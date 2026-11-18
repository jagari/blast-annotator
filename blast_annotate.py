#!/usr/bin/env python3
"""
기능 미상 서열 자동 BLAST + 엑셀 정리 프로그램
================================================

서열(FASTA 파일 또는 붙여넣기)을 넣으면
  1) NCBI BLAST 웹 서버에 자동으로 제출하고
  2) 결과를 기다렸다가 받아와서
  3) 가장 그럴듯한 기능(유전자/단백질 이름)을 뽑아
  4) 엑셀 파일(.xlsx)로 정리합니다.

필요한 것: Python 3.8+, openpyxl  (pip3 install openpyxl)

사용 예)
  python3 blast_annotate.py my_seqs.fasta
  python3 blast_annotate.py my_seqs.fasta -o 결과.xlsx --program blastn
  python3 blast_annotate.py            # 실행 후 서열을 직접 붙여넣기

  # 로컬 모드 (빠름): brew install blast 후
  python3 blast_annotate.py --setup-local          # 처음 한 번, DB 준비
  python3 blast_annotate.py my_seqs.fasta --local
"""

import argparse
import gzip
import shutil
import subprocess
import tempfile
import threading
import hashlib
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl 이 필요합니다. 터미널에서 다음을 실행하세요:  pip3 install openpyxl")

BLAST_URL = "https://blast.ncbi.nlm.nih.gov/Blast.cgi"
TOOL_NAME = "blast_annotate_py"

NUC_CHARS = set("ACGTUN")
PROT_ONLY_CHARS = set("EFILPQ*")  # 염기서열에는 안 나오는 아미노산 문자

UNINFORMATIVE = re.compile(
    r"hypothetical|uncharacteri[sz]ed|unknown|unnamed|predicted protein|"
    r"putative uncharacterized|low quality protein|domain of unknown function|DUF\d",
    re.I,
)


# ----------------------------------------------------------------------------
# 입력 처리
# ----------------------------------------------------------------------------
def parse_sequences(text):
    """FASTA 또는 '한 줄에 서열 하나' 형식을 모두 받아 [(이름, 서열)] 반환."""
    records = []
    if ">" in text:
        name, buf = None, []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith(">"):
                if name is not None:
                    records.append((name, "".join(buf)))
                name, buf = line[1:].strip() or f"seq{len(records) + 1}", []
            else:
                buf.append(line)
        if name is not None:
            records.append((name, "".join(buf)))
    else:
        for line in text.splitlines():
            line = line.strip()
            if line:
                records.append((f"seq{len(records) + 1}", line))

    cleaned = []
    for name, seq in records:
        seq = re.sub(r"[^A-Za-z*]", "", seq).upper()
        if seq:
            cleaned.append((name, seq))
    # 이름 중복 방지
    seen = Counter()
    out = []
    for name, seq in cleaned:
        seen[name] += 1
        out.append((name if seen[name] == 1 else f"{name}_{seen[name]}", seq))
    return out


def seq_type(seq):
    letters = set(seq)
    if letters & PROT_ONLY_CHARS:
        return "protein"
    acgt = sum(seq.count(c) for c in NUC_CHARS)
    return "nucleotide" if acgt / max(len(seq), 1) >= 0.9 else "protein"


def default_program(stype):
    # 기능을 알고 싶을 때는 DNA -> 단백질 번역 비교(blastx)가 기능 이름을 훨씬 잘 알려줌
    return "blastx" if stype == "nucleotide" else "blastp"


def default_db(program):
    return {"blastn": "core_nt", "megablast": "core_nt", "tblastx": "core_nt"}.get(program, "nr")


# ----------------------------------------------------------------------------
# NCBI BLAST URL API
# ----------------------------------------------------------------------------
def http(params, timeout=120):
    data = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(BLAST_URL, data=data, headers={"User-Agent": TOOL_NAME})
    for attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode("utf-8", errors="replace")
        except Exception as e:  # 네트워크 일시 오류 재시도
            wait = 15 * (attempt + 1)
            log(f"  ! 네트워크 오류({e}) - {wait}초 후 재시도")
            sleep_cancellable(wait)
    raise RuntimeError("NCBI 서버에 연결할 수 없습니다.")


def submit(fasta, program, db, hitlist, email, entrez_query=None):
    params = {
        "CMD": "Put",
        "PROGRAM": "blastn" if program == "megablast" else program,
        "DATABASE": db,
        "QUERY": fasta,
        "HITLIST_SIZE": str(hitlist),
        "EXPECT": "10",
        "FORMAT_TYPE": "XML",
        "TOOL": TOOL_NAME,
    }
    if program == "megablast":
        params["MEGABLAST"] = "on"
    if email:
        params["EMAIL"] = email
    if entrez_query:
        params["ENTREZ_QUERY"] = entrez_query
    page = http(params)
    rid = re.search(r"^\s*RID = (\S+)", page, re.M)
    rtoe = re.search(r"^\s*RTOE = (\d+)", page, re.M)
    if not rid:
        msg = re.search(r'<p class="error">(.*?)</p>', page, re.S)
        raise RuntimeError("BLAST 제출 실패: " + (msg.group(1).strip() if msg else page[:300]))
    return rid.group(1), int(rtoe.group(1)) if rtoe else 30


def check_status(rid, email):
    """READY / WAITING / FAILED / UNKNOWN"""
    page = http({"CMD": "Get", "FORMAT_OBJECT": "SearchInfo", "RID": rid,
                 "TOOL": TOOL_NAME, **({"EMAIL": email} if email else {})})
    m = re.search(r"Status=(\w+)", page)
    return m.group(1) if m else "UNKNOWN"


def fetch_result(rid, email):
    return http({"CMD": "Get", "RID": rid, "FORMAT_TYPE": "XML", "TOOL": TOOL_NAME,
                 **({"EMAIL": email} if email else {})}, timeout=300)


def run_web_jobs(jobs, email):
    """모든 묶음을 먼저 제출하고, 동시에 기다린다.
    NCBI 규칙: 제출은 10초에 1번 이하, 같은 RID 상태 확인은 1분에 1번 이하."""
    pending = [j for j in jobs if j["xml"] is None]
    last_submit = 0.0
    for j in pending:
        gap = time.time() - last_submit
        if gap < 11:
            sleep_cancellable(11 - gap)
        check_cancel()
        j["rid"], rtoe = submit(j["fasta"], j["program"], j["db"], j["hits"], email, j["organism"])
        last_submit = time.time()
        j["next_check"] = last_submit + min(max(rtoe, 15), 120)
        j["submitted"] = last_submit
        log(f"  제출 완료 [{j['label']}] RID {j['rid']} (NCBI 예상 {rtoe}초)")

    while pending:
        now = time.time()
        due = [j for j in pending if j["next_check"] <= now]
        if not due:
            sleep_cancellable(max(1, min(j["next_check"] for j in pending) - now))
            continue
        for j in due:
            status = check_status(j["rid"], email)
            elapsed = int(time.time() - j["submitted"])
            if status == "READY":
                j["xml"] = fetch_result(j["rid"], email)
                with open(j["cache_file"], "w", encoding="utf-8") as f:
                    f.write(j["xml"])
                pending.remove(j)
                log(f"  결과 도착 [{j['label']}] ({elapsed // 60}분 {elapsed % 60}초) - 남은 묶음 {len(pending)}개")
            elif status in ("FAILED", "UNKNOWN"):
                raise RuntimeError(f"BLAST 검색 실패 (RID {j['rid']}, 상태 {status}). 다시 실행하면 "
                                   "끝난 묶음은 캐시에서 불러옵니다.")
            else:
                j["next_check"] = time.time() + 60
        if pending and due:
            waits = ", ".join(f"{j['label']} {int(time.time() - j['submitted']) // 60}분"
                              for j in pending)
            log(f"  검색 중... ({waits})")


# ----------------------------------------------------------------------------
# 결과 해석
# ----------------------------------------------------------------------------
def txt(el, tag, default=""):
    x = el.find(tag)
    return x.text if x is not None and x.text is not None else default


def parse_blast_xml(xml_text, query_lengths):
    """BLAST XML -> {쿼리순번: [hit dict, ...]}"""
    root = ET.fromstring(xml_text)
    results = {}
    for it_idx, it in enumerate(root.iter("Iteration")):
        qdef = txt(it, "Iteration_query-def")
        qlen = int(txt(it, "Iteration_query-len", "0")) or 1
        hits = []
        for hit in it.iter("Hit"):
            hsps = list(hit.iter("Hsp"))
            if not hsps:
                continue
            best = min(hsps, key=lambda h: float(txt(h, "Hsp_evalue", "1e9")))
            ident = int(txt(best, "Hsp_identity", "0"))
            alen = int(txt(best, "Hsp_align-len", "1")) or 1
            # 쿼리 커버리지: 모든 HSP의 쿼리 구간 합집합
            spans = []
            for h in hsps:
                a, b = int(txt(h, "Hsp_query-from", "0")), int(txt(h, "Hsp_query-to", "0"))
                spans.append((min(a, b), max(a, b)))
            spans.sort()
            covered, cur_s, cur_e = 0, None, None
            for s, e in spans:
                if cur_e is None or s > cur_e + 1:
                    if cur_e is not None:
                        covered += cur_e - cur_s + 1
                    cur_s, cur_e = s, e
                else:
                    cur_e = max(cur_e, e)
            if cur_e is not None:
                covered += cur_e - cur_s + 1

            desc = txt(hit, "Hit_def")
            # nr 결과는 여러 설명이 '>'로 이어져 있음 -> 첫 번째만
            desc = desc.split(" >")[0].strip()
            acc = txt(hit, "Hit_accession")
            hits.append({
                "description": desc,
                "accession": acc,
                "hit_len": int(txt(hit, "Hit_len", "0")),
                "identity": round(100.0 * ident / alen, 1),
                "coverage": round(min(100.0, 100.0 * covered / qlen), 1),
                "evalue": float(txt(best, "Hsp_evalue", "1e9")),
                "bitscore": float(txt(best, "Hsp_bit-score", "0")),
                "align_len": alen,
            })
        results[it_idx] = {"query_def": qdef, "hits": hits}
    return results


def split_desc(desc):
    """'beta-lactamase TEM [Escherichia coli]' -> ('beta-lactamase TEM', 'Escherichia coli')"""
    if " OS=" in desc:  # UniProt Swiss-Prot (로컬 DB) 형식
        func = desc.split(" OS=")[0].strip()
        org = re.search(r" OS=(.*?)(?: OX=| GN=| PE=| SV=|$)", desc)
        return func, org.group(1).strip() if org else ""
    m = re.match(r"^(.*?)\s*\[([^\[\]]+)\]\s*$", desc)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # 염기 DB 제목: 'Escherichia coli strain X blaTEM-1 gene, complete cds'
    words = desc.split()
    organism = " ".join(words[:2]) if len(words) >= 2 and words[0][:1].isupper() else ""
    return desc, organism


def gene_name(desc):
    m = re.search(r" GN=(\S+)", desc)
    return m.group(1) if m else ""


def clean_function(name):
    name = re.sub(r"^(MULTISPECIES|RecName|PREDICTED|LOW QUALITY PROTEIN):\s*", "", name, flags=re.I)
    name = re.sub(r"^Full=", "", name)
    name = re.sub(r";.*$", "", name)  # RecName: Full=X; AltName=... -> X
    return name.strip()


def core_function(name):
    """변이/아형 표기를 떼어낸 핵심 기능명: 'Beta-lactamase TEM-12' -> 'beta-lactamase'"""
    words = re.sub(r"\s*\(.*?\)", "", name).split()
    while len(words) > 1 and re.fullmatch(r"(?=.*[A-Z0-9])[A-Za-z0-9-]{1,7}", words[-1]) \
            and not words[-1].islower():
        words.pop()
    return " ".join(words).lower()


def confidence(hit, program):
    if hit is None:
        return "결과 없음"
    ide, cov, ev = hit["identity"], hit["coverage"], hit["evalue"]
    if program in ("blastn", "megablast"):
        if ide >= 95 and cov >= 80:
            return "높음"
        if ide >= 80 and cov >= 50 and ev < 1e-20:
            return "보통"
    else:
        if ide >= 60 and cov >= 70 and ev < 1e-30:
            return "높음"
        if ide >= 30 and cov >= 40 and ev < 1e-5:
            return "보통"
    return "낮음"


def summarize(name, seq, stype, program, res):
    hits = res["hits"] if res else []
    top = hits[0] if hits else None
    informative = next((h for h in hits if not UNINFORMATIVE.search(h["description"])), None)
    chosen = informative or top

    function, organism = ("", "")
    if chosen:
        function, organism = split_desc(chosen["description"])
        function = clean_function(function)

    # 상위 히트들에서 가장 많이 나온 기능 이름 (합의)
    names = [core_function(clean_function(split_desc(h["description"])[0]))
             for h in hits[:10] if not UNINFORMATIVE.search(h["description"])]
    consensus = ""
    if names:
        cname, cnt = Counter(names).most_common(1)[0]
        consensus = f"{cname} ({cnt}/{min(len(hits), 10)})"

    note = ""
    if top and chosen is not top:
        note = "최상위 히트가 기능 미상(hypothetical)이라 다음 정보성 히트를 사용"
    if not hits:
        note = "유의한 일치 서열 없음 - 새로운 서열이거나 서열 품질 확인 필요"

    return {
        "name": name, "length": len(seq), "type": "염기" if stype == "nucleotide" else "단백질",
        "program": program, "function": function or "(일치 없음)", "consensus": consensus,
        "organism": organism, "gene": gene_name(chosen["description"]) if chosen else "",
        "accession": chosen["accession"] if chosen else "",
        "identity": chosen["identity"] if chosen else None,
        "coverage": chosen["coverage"] if chosen else None,
        "evalue": chosen["evalue"] if chosen else None,
        "bitscore": chosen["bitscore"] if chosen else None,
        "confidence": confidence(chosen, program), "n_hits": len(hits),
        "top_desc": top["description"] if top else "", "note": note, "hits": hits,
    }


# ----------------------------------------------------------------------------
# 엑셀 작성
# ----------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
CONF_FILL = {"높음": "C6EFCE", "보통": "FFEB9C", "낮음": "FFC7CE", "결과 없음": "D9D9D9"}


def ncbi_link(acc, program):
    if not acc:
        return ""
    if re.fullmatch(r"[OPQ][0-9][A-Z0-9]{3}[0-9]|[A-NR-Z][0-9](?:[A-Z][A-Z0-9]{2}[0-9]){1,2}", acc):
        return f"https://www.uniprot.org/uniprotkb/{acc}"
    db = "nuccore" if program in ("blastn", "megablast", "tblastx") else "protein"
    return f"https://www.ncbi.nlm.nih.gov/{db}/{acc}"


def write_header(ws, headers, widths):
    ws.append(headers)
    for i, w in enumerate(widths, 1):
        c = ws.cell(row=1, column=i)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 30


def fmt_e(e):
    return None if e is None else (0 if e == 0 else float(f"{e:.2e}"))


def write_excel(path, summaries, records, settings):
    wb = Workbook()

    # 1) 요약
    ws = wb.active
    ws.title = "기능 요약"
    headers = ["서열 이름", "길이", "종류", "BLAST 종류", "추정 기능", "신뢰도",
               "상위 10개 합의 기능", "유래 생물", "유전자명", "Accession", "Identity(%)",
               "Query 커버리지(%)", "E-value", "Bit score", "히트 수", "최상위 히트 원문",
               "비고", "NCBI 링크"]
    write_header(ws, headers, [22, 8, 8, 10, 40, 9, 32, 26, 10, 15, 11, 12, 11, 10, 8, 50, 36, 40])
    for s in summaries:
        link = ncbi_link(s["accession"], s["program"])
        ws.append([s["name"], s["length"], s["type"], s["program"], s["function"], s["confidence"],
                   s["consensus"], s["organism"], s["gene"], s["accession"], s["identity"], s["coverage"],
                   fmt_e(s["evalue"]), s["bitscore"], s["n_hits"], s["top_desc"], s["note"], link])
        r = ws.max_row
        ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor=CONF_FILL[s["confidence"]])
        ws.cell(row=r, column=13).number_format = "0.00E+00"
        if link:
            c = ws.cell(row=r, column=18)
            c.hyperlink, c.font = link, Font(color="0563C1", underline="single")
        for col in (5, 7, 16, 17):
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.auto_filter.ref = ws.dimensions

    # 2) 상세 히트
    wd = wb.create_sheet("상세 히트")
    write_header(wd, ["서열 이름", "순위", "기능(설명)", "유래 생물", "유전자명", "Accession", "Identity(%)",
                      "Query 커버리지(%)", "정렬 길이", "E-value", "Bit score", "NCBI 링크"],
                 [22, 6, 50, 28, 10, 15, 11, 12, 9, 11, 10, 40])
    for s in summaries:
        for i, h in enumerate(s["hits"], 1):
            func, org = split_desc(h["description"])
            link = ncbi_link(h["accession"], s["program"])
            wd.append([s["name"], i, clean_function(func), org, gene_name(h["description"]),
                       h["accession"], h["identity"],
                       h["coverage"], h["align_len"], fmt_e(h["evalue"]), h["bitscore"], link])
            r = wd.max_row
            wd.cell(row=r, column=10).number_format = "0.00E+00"
            c = wd.cell(row=r, column=12)
            c.hyperlink, c.font = link, Font(color="0563C1", underline="single")
    wd.auto_filter.ref = wd.dimensions

    # 3) 입력 서열
    wi = wb.create_sheet("입력 서열")
    write_header(wi, ["서열 이름", "길이", "서열"], [22, 8, 100])
    for name, seq in records:
        wi.append([name, len(seq), seq])
        wi.cell(row=wi.max_row, column=3).alignment = Alignment(wrap_text=True)

    # 4) 실행 정보
    wsx = wb.create_sheet("실행 정보")
    for k, v in settings.items():
        wsx.append([k, v])
    wsx.append([])
    wsx.append(["신뢰도 기준", "blastx/blastp: 높음 = identity≥60%, 커버리지≥70%, E<1e-30 / "
                "보통 = identity≥30%, 커버리지≥40%, E<1e-5"])
    wsx.append(["", "blastn: 높음 = identity≥95%, 커버리지≥80% / 보통 = identity≥80%, 커버리지≥50%, E<1e-20"])
    wsx.column_dimensions["A"].width = 18
    wsx.column_dimensions["B"].width = 100

    wb.save(path)


# ----------------------------------------------------------------------------
# 메인
# ----------------------------------------------------------------------------
LOG_HOOK = None                 # GUI 가 지정하면 로그가 그쪽으로 감
CANCEL = threading.Event()      # GUI '중지' 버튼


class Cancelled(Exception):
    pass


def log(msg):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    if LOG_HOOK:
        LOG_HOOK(line)
    else:
        print(line, flush=True)


def check_cancel():
    if CANCEL.is_set():
        raise Cancelled()


def sleep_cancellable(sec):
    if CANCEL.wait(max(0, sec)):
        raise Cancelled()


def run_cmd(cmd, **kw):
    """Windows 에서 검은 콘솔창이 뜨지 않게 subprocess 실행."""
    if os.name == "nt":
        kw["creationflags"] = kw.get("creationflags", 0) | 0x08000000  # CREATE_NO_WINDOW
    return subprocess.run(cmd, **kw)


def read_input(paths):
    if paths:
        text = ""
        for p in paths:
            with open(p, encoding="utf-8", errors="replace") as f:
                text += f.read() + "\n"
        return text
    print("서열을 붙여넣으세요 (FASTA 또는 한 줄에 하나).")
    print("다 넣었으면 빈 줄에서 Ctrl+D (Windows는 Ctrl+Z 후 Enter) 를 누르세요.\n")
    return sys.stdin.read()


# ----------------------------------------------------------------------------
# 로컬 BLAST+ 모드
# ----------------------------------------------------------------------------
SPROT_URL = ("https://ftp.uniprot.org/pub/databases/uniprot/current_release/"
             "knowledgebase/complete/uniprot_sprot.fasta.gz")
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def user_data_dir():
    """캐시 저장 위치 (설치 폴더가 쓰기 금지여도 동작하도록 사용자 폴더 사용)."""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~/Library/Caches")
    d = os.path.join(base, "BlastAnnotator")
    os.makedirs(d, exist_ok=True)
    return d
DEFAULT_LOCAL_DB = os.path.join(SCRIPT_DIR, "blastdb", "swissprot")


def find_blast_bin(name):
    import glob
    if os.name == "nt" and not name.endswith(".exe"):
        name += ".exe"
    # 1순위: 프로그램 폴더의 blast/bin 또는 tools/ncbi-blast-*/bin
    cands = [os.path.join(SCRIPT_DIR, "blast", "bin")]
    cands += sorted(glob.glob(os.path.join(SCRIPT_DIR, "tools", "ncbi-blast-*", "bin")), reverse=True)
    for d in cands:
        if os.path.exists(os.path.join(d, name)):
            return os.path.join(d, name)
    path = shutil.which(name)
    if not path:
        for d in ("/opt/homebrew/bin", "/usr/local/bin", os.path.expanduser("~/miniconda3/bin")):
            if os.path.exists(os.path.join(d, name)):
                return os.path.join(d, name)
        raise RuntimeError(f"'{name}' 이(가) 없습니다. BLAST+ 를 먼저 설치하세요 (사용법.md 참고)")
    return path


def local_db_exists(db):
    return any(os.path.exists(db + ext) for ext in (".pin", ".pal", ".nin", ".nal", ".psq", ".nsq"))


def setup_local_db():
    """UniProt Swiss-Prot 을 내려받아 로컬 BLAST DB 로 만든다 (약 90MB 다운로드, 1~3분)."""
    makeblastdb = find_blast_bin("makeblastdb")
    os.makedirs(os.path.dirname(DEFAULT_LOCAL_DB), exist_ok=True)
    gz_path = DEFAULT_LOCAL_DB + ".fasta.gz"
    fa_path = DEFAULT_LOCAL_DB + ".fasta"
    if os.path.exists(fa_path) and os.path.getsize(fa_path) > 100 << 20:
        log("이미 받은 Swiss-Prot 파일 사용")
        return _build_db(makeblastdb, gz_path, fa_path)
    log(f"Swiss-Prot 다운로드 중: {SPROT_URL}")
    req = urllib.request.Request(SPROT_URL, headers={"User-Agent": TOOL_NAME})
    with urllib.request.urlopen(req, timeout=600) as r, open(gz_path, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        done, last = 0, 0
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if done - last >= 10 << 20:
                last = done
                log(f"  {done >> 20}MB" + (f" / {total >> 20}MB" if total else ""))
    log("압축 해제 중...")
    with gzip.open(gz_path, "rb") as src, open(fa_path, "wb") as dst:
        shutil.copyfileobj(src, dst)
    return _build_db(makeblastdb, gz_path, fa_path)


def _build_db(makeblastdb, gz_path, fa_path):
    log("BLAST DB 만드는 중 (makeblastdb)...")
    # BLAST+ 는 경로에 공백이 있으면 실패하므로 DB 폴더 안에서 상대 경로로 실행
    dbdir, dbname = os.path.split(DEFAULT_LOCAL_DB)
    run_cmd([makeblastdb, "-in", os.path.basename(fa_path), "-dbtype", "prot",
                    "-parse_seqids", "-out", dbname, "-title", "UniProt Swiss-Prot"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, cwd=dbdir)
    # makeblastdb 의 마지막 자체 점검이 공백 경로에서 실패하는 경우가 있어, 직접 확인
    chk = run_cmd([find_blast_bin("blastdbcmd"), "-db", dbname, "-info"],
                         capture_output=True, text=True, cwd=dbdir)
    if chk.returncode != 0:
        raise RuntimeError("로컬 DB 생성 실패: " + chk.stderr.strip()[:300])
    n = re.search(r"([\d,]+) sequences", chk.stdout)
    log(f"  단백질 {n.group(1) if n else '?'}개 등록")
    for f in (gz_path, fa_path):
        if os.path.exists(f):
            os.remove(f)
    log(f"로컬 DB 준비 완료: {DEFAULT_LOCAL_DB}")


def run_local_job(job, threads):
    check_cancel()
    exe = find_blast_bin("blastn" if job["program"] == "megablast" else job["program"])
    dbdir, dbname = os.path.split(job["db"])  # 공백/한글 경로 회피: DB 폴더에서 상대 이름으로 실행
    # 쿼리는 표준입력(-query 기본값)으로 전달 -> 임시파일 경로 문제 없음
    cmd = [exe, "-db", dbname, "-outfmt", "5", "-evalue", "10",
           "-max_target_seqs", str(job["hits"]), "-num_threads", str(threads)]
    if job["program"] == "megablast":
        cmd += ["-task", "megablast"]
    elif job["program"] == "blastn":
        cmd += ["-task", "blastn"]
    t0 = time.time()
    r = run_cmd(cmd, input=job["fasta"] + "\n", capture_output=True, text=True,
                encoding="utf-8", errors="replace", cwd=dbdir or ".")
    if r.returncode != 0:
        raise RuntimeError(f"로컬 {job['program']} 실행 실패: {r.stderr.strip()[:500]}")
    log(f"  완료 [{job['label']}] ({time.time() - t0:.1f}초)")
    return r.stdout


# ----------------------------------------------------------------------------
# 메인
# ----------------------------------------------------------------------------
LOG_HOOK = None                 # GUI 가 지정하면 로그가 그쪽으로 감
CANCEL = threading.Event()      # GUI '중지' 버튼


class Cancelled(Exception):
    pass


def log(msg):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    if LOG_HOOK:
        LOG_HOOK(line)
    else:
        print(line, flush=True)


def check_cancel():
    if CANCEL.is_set():
        raise Cancelled()


def sleep_cancellable(sec):
    if CANCEL.wait(max(0, sec)):
        raise Cancelled()


def run_cmd(cmd, **kw):
    """Windows 에서 검은 콘솔창이 뜨지 않게 subprocess 실행."""
    if os.name == "nt":
        kw["creationflags"] = kw.get("creationflags", 0) | 0x08000000  # CREATE_NO_WINDOW
    return subprocess.run(cmd, **kw)


def read_input(paths):
    if paths:
        text = ""
        for p in paths:
            with open(p, encoding="utf-8", errors="replace") as f:
                text += f.read() + "\n"
        return text
    print("서열을 붙여넣으세요 (FASTA 또는 한 줄에 하나).")
    print("다 넣었으면 빈 줄에서 Ctrl+D (Windows는 Ctrl+Z 후 Enter) 를 누르세요.\n")
    return sys.stdin.read()


# ----------------------------------------------------------------------------
# 로컬 BLAST+ 모드
# ----------------------------------------------------------------------------
SPROT_URL = ("https://ftp.uniprot.org/pub/databases/uniprot/current_release/"
             "knowledgebase/complete/uniprot_sprot.fasta.gz")
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def user_data_dir():
    """캐시 저장 위치 (설치 폴더가 쓰기 금지여도 동작하도록 사용자 폴더 사용)."""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~/Library/Caches")
    d = os.path.join(base, "BlastAnnotator")
    os.makedirs(d, exist_ok=True)
    return d
DEFAULT_LOCAL_DB = os.path.join(SCRIPT_DIR, "blastdb", "swissprot")


def find_blast_bin(name):
    import glob
    if os.name == "nt" and not name.endswith(".exe"):
        name += ".exe"
    # 1순위: 프로그램 폴더의 blast/bin 또는 tools/ncbi-blast-*/bin
    cands = [os.path.join(SCRIPT_DIR, "blast", "bin")]
    cands += sorted(glob.glob(os.path.join(SCRIPT_DIR, "tools", "ncbi-blast-*", "bin")), reverse=True)
    for d in cands:
        if os.path.exists(os.path.join(d, name)):
            return os.path.join(d, name)
    path = shutil.which(name)
    if not path:
        for d in ("/opt/homebrew/bin", "/usr/local/bin", os.path.expanduser("~/miniconda3/bin")):
            if os.path.exists(os.path.join(d, name)):
                return os.path.join(d, name)
        raise RuntimeError(f"'{name}' 이(가) 없습니다. BLAST+ 를 먼저 설치하세요 (사용법.md 참고)")
    return path


def local_db_exists(db):
    return any(os.path.exists(db + ext) for ext in (".pin", ".pal", ".nin", ".nal", ".psq", ".nsq"))


def setup_local_db():
    """UniProt Swiss-Prot 을 내려받아 로컬 BLAST DB 로 만든다 (약 90MB 다운로드, 1~3분)."""
    makeblastdb = find_blast_bin("makeblastdb")
    os.makedirs(os.path.dirname(DEFAULT_LOCAL_DB), exist_ok=True)
    gz_path = DEFAULT_LOCAL_DB + ".fasta.gz"
    fa_path = DEFAULT_LOCAL_DB + ".fasta"
    if os.path.exists(fa_path) and os.path.getsize(fa_path) > 100 << 20:
        log("이미 받은 Swiss-Prot 파일 사용")
        return _build_db(makeblastdb, gz_path, fa_path)
    log(f"Swiss-Prot 다운로드 중: {SPROT_URL}")
    req = urllib.request.Request(SPROT_URL, headers={"User-Agent": TOOL_NAME})
    with urllib.request.urlopen(req, timeout=600) as r, open(gz_path, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        done, last = 0, 0
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if done - last >= 10 << 20:
                last = done
                log(f"  {done >> 20}MB" + (f" / {total >> 20}MB" if total else ""))
    log("압축 해제 중...")
    with gzip.open(gz_path, "rb") as src, open(fa_path, "wb") as dst:
        shutil.copyfileobj(src, dst)
    return _build_db(makeblastdb, gz_path, fa_path)


def _build_db(makeblastdb, gz_path, fa_path):
    log("BLAST DB 만드는 중 (makeblastdb)...")
    # BLAST+ 는 경로에 공백이 있으면 실패하므로 DB 폴더 안에서 상대 경로로 실행
    dbdir, dbname = os.path.split(DEFAULT_LOCAL_DB)
    run_cmd([makeblastdb, "-in", os.path.basename(fa_path), "-dbtype", "prot",
                    "-parse_seqids", "-out", dbname, "-title", "UniProt Swiss-Prot"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, cwd=dbdir)
    # makeblastdb 의 마지막 자체 점검이 공백 경로에서 실패하는 경우가 있어, 직접 확인
    chk = run_cmd([find_blast_bin("blastdbcmd"), "-db", dbname, "-info"],
                         capture_output=True, text=True, cwd=dbdir)
    if chk.returncode != 0:
        raise RuntimeError("로컬 DB 생성 실패: " + chk.stderr.strip()[:300])
    n = re.search(r"([\d,]+) sequences", chk.stdout)
    log(f"  단백질 {n.group(1) if n else '?'}개 등록")
    for f in (gz_path, fa_path):
        if os.path.exists(f):
            os.remove(f)
    log(f"로컬 DB 준비 완료: {DEFAULT_LOCAL_DB}")


def run_local_job(job, threads):
    exe = find_blast_bin("blastn" if job["program"] == "megablast" else job["program"])
    with tempfile.NamedTemporaryFile("w", suffix=".fasta", delete=False) as tf:
        tf.write(job["fasta"] + "\n")
        qpath = tf.name
    dbdir, dbname = os.path.split(job["db"])  # 공백 경로 회피: DB 폴더에서 실행
    cmd = [exe, "-query", qpath, "-db", dbname, "-outfmt", "5", "-evalue", "10",
           "-max_target_seqs", str(job["hits"]), "-num_threads", str(threads)]
    if job["program"] == "megablast":
        cmd += ["-task", "megablast"]
    elif job["program"] == "blastn":
        cmd += ["-task", "blastn"]
    try:
        t0 = time.time()
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=dbdir or ".")
    finally:
        os.remove(qpath)
    if r.returncode != 0:
        raise RuntimeError(f"로컬 {job['program']} 실행 실패: {r.stderr.strip()[:500]}")
    log(f"  완료 [{job['label']}] ({time.time() - t0:.1f}초)")
    return r.stdout


# ----------------------------------------------------------------------------
# 메인
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="서열을 BLAST로 자동 검색하고 기능을 엑셀로 정리합니다.")
    ap.add_argument("inputs", nargs="*", help="FASTA/텍스트 파일 (생략하면 직접 붙여넣기)")
    ap.add_argument("-o", "--output", help="엑셀 파일 이름 (기본: blast_결과_날짜시간.xlsx)")
    ap.add_argument("-p", "--program", choices=["auto", "blastx", "blastn", "megablast", "blastp", "tblastx"],
                    default="auto", help="auto(기본): DNA->blastx, 단백질->blastp")
    ap.add_argument("-d", "--db", help="데이터베이스. 웹: 기본 nr (blastn은 core_nt), swissprot 추천. "
                                       "로컬: DB 경로 (기본: blastdb/swissprot)")
    ap.add_argument("-n", "--hits", type=int, default=10, help="서열당 저장할 히트 수 (기본 10)")
    ap.add_argument("-b", "--batch", type=int, default=10, help="한 번에 제출할 서열 수 (기본 10)")
    ap.add_argument("--local", action="store_true",
                    help="NCBI 서버 대신 내 컴퓨터에서 검색 (빠름, 대기열 없음)")
    ap.add_argument("--setup-local", action="store_true",
                    help="로컬 검색용 Swiss-Prot DB 를 내려받아 준비 (처음 한 번)")
    ap.add_argument("--threads", type=int, default=os.cpu_count() or 4, help="로컬 검색 CPU 수")
    ap.add_argument("--organism", help="(웹 전용) 특정 생물로 제한 (예: 'Escherichia coli[organism]')")
    ap.add_argument("--email", default=os.environ.get("NCBI_EMAIL", ""),
                    help="(웹 전용) NCBI에 알릴 연락 이메일 (권장, 또는 환경변수 NCBI_EMAIL)")
    ap.add_argument("--cache", default=None, help="결과 캐시 폴더 (기본: 사용자 폴더, 중단 후 재실행 시 재사용)")
    args = ap.parse_args()

    if args.setup_local:
        setup_local_db()
        if not args.inputs:
            return

    if args.cache is None:
        args.cache = os.path.join(user_data_dir(), "cache")
    records = parse_sequences(read_input(args.inputs))
    run_analysis(records, args)


class Options:
    """GUI 에서 쓰는 설정 묶음 (argparse 결과와 같은 이름)."""
    def __init__(self, **kw):
        self.local, self.program, self.db, self.hits, self.batch = True, "auto", None, 10, 10
        self.threads = os.cpu_count() or 4
        self.organism, self.email, self.output = None, "", None
        self.cache = os.path.join(user_data_dir(), "cache")
        self.__dict__.update(kw)


def run_analysis(records, args):
    """서열 목록을 검색하고 엑셀을 만든다. (엑셀 경로, 요약 목록) 반환."""
    CANCEL.clear()
    if args.local:
        db_path = os.path.abspath(args.db) if args.db else DEFAULT_LOCAL_DB
        if not local_db_exists(db_path):
            raise RuntimeError(f"로컬 DB 가 없습니다 ({db_path}).\n"
                               "  먼저 실행:  python3 blast_annotate.py --setup-local")

    if not records:
        raise RuntimeError("서열을 찾지 못했습니다.")
    log(f"서열 {len(records)}개 읽음 ({'로컬' if args.local else 'NCBI 웹'} 모드)")

    # 프로그램/DB 별로 묶기
    groups = {}
    for idx, (name, seq) in enumerate(records):
        st = seq_type(seq)
        prog = default_program(st) if args.program == "auto" else args.program
        if st == "protein" and prog != "blastp":
            log(f"  ! '{name}' 은 단백질 서열이라 blastp 로 검색합니다.")
            prog = "blastp"
        if st == "nucleotide" and prog == "blastp":
            prog = "blastx"
        db = db_path if args.local else (args.db or default_db(prog))
        groups.setdefault((prog, db), []).append(idx)

    # 묶음(job) 만들기
    os.makedirs(args.cache, exist_ok=True)
    jobs = []
    for (prog, db), idxs in groups.items():
        for b in range(0, len(idxs), args.batch):
            chunk = idxs[b:b + args.batch]
            # 쿼리 이름은 내부 번호로 (이름에 특수문자 있어도 안전)
            fasta = "\n".join(f">q{i}\n{records[i][1]}" for i in chunk)
            mode = "local" if args.local else "web"
            key = hashlib.md5(f"{mode}|{prog}|{db}|{args.hits}|{args.organism}|{fasta}".encode()
                              ).hexdigest()[:16]
            if not args.local:  # 이전 버전과 캐시 호환
                key = hashlib.md5(f"{prog}|{db}|{args.hits}|{args.organism}|{fasta}".encode()
                                  ).hexdigest()[:16]
            job = {"program": prog, "db": db, "hits": args.hits, "organism": args.organism,
                   "fasta": fasta, "chunk": chunk, "xml": None,
                   "label": f"{prog} #{len(jobs) + 1}",
                   "cache_file": os.path.join(args.cache, f"{key}.xml")}
            if os.path.exists(job["cache_file"]):
                with open(job["cache_file"], encoding="utf-8") as f:
                    job["xml"] = f.read()
            jobs.append(job)

    todo = [j for j in jobs if j["xml"] is None]
    log(f"검색 묶음 {len(jobs)}개 (캐시 사용 {len(jobs) - len(todo)}개, 새로 검색 {len(todo)}개)")
    if args.local:
        for n, j in enumerate(todo, 1):
            log(f"  로컬 검색 {n}/{len(todo)}")
            j["xml"] = run_local_job(j, args.threads)
            with open(j["cache_file"], "w", encoding="utf-8") as f:
                f.write(j["xml"])
    else:
        run_web_jobs(jobs, args.email)
    check_cancel()

    results, prog_of = {}, {}
    for j in jobs:
        try:
            parsed = parse_blast_xml(j["xml"], None)
        except ET.ParseError:
            os.remove(j["cache_file"])
            raise RuntimeError("BLAST 결과를 해석하지 못했습니다. 다시 실행해 주세요.")
        by_qdef = {v["query_def"]: v for v in parsed.values()}
        for pos, i in enumerate(j["chunk"]):
            results[i] = by_qdef.get(f"q{i}") or parsed.get(pos)
            prog_of[i] = j["program"]

    summaries = []
    for i, (name, seq) in enumerate(records):
        s = summarize(name, seq, seq_type(seq), prog_of[i], results.get(i))
        summaries.append(s)
        log(f"  {name}: {s['function']} [{s['organism']}] - 신뢰도 {s['confidence']}")

    out = args.output or f"blast_결과_{datetime.now():%Y%m%d_%H%M}.xlsx"
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"
    settings = {
        "실행 시각": f"{datetime.now():%Y-%m-%d %H:%M}",
        "검색 방식": "로컬 BLAST+ (UniProt Swiss-Prot)" if args.local else "NCBI 웹 BLAST",
        "입력 서열 수": len(records),
        "검색 설정": "; ".join(f"{p} vs {os.path.basename(d)} ({len(v)}개)"
                           for (p, d), v in groups.items()),
        "서열당 히트 수": args.hits,
        "생물 제한": args.organism or "없음",
    }
    write_excel(out, summaries, records, settings)
    log(f"완료! 엑셀 파일: {os.path.abspath(out)}")
    return os.path.abspath(out), summaries


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, Cancelled):
        sys.exit("\n중단됨. 다시 실행하면 이미 끝난 검색은 캐시에서 바로 불러옵니다.")
    except (RuntimeError, subprocess.CalledProcessError) as e:
        sys.exit(f"오류: {e}")
