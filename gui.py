#!/usr/bin/env python3
"""서열 기능 자동 분석기 - 간단한 화면(GUI)

서열을 붙여넣거나 FASTA 파일을 열고 [분석 시작]을 누르면
BLAST 검색 후 기능을 정리한 엑셀 파일을 만들어 줍니다.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import blast_annotate as engine

APP_TITLE = "서열 기능 자동 분석기"
HINT = "#8a8a8a"     # 밝은/어두운 화면 모두에서 읽히는 회색
ACCENT = "#2F80ED"   # 밝은/어두운 화면 모두에서 읽히는 파랑
CONF_COLORS = {"높음": "#C6EFCE", "보통": "#FFEB9C", "낮음": "#FFC7CE", "결과 없음": "#E0E0E0"}


def default_output_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    base = docs if os.path.isdir(docs) else os.path.expanduser("~")
    d = os.path.join(base, "BLAST분석결과")
    return d


def open_path(path):
    try:
        if os.name == "nt":
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"열 수 없습니다:\n{path}\n\n{e}")


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.worker = None
        self.result_path = None
        self.out_dir = default_output_dir()

        root.title(APP_TITLE)
        root.geometry("940x820")
        root.minsize(760, 620)
        self._setup_fonts()
        self._build()
        self._poll_queue()
        self._update_count()

    # ------------------------------------------------------------------ 화면
    def _setup_fonts(self):
        family = "Malgun Gothic" if os.name == "nt" else ("Apple SD Gothic Neo" if sys.platform == "darwin" else None)
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            f = tkfont.nametofont(name)
            if family:
                f.configure(family=family)
            f.configure(size=10 if os.name == "nt" else 13)
        self.big = tkfont.Font(family=family, size=16 if os.name == "nt" else 20, weight="bold")
        self.bold = tkfont.Font(family=family, size=10 if os.name == "nt" else 13, weight="bold")
        mono = "Consolas" if os.name == "nt" else "Menlo"
        self.mono = tkfont.Font(family=mono, size=9 if os.name == "nt" else 12)
        style = ttk.Style()
        style.configure("Start.TButton", font=self.bold, padding=(18, 8))
        style.configure("TLabelframe.Label", font=self.bold)
        style.configure("Treeview", rowheight=24 if os.name == "nt" else 26)

    def _build(self):
        pad = {"padx": 14}
        top = ttk.Frame(self.root)
        top.pack(fill="x", pady=(12, 4), **pad)
        ttk.Label(top, text=APP_TITLE, font=self.big).pack(side="left")
        ttk.Label(top, text="서열을 넣고 [분석 시작]을 누르면 기능을 찾아 엑셀로 정리합니다",
                  foreground=HINT).pack(side="left", padx=12, pady=(6, 0))

        # 1. 서열 입력
        box1 = ttk.LabelFrame(self.root, text=" ① 서열 입력 ")
        box1.pack(fill="both", expand=False, pady=6, **pad)
        bar = ttk.Frame(box1)
        bar.pack(fill="x", padx=8, pady=(6, 2))
        ttk.Button(bar, text="📂 서열 파일 열기 (FASTA/TXT)", command=self.open_file).pack(side="left")
        ttk.Button(bar, text="지우기", command=self.clear_input).pack(side="left", padx=6)
        ttk.Button(bar, text="예시 넣기", command=self.insert_example).pack(side="left")
        self.count_lbl = ttk.Label(bar, text="", font=self.bold)
        self.count_lbl.pack(side="right")
        ttk.Label(box1, text="아래 칸에 서열을 붙여넣으세요. 여러 개는 FASTA 형식(>이름 다음 줄에 서열) "
                             "또는 한 줄에 하나씩.", foreground=HINT).pack(anchor="w", padx=8)
        tf = ttk.Frame(box1)
        tf.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        self.input = tk.Text(tf, height=6, wrap="char", font=self.mono, undo=True)
        sb = ttk.Scrollbar(tf, command=self.input.yview)
        self.input.configure(yscrollcommand=sb.set)
        self.input.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.input.bind("<<Modified>>", self._on_modified)

        # 2. 검색 방식
        box2 = ttk.LabelFrame(self.root, text=" ② 검색 방식 ")
        box2.pack(fill="x", pady=6, **pad)
        self.mode = tk.StringVar(value="local")
        ttk.Radiobutton(box2, variable=self.mode, value="local",
                        text="빠른 검색 (추천) - 내 컴퓨터에서 바로 검색, 서열당 1~2초, 인터넷 불필요"
                        ).pack(anchor="w", padx=8, pady=(6, 2))
        ttk.Radiobutton(box2, variable=self.mode, value="web",
                        text="정밀 검색 - NCBI 서버의 전체 DB에서 검색, 서열당 수 분~수십 분, 인터넷 필요"
                        ).pack(anchor="w", padx=8, pady=(0, 2))
        ttk.Label(box2, text="※ 빠른 검색에서 신뢰도 '낮음'/'결과 없음'이 나온 서열만 정밀 검색으로 다시 돌리는 것을 추천합니다.",
                  foreground=HINT).pack(anchor="w", padx=8, pady=(0, 6))

        # 3. 실행
        box3 = ttk.Frame(self.root)
        box3.pack(fill="x", pady=6, **pad)
        self.start_btn = ttk.Button(box3, text="▶  분석 시작", style="Start.TButton", command=self.start)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(box3, text="■ 중지", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(box3, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=8)
        self.status = ttk.Label(box3, text="대기 중", width=18)
        self.status.pack(side="left")

        # 결과
        box4 = ttk.LabelFrame(self.root, text=" ③ 결과 ")
        box4.pack(fill="both", expand=True, pady=(6, 12), **pad)
        rb = ttk.Frame(box4)
        rb.pack(fill="x", padx=8, pady=(6, 4))
        self.open_xlsx_btn = ttk.Button(rb, text="📊 엑셀 파일 열기", command=self.open_result, state="disabled")
        self.open_xlsx_btn.pack(side="left")
        ttk.Button(rb, text="📁 결과 폴더 열기", command=self.open_folder).pack(side="left", padx=6)
        ttk.Button(rb, text="저장 위치 변경", command=self.change_dir).pack(side="left")
        self.dir_lbl = ttk.Label(rb, text="", foreground=HINT)
        self.dir_lbl.pack(side="left", padx=8)
        self._show_dir()

        nb = ttk.Notebook(box4)
        nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        # 결과 표
        tframe = ttk.Frame(nb)
        cols = ("name", "func", "org", "conf", "ident")
        self.tree = ttk.Treeview(tframe, columns=cols, show="headings", height=6)
        for c, t, w in zip(cols, ("서열 이름", "추정 기능", "유래 생물", "신뢰도", "일치도(%)"),
                           (140, 300, 190, 60, 80)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c in ("conf", "ident") else "w")
        for k, v in CONF_COLORS.items():
            self.tree.tag_configure(k, background=v, foreground="#000000")  # 다크 모드에서도 글자 보이게
        tsb = ttk.Scrollbar(tframe, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        nb.add(tframe, text="결과 요약")
        # 진행 기록
        lframe = ttk.Frame(nb)
        self.logbox = tk.Text(lframe, height=8, font=self.mono, state="disabled", wrap="word")
        lsb = ttk.Scrollbar(lframe, command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=lsb.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        nb.add(lframe, text="진행 기록")
        self.nb = nb

    # ------------------------------------------------------------------ 입력
    def _on_modified(self, _e=None):
        if self.input.edit_modified():
            self.input.edit_modified(False)
            self._update_count()

    def _update_count(self):
        recs = engine.parse_sequences(self.input.get("1.0", "end"))
        if not recs:
            self.count_lbl.configure(text="서열 0개", foreground=HINT)
        else:
            nuc = sum(1 for _, s in recs if engine.seq_type(s) == "nucleotide")
            self.count_lbl.configure(text=f"서열 {len(recs)}개 인식 (염기 {nuc} · 단백질 {len(recs) - nuc})",
                                     foreground=ACCENT)

    def open_file(self):
        paths = filedialog.askopenfilenames(
            title="서열 파일 선택",
            filetypes=[("서열 파일", "*.fasta *.fa *.fas *.fna *.faa *.fsa *.seq *.txt"), ("모든 파일", "*.*")])
        if not paths:
            return
        text = ""
        for p in paths:
            try:
                with open(p, encoding="utf-8", errors="replace") as f:
                    text += f.read().strip() + "\n"
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"파일을 읽을 수 없습니다:\n{p}\n\n{e}")
                return
        self.input.delete("1.0", "end")
        self.input.insert("1.0", text)
        self._update_count()

    def clear_input(self):
        self.input.delete("1.0", "end")
        self._update_count()

    def insert_example(self):
        ex = os.path.join(engine.SCRIPT_DIR, "example.fasta")
        if os.path.exists(ex):
            with open(ex, encoding="utf-8") as f:
                self.input.delete("1.0", "end")
                self.input.insert("1.0", f.read())
            self._update_count()

    # ------------------------------------------------------------------ 결과 위치
    def _show_dir(self):
        self.dir_lbl.configure(text=f"저장 위치: {self.out_dir}")

    def change_dir(self):
        d = filedialog.askdirectory(title="결과 엑셀을 저장할 폴더", initialdir=self.out_dir)
        if d:
            self.out_dir = d
            self._show_dir()

    def open_folder(self):
        os.makedirs(self.out_dir, exist_ok=True)
        open_path(self.out_dir)

    def open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            open_path(self.result_path)

    # ------------------------------------------------------------------ 실행
    def log(self, line):
        self.q.put(("log", line))

    def start(self):
        records = engine.parse_sequences(self.input.get("1.0", "end"))
        if not records:
            messagebox.showwarning(APP_TITLE, "서열을 먼저 넣어 주세요.\n(붙여넣기 또는 [서열 파일 열기])")
            return
        local = self.mode.get() == "local"
        if local and not engine.local_db_exists(engine.DEFAULT_LOCAL_DB):
            messagebox.showerror(APP_TITLE, "빠른 검색용 데이터베이스(blastdb 폴더)가 없습니다.\n"
                                            "프로그램 폴더 전체가 그대로 있는지 확인해 주세요.")
            return
        if not local:
            mins = max(5, len(records) * 3)
            if not messagebox.askyesno(APP_TITLE, f"정밀 검색은 NCBI 서버 상황에 따라 오래 걸립니다 "
                                                  f"(서열 {len(records)}개 → 대략 {mins}분 이상).\n"
                                                  "창을 닫지 말고 기다려 주세요. 시작할까요?"):
                return
        os.makedirs(self.out_dir, exist_ok=True)
        out = os.path.join(self.out_dir, f"BLAST결과_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        opts = engine.Options(local=local, output=out)

        self.tree.delete(*self.tree.get_children())
        self.logbox.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.configure(state="disabled")
        self.result_path = None
        self.open_xlsx_btn.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress.start(12)
        self.status.configure(text="분석 중...")
        self.nb.select(1)

        engine.LOG_HOOK = self.log

        def work():
            try:
                path, summaries = engine.run_analysis(records, opts)
                self.q.put(("done", (path, summaries)))
            except engine.Cancelled:
                self.q.put(("cancelled", None))
            except Exception as e:  # 사용자에게 보여줄 오류
                self.q.put(("error", str(e)))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def stop(self):
        engine.CANCEL.set()
        self.status.configure(text="중지하는 중...")
        self.stop_btn.configure(state="disabled")

    def _finish(self):
        self.progress.stop()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self.logbox.configure(state="normal")
                    self.logbox.insert("end", data + "\n")
                    self.logbox.see("end")
                    self.logbox.configure(state="disabled")
                elif kind == "done":
                    path, summaries = data
                    self._finish()
                    self.result_path = path
                    for s in summaries:
                        func = s["function"] + (f"  ({s['gene']})" if s.get("gene") else "")
                        self.tree.insert("", "end", values=(
                            s["name"], func, s["organism"], s["confidence"],
                            "" if s["identity"] is None else s["identity"]), tags=(s["confidence"],))
                    self.open_xlsx_btn.configure(state="normal")
                    self.status.configure(text=f"완료 ({len(summaries)}개)")
                    self.nb.select(0)
                    if messagebox.askyesno(APP_TITLE, f"분석이 끝났습니다. 서열 {len(summaries)}개\n\n"
                                                      f"{path}\n\n엑셀 파일을 지금 열까요?"):
                        open_path(path)
                elif kind == "cancelled":
                    self._finish()
                    self.status.configure(text="중지됨")
                    self.log("중지되었습니다. 끝난 검색은 저장되어 있어 다시 시작하면 바로 불러옵니다.")
                elif kind == "error":
                    self._finish()
                    self.status.configure(text="오류")
                    messagebox.showerror(APP_TITLE, f"문제가 생겼습니다:\n\n{data}\n\n"
                                                    "'진행 기록' 탭 내용을 개발자에게 보내 주세요.")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


def selftest(out_path):
    """빌드 검증용: 화면 없이 예시 서열을 로컬 검색해 엑셀을 만든다. 성공 시 종료코드 0."""
    log_path = out_path + ".log"
    with open(log_path, "w", encoding="utf-8") as lf:
        engine.LOG_HOOK = lambda line: (lf.write(line + "\n"), lf.flush())
        try:
            with open(os.path.join(engine.SCRIPT_DIR, "example.fasta"), encoding="utf-8") as f:
                records = engine.parse_sequences(f.read())
            _, summaries = engine.run_analysis(records, engine.Options(local=True, output=out_path))
            funcs = " | ".join(s["function"] for s in summaries)
            lf.write("FUNCTIONS: " + funcs + "\n")
            ok = "lactamase" in funcs.lower() and "fluorescent" in funcs.lower()
        except Exception as e:
            lf.write(f"SELFTEST ERROR: {e!r}\n")
            ok = False
        lf.write("SELFTEST " + ("PASS" if ok else "FAIL") + "\n")
    sys.exit(0 if ok else 1)


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--selftest":
        selftest(os.path.abspath(sys.argv[2]))
    try:
        if os.name == "nt":
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 고해상도 화면에서 흐리지 않게
    except Exception:
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
