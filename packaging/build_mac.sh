#!/bin/bash
# =====================================================================
#  서열 기능 자동 분석기 - Mac 빌드 스크립트
#  결과물: build_mac/BLAST_Analyzer_Mac_<AppleSilicon|Intel>.dmg
#
#  직접 빌드:  PYTHON=/path/to/python3.12 bash packaging/build_mac.sh
#  (Tk 8.6 이상이 들어있는 Python 3.10+ 필요. macOS 기본 /usr/bin/python3 은 Tk 가 낡아서 안 됨)
#  GitHub Actions 에서도 이 스크립트를 그대로 사용
# =====================================================================
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$ROOT/build_mac"
APP="BLAST_Analyzer"
BLAST_VER="2.17.0"
PY="${PYTHON:-python3}"
# 상대 경로로 지정해도 되도록 절대 경로로 변환 (아래에서 cd 하므로)
case "$PY" in */*) PY="$(cd "$(dirname "$PY")" && pwd)/$(basename "$PY")" ;; esac

case "$(uname -m)" in
  arm64)  BLAST_ARCH="aarch64"; LABEL="AppleSilicon" ;;
  x86_64) BLAST_ARCH="x64";     LABEL="Intel" ;;
  *) echo "지원하지 않는 CPU: $(uname -m)"; exit 1 ;;
esac

step() { printf '\n=== %s ===\n' "$1"; }
mkdir -p "$WORK"
cd "$WORK"

step "1. Python 확인 / 패키지 설치"
"$PY" -c "import sys, tkinter; assert sys.version_info >= (3, 10) and tkinter.TkVersion >= 8.6, 'old'; print(sys.version.split()[0], 'Tk', tkinter.TkVersion)" \
  || { echo "Python 3.10+ / Tk 8.6+ 가 필요합니다. PYTHON=... 으로 지정하세요."; exit 1; }
"$PY" -m pip install -q openpyxl pyinstaller

step "2. NCBI BLAST+ $BLAST_VER (Mac $BLAST_ARCH) 내려받기 + 무결성 확인"
TAR="ncbi-blast-$BLAST_VER+-$BLAST_ARCH-macosx.tar.gz"
URL="https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/$BLAST_VER/$TAR"
if [ ! -f "$TAR" ]; then
  curl -fsSLO "$URL"
  curl -fsSLO "$URL.md5"
fi
[ "$(awk '{print $1}' "$TAR.md5")" = "$(md5 -q "$TAR")" ] || { rm -f "$TAR"; echo "BLAST+ 파일 손상 (MD5 불일치)"; exit 1; }
[ -d "ncbi-blast-$BLAST_VER+" ] || tar xzf "$TAR"
BIN="$WORK/ncbi-blast-$BLAST_VER+/bin"

step "3. UniProt Swiss-Prot 데이터베이스"
mkdir -p blastdb
cd blastdb
if [ ! -f swissprot.pin ]; then
  if [ -f "$ROOT/blastdb/swissprot.pin" ]; then
    echo "프로젝트의 기존 DB 사용"
    cp "$ROOT"/blastdb/swissprot.* .
  else
    curl -fsSL -o sprot.fasta.gz \
      "https://ftp.uniprot.org/pub/databases/uniprot/current_release/knowledgebase/complete/uniprot_sprot.fasta.gz"
    gunzip -f sprot.fasta.gz
    "$BIN/makeblastdb" -in sprot.fasta -dbtype prot -parse_seqids -out swissprot -title "UniProt Swiss-Prot" || true
    rm -f sprot.fasta
  fi
fi
"$BIN/blastdbcmd" -db swissprot -info | head -2
cd "$WORK"

step "4. PyInstaller 로 .app 만들기"
"$PY" -m PyInstaller --noconfirm --windowed --name "$APP" \
  --osx-bundle-identifier io.github.jagari.blastanalyzer \
  --distpath "$WORK/dist" --workpath "$WORK/pyi" --specpath "$WORK/pyi" \
  "$ROOT/gui.py" > "$WORK/pyinstaller.log" 2>&1 || { tail -30 "$WORK/pyinstaller.log"; exit 1; }
APPDIR="$WORK/dist/$APP.app"
RES="$APPDIR/Contents/Resources"

step "5. BLAST+ / DB / 예시 파일 넣기"
mkdir -p "$RES/blast/bin" "$RES/blastdb"
for exe in blastx blastp blastn blastdbcmd; do cp "$BIN/$exe" "$RES/blast/bin/"; done
cp blastdb/swissprot.* "$RES/blastdb/"
cp "$ROOT/example.fasta" "$RES/"
# Finder / Dock 에 보이는 이름
/usr/libexec/PlistBuddy -c "Add :CFBundleDisplayName string 서열 기능 분석기" "$APPDIR/Contents/Info.plist" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Set :CFBundleDisplayName 서열 기능 분석기" "$APPDIR/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString $(echo "${VERSION:-1.0}" | sed 's/^v//')" "$APPDIR/Contents/Info.plist"

step "6. 서명 (ad-hoc: Apple 개발자 인증서 없이 실행 가능하게)"
codesign --force --deep --sign - "$APPDIR"
codesign --verify --deep --strict "$APPDIR" && echo "서명 확인 OK"

step "7. 자체 테스트 (.app 으로 예시 서열 분석)"
rm -f "$WORK/selftest.xlsx" "$WORK/selftest.xlsx.log"
if ! "$APPDIR/Contents/MacOS/$APP" --selftest "$WORK/selftest.xlsx"; then
  cat "$WORK/selftest.xlsx.log" || true
  echo "자체 테스트 실패"; exit 1
fi
tail -2 "$WORK/selftest.xlsx.log"

step "8. DMG 만들기"
STAGE="$WORK/dmg"
rm -rf "$STAGE"
mkdir -p "$STAGE"
cp -R "$APPDIR" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
cp "$ROOT/packaging/사용설명서.txt" "$STAGE/"
DMG="$WORK/${APP}_Mac_${LABEL}.dmg"
rm -f "$DMG"
hdiutil create -quiet -volname "서열 기능 분석기" -srcfolder "$STAGE" -ov -format UDZO "$DMG"

step "완료"
ls -lh "$DMG"
