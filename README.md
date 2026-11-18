# 서열 기능 자동 분석기 (BLAST Analyzer)

기능을 모르는 **염기서열(DNA)이나 단백질 서열**을 넣으면, BLAST로 비슷한 서열을 찾아
**어떤 기능의 유전자/단백질인지 자동으로 정리해 엑셀 파일로** 만들어 주는 Windows 프로그램입니다.
서열을 하나씩 BLAST 웹사이트에 붙여넣던 반복 작업을 없애 줍니다.

## ⬇️ 다운로드

### [👉 최신 버전 다운로드 (Releases)](https://github.com/jagari/blast-annotator/releases/latest)

| 파일 | 설명 |
|---|---|
| [**BLAST_Analyzer_Setup.exe** (92MB)](https://github.com/jagari/blast-annotator/releases/latest/download/BLAST_Analyzer_Setup.exe) | **추천.** 클릭하면 바로 다운로드됩니다. 더블클릭 → 다음 → 설치 |
| [**BLAST_Analyzer_portable.zip** (124MB)](https://github.com/jagari/blast-annotator/releases/latest/download/BLAST_Analyzer_portable.zip) | 설치 없이 쓰는 버전입니다. 압축을 풀고 `BLAST_Analyzer.exe` 실행 |

- Windows 10/11 (64비트)
- 따로 설치할 것이 없습니다. 검색 프로그램과 데이터베이스가 모두 들어 있습니다.
- 처음 실행할 때 파란색 **"Windows의 PC 보호"** 창이 뜨면 **[추가 정보] → [실행]** 을 누르세요.
  (코드 서명 인증서가 없는 프로그램이라 뜨는 안내입니다)

## 사용 방법

1. **서열 넣기**: [서열 파일 열기]로 FASTA/TXT 파일을 고르거나, 입력 칸에 붙여넣기 (여러 개 가능)
2. **검색 방식 고르기**
   - **빠른 검색 (추천)**: 내 컴퓨터에서 검색. 서열당 1~2초, 인터넷 필요 없음
   - **정밀 검색**: NCBI 서버의 전체 DB에서 검색. 서열당 수 분~수십 분, 인터넷 필요
3. **[▶ 분석 시작]** → 결과 표 확인 → 엑셀 파일 열기

엑셀에는 추정 기능, 신뢰도(색 표시), 유래 생물, 유전자명, 일치도, E-value, 단백질 정보 링크가 정리됩니다.
자세한 내용은 프로그램 폴더 안의 `사용설명서.txt` 를 보세요.

## 개발자용

- `blast_annotate.py`: 검색/분석/엑셀 엔진 (명령줄에서도 사용 가능, [사용법.md](사용법.md) 참고)
- `gui.py`: 프로그램 화면 (tkinter)
- `packaging/build_windows.ps1`: Windows 실행 파일과 설치 파일 빌드
- 새 버전 빌드: GitHub → Actions → **Build Windows app** → Run workflow

## 출처 및 라이선스

- 검색 엔진: [NCBI BLAST+](https://blast.ncbi.nlm.nih.gov/) 2.17.0 (미국 정부 저작물, 퍼블릭 도메인)
- 단백질 데이터베이스: [UniProt](https://www.uniprot.org/) Swiss-Prot. The UniProt Consortium,
  [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
- 정밀 검색은 NCBI BLAST 공용 웹 서비스를 사용하며 [NCBI 이용 지침](https://blast.ncbi.nlm.nih.gov/doc/blast-help/developerinfo.html)을 따릅니다.

분석 결과는 "비슷한 서열이 이런 기능을 가진다"는 **추정**입니다. 중요한 결과는 링크된 원본 정보로 직접 확인하세요.
