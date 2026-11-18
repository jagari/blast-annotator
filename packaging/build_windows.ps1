# =====================================================================
#  서열 기능 자동 분석기 - Windows 빌드 스크립트
#  결과물: build_win\BLAST_Analyzer_Setup.exe (설치 파일)
#          build_win\BLAST_Analyzer_portable.zip (압축만 풀면 되는 버전)
#
#  Windows PC 에서 직접 빌드:  PowerShell 에서
#     powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
#  (Python 3.12 이 설치되어 있어야 함. GitHub Actions 에서도 이 스크립트를 그대로 사용)
# =====================================================================
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"   # Invoke-WebRequest 속도 개선

$Root  = Split-Path -Parent $PSScriptRoot
$Work  = Join-Path $Root "build_win"
$App   = "BLAST_Analyzer"
$BlastVer = "2.17.0"
New-Item -ItemType Directory -Force -Path $Work | Out-Null
Set-Location $Work

function Step($msg) { Write-Host "`n=== $msg ===" -ForegroundColor Cyan }

Step "1. Python 패키지 설치"
python -m pip install --upgrade pip | Out-Null
python -m pip install openpyxl pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip install 실패" }

Step "2. NCBI BLAST+ $BlastVer (Windows) 내려받기 + 무결성 확인"
$tar = "ncbi-blast-$BlastVer+-x64-win64.tar.gz"
$url = "https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/$BlastVer/$tar"
if (-not (Test-Path $tar)) {
    Invoke-WebRequest $url -OutFile $tar
    Invoke-WebRequest "$url.md5" -OutFile "$tar.md5"
}
$expected = ((Get-Content "$tar.md5") -split "\s+")[0].ToLower()
$actual = (Get-FileHash $tar -Algorithm MD5).Hash.ToLower()
if ($expected -ne $actual) { Remove-Item $tar; throw "BLAST+ 파일 손상 (MD5 불일치). 다시 실행하세요." }
if (-not (Test-Path "ncbi-blast-$BlastVer+")) { tar -xzf $tar }
$BlastBin = Join-Path $Work "ncbi-blast-$BlastVer+\bin"

Step "3. UniProt Swiss-Prot 데이터베이스 만들기"
$DbDir = Join-Path $Work "blastdb"
New-Item -ItemType Directory -Force -Path $DbDir | Out-Null
Set-Location $DbDir
if (-not (Test-Path "swissprot.pin")) {
    Invoke-WebRequest "https://ftp.uniprot.org/pub/databases/uniprot/current_release/knowledgebase/complete/uniprot_sprot.fasta.gz" -OutFile "sprot.fasta.gz"
    python -c "import gzip,shutil; shutil.copyfileobj(gzip.open('sprot.fasta.gz','rb'), open('sprot.fasta','wb'))"
    & "$BlastBin\makeblastdb.exe" -in sprot.fasta -dbtype prot -parse_seqids -out swissprot -title "UniProt Swiss-Prot"
    Remove-Item sprot.fasta, sprot.fasta.gz
}
& "$BlastBin\blastdbcmd.exe" -db swissprot -info
if ($LASTEXITCODE -ne 0) { throw "DB 확인 실패" }
Set-Location $Work

Step "4. PyInstaller 로 실행 파일 만들기"
python -m PyInstaller --noconfirm --onedir --windowed --name $App `
    --distpath "$Work\dist" --workpath "$Work\pyi" --specpath "$Work\pyi" `
    (Join-Path $Root "gui.py")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 실패" }
$Dist = Join-Path $Work "dist\$App"

Step "5. BLAST+ / DB / 예시 파일 넣기"
New-Item -ItemType Directory -Force -Path "$Dist\blast\bin", "$Dist\blastdb" | Out-Null
foreach ($exe in "blastx", "blastp", "blastn", "blastdbcmd") {
    Copy-Item "$BlastBin\$exe.exe" "$Dist\blast\bin\"
}
Copy-Item "$BlastBin\*.dll" "$Dist\blast\bin\" -ErrorAction SilentlyContinue
Copy-Item "$DbDir\swissprot.*" "$Dist\blastdb\"
Copy-Item (Join-Path $Root "example.fasta") $Dist
Copy-Item (Join-Path $Root "packaging\사용설명서.txt") $Dist

Step "6. 자체 테스트 (실행 파일로 예시 서열 분석)"
$testOut = Join-Path $Work "selftest.xlsx"
$p = Start-Process -FilePath "$Dist\$App.exe" -ArgumentList "--selftest", "`"$testOut`"" -Wait -PassThru
Get-Content "$testOut.log" -Encoding UTF8 | Select-Object -Last 5
if ($p.ExitCode -ne 0) { throw "자체 테스트 실패 - 위 로그 확인" }
Write-Host "자체 테스트 통과" -ForegroundColor Green

Step "7. 압축 파일 (설치 없이 쓰는 버전)"
$zip = Join-Path $Work "${App}_portable.zip"
if (Test-Path $zip) { Remove-Item $zip }
$7z = "C:\Program Files\7-Zip\7z.exe"
if (Test-Path $7z) { & $7z a -tzip -mx=7 $zip "$Dist" | Out-Null }   # 한글 파일명 깨짐 방지
else { Compress-Archive -Path $Dist -DestinationPath $zip }

Step "8. 설치 파일 (Inno Setup)"
$iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    Write-Host "Inno Setup 이 없어 설치 파일은 건너뜁니다 (choco install innosetup 으로 설치 가능)" -ForegroundColor Yellow
} else {
    $ver = if ($env:VERSION) { $env:VERSION.TrimStart("v") } else { "1.0" }
    & $iscc "/DSourceDir=$Dist" "/DOutDir=$Work" "/DAppVersion=$ver" (Join-Path $Root "packaging\installer.iss")
    if ($LASTEXITCODE -ne 0) { throw "설치 파일 생성 실패" }
}

Step "완료"
Get-ChildItem $Work -Filter "${App}*" | Format-Table Name, @{n="MB"; e={[math]::Round($_.Length/1MB)}}
