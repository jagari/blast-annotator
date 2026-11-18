; 서열 기능 자동 분석기 - Inno Setup 설치 스크립트 (build_windows.ps1 이 호출)
#ifndef SourceDir
  #define SourceDir "..\build_win\dist\BLAST_Analyzer"
#endif
#ifndef AppVersion
  #define AppVersion "1.0"
#endif
#ifndef OutDir
  #define OutDir "..\build_win"
#endif

[Setup]
AppName=서열 기능 자동 분석기
AppVersion={#AppVersion}
AppPublisher=BLAST Analyzer
DefaultDirName={autopf}\BLAST_Analyzer
DefaultGroupName=서열 기능 자동 분석기
OutputDir={#OutDir}
OutputBaseFilename=BLAST_Analyzer_Setup
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
WizardStyle=modern
UninstallDisplayIcon={app}\BLAST_Analyzer.exe

[Languages]
#if FileExists(AddBackslash(CompilerPath) + "Languages\Korean.isl")
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
#else
Name: "default"; MessagesFile: "compiler:Default.isl"
#endif

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; Flags: checkedonce

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\서열 기능 자동 분석기"; Filename: "{app}\BLAST_Analyzer.exe"
Name: "{group}\사용설명서"; Filename: "{app}\사용설명서.txt"
Name: "{group}\제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\서열 기능 자동 분석기"; Filename: "{app}\BLAST_Analyzer.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\BLAST_Analyzer.exe"; Description: "지금 실행"; Flags: nowait postinstall skipifsilent
