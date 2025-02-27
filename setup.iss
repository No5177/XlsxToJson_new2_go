[Setup]
AppName=xlsx_to_json
AppVersion=1.0
DefaultDirName={commonpf}\xlsx_to_json
DefaultGroupName=xlsx_to_json
UninstallDisplayIcon={app}\program\python.exe
Compression=lzma2
SolidCompression=yes
OutputDir=.
OutputBaseFilename=xlsx_to_json_Installer
SetupIconFile=program\icon_image\icon_517.ico
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "python-3.13.1-amd64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "program\*"; DestDir: "{app}\program"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "XlsxToJson.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "CheckPython.bat"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{tmp}\python-3.13.1-amd64.exe"; Parameters: "/quiet InstallAllUsers=0 TargetDir=""{app}\python"" PrependPath=1"; Flags: waituntilterminated; StatusMsg: "Installing Python..."
Filename: "{app}\CheckPython.bat"; Flags: runhidden waituntilterminated
Filename: "{app}\python\python.exe"; Parameters: "-m pip install --no-index --find-links ""{app}\program"" -r ""{app}\program\requirements.txt"""; Flags: waituntilterminated

[Icons]
Name: "{group}\xlsx_to_json"; Filename: "{app}\XlsxToJson.bat"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
