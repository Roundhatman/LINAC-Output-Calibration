; Output Calculator Inno setup script
; L.Swarnajith

#define MyAppName "Output Calculator NCI"
#define MyAppVersion "1.0"
#define MyAppPublisher "L.Swarnajith | Spetsgruppa Commandos"
#define MyAppExeName "nci_calc_001.exe"

[Setup]
AppId={{989A25FB-F253-43F4-B337-5606D52F4EF6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequiredOverridesAllowed=dialog
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "D:\Academic\USJP EES Course File\Special II\PHY 455 6.0 Internship\NCISL\AppDevelopment\nci_calc_001\nci_calc_001\bin\Debug\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Academic\USJP EES Course File\Special II\PHY 455 6.0 Internship\NCISL\AppDevelopment\nci_calc_001\nci_calc_001\bin\Debug\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

