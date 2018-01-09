#define MyAppName "Kempus"
#define MyAppVersion "1.0"
#define MyAppPublisher "JR Company © 2014"
#define MyAppURL "http://www.JR-07.Net/"
#define MyAppExeName "Keuangan Kampus.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{91FCDD7B-DAAB-42D4-82F3-E8952809E8D3}}
AppName={#MyAppName}
AppVersion=1.0
VersionInfoVersion= 1.0.0.0
VersionInfoDescription=Kempus
AppVerName={#MyAppName} {#MyAppVersion}
AppCopyright=JR Company (C) 2014, Inc.
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=C:\Users\JR\Desktop\Keuangan Kampus\Installer\License.txt
InfoBeforeFile=C:\Users\JR\Desktop\Keuangan Kampus\Installer\Show Before.txt
InfoAfterFile=C:\Users\JR\Desktop\Keuangan Kampus\Installer\Show After.txt
OutputDir=C:\Users\JR\Desktop\Keuangan Kampus\Installer
OutputBaseFilename=Kempus
SetupIconFile=C:\Users\JR\Desktop\Keuangan Kampus\Image\logo_setup.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Keuangan Kampus.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Log.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Crystl32.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Crystl32.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Crystl32.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Crystl32.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.DEP"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.SRG"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.DEP"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.SRG"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace 
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.DEP"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.SRG"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.DEP"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.SRG"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.DEP"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.DEP"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSDATGRD.OCX"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSBIND.DLL"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSBIND.DLL"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSBIND.DLL"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSSTDFMT.DLL"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSSTDFMT.DLL"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\MSSTDFMT.DLL"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\msvbvm60.dll"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\msvbvm60.dll"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\VB6.OLB"; DestDir: "C:\Program Files\Microsoft Visual Studio\VB98\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\stdole2.tlb"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\DAO350.dll"; DestDir: "C:\Program Files\Common Files\Microsoft Shared\DAO\"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\DAO350.dll"; DestDir: "C:\Program Files\Common Files\Microsoft Shared\DAO\"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\msador28.tlb"; DestDir: "C:\Program Files\Common Files\System\ado\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\msado60.tlb"; DestDir: "C:\Program Files\Common Files\System\ado\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Database\*"; DestDir: "{app}\Database"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\BuktiPembayaran.rpt"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\LaporanKeuangan.rpt"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\logger_asserts.txt"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\JR\Desktop\Keuangan Kampus\Installer\Component\Crystal Report Redistribution 8.5 Runtime\*"; DestDir: "{app}\Component\Crystal Report Redistribution 8.5 Runtime"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\Component\Crystal Report Redistribution 8.5 Runtime\setup.exe"; Description: "Menginstal Crystal Report 8.5 Runtime Destribution"; StatusMsg: "Menginstal Crystal Reports 8.5"; Flags: skipifsilent
Filename: "{app}\{#MyAppExeName}"; Description: "Jalankan Kempus"; Flags: nowait postinstall skipifsilent
