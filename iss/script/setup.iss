[Setup]
OutputDir=C:\Documents and Settings\username\My Documents\Development\vb6\eikofp\iss\setup
SolidCompression=true
VersionInfoVersion=0.8.7
VersionInfoCompany=Eiko Software
ShowLanguageDialog=no
AppID={{070B65A6-7C5E-433C-9426-E669A6CC5E12}
UninstallLogMode=overwrite
SetupLogging=true
AppCopyright=Rossano M. Szczepanski
AppName=Eiko Finanças Pessoais
AppVerName=Eiko Finanças Pessoais 0.8.7
AllowCancelDuringInstall=false
DisableDirPage=true
DefaultDirName={pf}\Eiko Soft\Finanças Pessoais
DefaultGroupName=Eiko Finanças Pessoais
DisableProgramGroupPage=true
AppVersion=Eiko Finanças Pessoais 0.8.7
AppPublisher=Rossano M. Szczepanski
AppPublisherURL=
AppUpdatesURL=
UninstallDisplayName=Eiko Finanças Pessoais 0.8.7
AppContact=rm.szc81@gmail.com
UninstallDisplayIcon={app}\EikoFP.ico

[Languages]
Name: brazilianportuguese; MessagesFile: compiler:Languages\BrazilianPortuguese.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Languages: 

[Icons]
Name: {group}\Eiko Finanças Pessoais; Filename: {app}\EikoFP.exe; WorkingDir: {app}; IconFilename: {app}\EikoFP.ico
Name: {group}\Eiko Finanças Pessoais - Modo Debug; Filename: {app}\EikoFP.exe; WorkingDir: {app}; IconFilename: {app}\EikoFP.ico; Parameters: /debug
Name: {group}\Eiko Finanças Pessoais - Modo Offline; Filename: {app}\EikoFP.exe; WorkingDir: {app}; IconFilename: {app}\EikoFP.ico; Parameters: /offline
Name: {group}\{cm:UninstallProgram,EikoFP}; Filename: {uninstallexe}
Name: {commondesktop}\Eiko Finanças Pessoais; Filename: {app}\EikoFP.exe; Tasks: desktopicon; WorkingDir: {app}; IconFilename: {app}\EikoFP.ico

[Run]
Filename: {app}\EikoFP.exe; WorkingDir: {app}; Flags: postinstall nowait unchecked; Description: Executar o Eiko Finanças Pessoais agora.

[Files]

;main files
Source: ..\..\exe\EikoFP.exe; DestDir: {app}; Attribs: readonly; Flags: overwritereadonly ignoreversion uninsremovereadonly replacesameversion
Source: ..\..\exe\EikoFP.ico; DestDir: {app}; Attribs: readonly; Flags: overwritereadonly ignoreversion uninsremovereadonly replacesameversion
Source: ..\..\exe\EikoFP.exe.manifest; DestDir: {app}; Attribs: readonly; Flags: overwritereadonly ignoreversion uninsremovereadonly replacesameversion
Source: ..\..\chm\eikoFP.chm; DestDir: {app}; Attribs: readonly; Flags: overwritereadonly ignoreversion uninsremovereadonly replacesameversion

;shared files
Source: ..\..\dll\libmySQL.dll; DestDir: {sys}; Attribs: readonly; Flags: uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\..\dll\MyVbQL.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\..\dll\newobjectspack1.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\..\dll\SmartMenuXP.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\..\dll\SmartSubClass.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\..\dll\SQLITE3COMUTF8.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly

;support files
Source: ..\support\ASYCFILT.DLL; DestDir: {sys}; Attribs: readonly; Flags: restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\COMCAT.DLL; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\msvbvm60.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\OLEAUT32.DLL; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\OLEPRO32.DLL; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\STDOLE2.TLB; DestDir: {sys}; Attribs: readonly; Flags: restartreplace sharedfile uninsneveruninstall onlyifdoesntexist overwritereadonly

;component files
Source: ..\..\ocx\SMARTMENUXP.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver uninsneveruninstall onlyifdoesntexist overwritereadonly
Source: ..\support\COMDLG32.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly
Source: ..\support\HHCTRL.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly
Source: ..\support\MSCHRT20.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly
Source: ..\support\MSCOMCT2.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly
Source: ..\support\MSCOMCTL.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly
Source: ..\support\MSFLXGRD.OCX; DestDir: {sys}; Attribs: readonly; Flags: regserver sharedfile onlyifdoesntexist restartreplace uninsneveruninstall overwritereadonly

;dependency files
Source: ..\support\hh.exe; DestDir: {sys}; Attribs: readonly; Flags: restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
Source: ..\support\itircl.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
Source: ..\support\itss.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
Source: ..\support\MSVCRT.DLL; DestDir: {sys}; Attribs: readonly; Flags: restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
Source: ..\support\scrrun.dll; DestDir: {sys}; Attribs: readonly; Flags: regserver restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
Source: ..\support\shfolder.dll; DestDir: {sys}; Attribs: readonly; Flags: restartreplace sharedfile uninsneveruninstall overwritereadonly onlyifdoesntexist
