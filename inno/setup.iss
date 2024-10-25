[Setup]
AppName=AiTexter
AppVersion=1.01
DefaultDirName={userappdata}\AiTexter
DefaultGroupName=AiTexter
OutputDir=.\Output
OutputBaseFilename=AiTexterInstaller
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes

[Files]
; Füge hier deine Dateien hinzu, die installiert werden sollen
Source: "D:\Programmieren\AiTexter\dist\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Programmieren\AiTexter\dist\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Desktop-Icon für das Programm erstellen
Name: "{userdesktop}\AiTexter"; Filename: "{app}\main.exe"; IconFilename: "{app}\icon.ico"

[Run]
; Programm nach Installation starten
Filename: "{app}\main.exe"; Description: "{cm:LaunchProgram,AiTexter}"; Flags: nowait postinstall skipifsilent
