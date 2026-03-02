[Setup]
AppName=Converter
AppVersion=0.0.0
DefaultDirName={pf}\Converter
DefaultGroupName=Converter
OutputDir=installer
OutputBaseFilename=ConverterSetup
Compression=lzma
SolidCompression=yes
SetupIconFile=File_Converter_Logo.ico

[Files]
Source: "dist\Converter.exe"; DestDir: "{app}"; DestName: "Converter.exe"; Flags: ignoreversion

[Icons]
Name: "{group}\Converter"; Filename: "{app}\Converter.exe"; IconFilename: "{app}\Converter.exe"
Name: "{commondesktop}\Converter"; Filename: "{app}\Converter.exe"; IconFilename: "{app}\Converter.exe"

[Run]
Filename: "{app}\Converter.exe"; Description: "Launch Converter"; Flags: nowait postinstall skipifsilent
