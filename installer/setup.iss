[Setup]
AppName=DS Typing Trainer
AppVersion=1.0
DefaultDirName={pf}\DSTypingTrainer
DefaultGroupName=DS Typing Trainer
OutputDir=output
OutputBaseFilename=DSTypingTrainer_Setup
Compression=lzma
SolidCompression=yes

[Files]
Source: "..\dist\TypingTrainer.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\sounds\*"; DestDir: "{app}\sounds"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\DS Typing Trainer"; Filename: "{app}\TypingTrainer.exe"
Name: "{commondesktop}\DS Typing Trainer"; Filename: "{app}\TypingTrainer.exe"

[Run]
Filename: "{app}\TypingTrainer.exe"; Description: "Launch Typing Trainer"; Flags: nowait postinstall skipifsilent
