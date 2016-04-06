[_ISTool]
EnableISX=false

[Files]
Source: .\Form2.frx; DestDir: {app}
Source: .\Form2.frm; DestDir: {app}
Source: .\spSubclass.dll; DestDir: {app}; Flags: regserver
Source: .\test.vbp; DestDir: {app}
Source: .\test.vbw; DestDir: {app}
Source: .\LICENSE.txt; DestDir: {app}

[Run]
Filename: {app}\LICENSE.txt; Flags: postinstall shellexec; Description: View License Agreement / Documentation;

[Icons]
Name: {group}\spSubClass\License && Readme; Filename: {app}\LICENSE.txt;  Comment: License and Readme file
Name: {group}\spSubClass\Example Project; Filename: {app}\test.vbp; Comment: Test Project

[Setup]
AppName=Sandsprite Subclass Component
AppVerName=Sandsprite Subclass Component
AppVersion=1.0
AppPublisher=David Zimmer
AppPublisherURL=http://sandsprite.com/subclass
AppSupportURL=http://sandsprite.com/subclass
AppUpdatesURL=http://sandsprite.com/subclass
AppCopyright=Sleuth , copyright © 2003 David Zimmer
DefaultDirName={pf}\SandSprite\spSubClass\
DefaultGroupName=SandSprite
Compression=bzip/9
OutputBaseFilename=subclassSetup
OutputDir=.\
UninstallDisplayIcon={app}\compil32.exe
AllowNoIcons=yes
WizardImageFile=compiler:WizModernImage.bmp
WizardSmallImageFile=compiler:WizModernSmallImage.bmp
