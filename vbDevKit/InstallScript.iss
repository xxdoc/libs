[_ISTool]
EnableISX=false

[Files]
Source: vbDevKit.dll; DestDir: {app}; CopyMode: normal; Flags: regserver
Source: vbDevKit.chm; DestDir: {app}
Source: Code Examples\example.ini; DestDir: {app}\Code Examples\
Source: Code Examples\Form1.frm; DestDir: {app}\Code Examples\
Source: Code Examples\Project1.vbp; DestDir: {app}\Code Examples\
Source: Code Examples\Project1.vbw; DestDir: {app}\Code Examples\

[Dirs]
Name: {app}\Code Examples

[Run]
Filename: {app}\vbDevKit.chm; WorkingDir: {app}; Description: View Product Documentation; Flags: postinstall shellexec

[Icons]
Name: {group}\VbDevKit Help; Filename: {app}\vbDevKit.chm; WorkingDir: {app}
Name: {group}\Vb Code Examples; Filename: {app}\Code Examples\Project1.vbp; WorkingDir: {app}\Code Examples\

[Setup]
AppPublisher=David Zimmer
AppPublisherURL=http://sandsprite.com
AppSupportURL=http://sandsprite.com
AppUpdatesURL=http://sandsprite.com
DefaultGroupName=vbDevKit
Compression=bzip/9
UninstallDisplayIcon={app}\compil32.exe
AppCopyright=SandSprite.com
AppName=VbDevKit
AppVerName=1.0
OutputDir=./
OutputBaseFilename=vbdevkit
DefaultDirName={pf}\SandSprite\vbDevKit

