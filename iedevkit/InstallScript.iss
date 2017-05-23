[_ISTool]
EnableISX=false

[Files]
Source: Examples\Menu and Toolbar Example\cold.ico; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\example.html; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\exampleUI.html; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\Form1.frm; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\Form1.frx; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\hot.ico; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\Project1.vbp; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\Menu and Toolbar Example\Project1.vbw; DestDir: {app}\Examples\Menu and Toolbar Example
Source: Examples\demo.exe; DestDir: {app}\Examples\
Source: Examples\dhtml_events.html; DestDir: {app}\Examples\
Source: Examples\external_demo.html; DestDir: {app}\Examples\
Source: Examples\form1.frm; DestDir: {app}\Examples\
Source: Examples\Project1.vbw; DestDir: {app}\Examples\
Source: Examples\Project1.vbp; DestDir: {app}\Examples\
Source: IEDevKit.chm; DestDir: {app}
Source: IEDevKit2.dll; DestDir: {app}; Flags: regserver
Source: Product Homepage.url; DestDir: {app}

[Dirs]
Name: {app}\Examples
Name: {app}\Examples\Menu and Toolbar Example

[Run]
Filename: {app}\Examples\Project1.vbp; WorkingDir: {app}; Description: Startup Sample Project; Flags: postinstall shellexec

[Icons]
Name: {group}\IE DevKit Helpfile; Filename: {app}\IEDevKit.chm; WorkingDir: {app}; Flags: runmaximized
Name: {group}\Extend Wb Sample Project ; Filename: {app}\Examples\Project1.vbp; WorkingDir: {app}\Examples\
Name: {group}\Menu & Toolbar Example; Filename: {app}\Examples\Menu and Toolbar Example\Project1.vbp; WorkingDir: {app}\Examples\Menu and Toolbar Example\
Name: {group}\Product Homepage; Filename: {app}\Product Homepage.url; WorkingDir: {app}
Name: {group}\Uninstall; Filename: {app}\unins000.exe; WorkingDir: {app}

[Setup]
AppPublisher=David Zimmer
AppPublisherURL=http://sandsprite.com/iedevkit
AppSupportURL=http://sandsprite.com/iedevkit
AppUpdatesURL=http://sandsprite.com/iedevkit
DefaultGroupName=IEDevKit
Compression=bzip/9
UninstallDisplayIcon={app}\compil32.exe
AppCopyright=SandSprite.com
AppName=IEDev Kit
AppVerName=IE DevKit 1.0
OutputDir=./
OutputBaseFilename=iedevkit
DefaultDirName={pf}\SandSprite\ieDevKit

