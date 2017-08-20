program DemoApp;

uses
  Forms,
  Main in 'Main.pas' {Form1},
  BTMemoryModule in '..\BTMemoryModule\BTMemoryModule.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
