unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, BTMemoryModule, StdCtrls, xpman;

type
  TTestCallstd = procedure(f_Text: PChar); stdcall;
  TTestCallcdel = procedure(f_Text: PChar); cdecl;

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    m_TestCallstd: TTestCallstd;
    m_TestCallcdel: TTestCallcdel;
    mp_DllData: Pointer;
    m_DllDataSize: Integer;
    mp_MemoryModule: PBTMemoryModule;
    m_DllHandle: Cardinal;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
  MemoryStream: TMemoryStream;
begin
  MemoryStream := TMemoryStream.Create;
  MemoryStream.LoadFromFile('DemoDLL.dll');
  MemoryStream.Position := 0;
  m_DllDataSize := MemoryStream.Size;
  mp_DllData := GetMemory(m_DllDataSize);
  MemoryStream.Read(mp_DllData^, m_DllDataSize);
  MemoryStream.Free;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FreeMemory(mp_DllData);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  m_DllHandle := LoadLibrary('DemoDLL.dll');
  try
    if m_DllHandle = 0 then
      Abort;
    @m_TestCallstd := GetProcAddress(m_DllHandle, 'TestCallstd');
    if @m_TestCallstd = nil then
      Abort;
    @m_TestCallcdel := GetProcAddress(m_DllHandle, 'TestCallcdel');
    if @m_TestCallcdel = nil then
      Abort;
    m_TestCallstd('This is a Dll File call!');
    m_TestCallcdel('This is a Dll File call!');
  except
    Showmessage('An error occoured while loading the dll');
  end;
  if m_DllHandle <> 0 then
    FreeLibrary(m_DllHandle)
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  mp_MemoryModule := BTMemoryLoadLibary(mp_DllData, m_DllDataSize);
  try
    if mp_MemoryModule = nil then
      Abort;
    @m_TestCallstd := BTMemoryGetProcAddress(mp_MemoryModule, 'TestCallstd');
    if @m_TestCallstd = nil then
      Abort;
    @m_TestCallcdel := BTMemoryGetProcAddress(mp_MemoryModule, 'TestCallcdel');
    if @m_TestCallcdel = nil then
      Abort;
    m_TestCallstd('This is a Dll Memory call!');
    m_TestCallcdel('This is a Dll Memory call!');
  except
    Showmessage('An error occoured while loading the dll: ' +
      BTMemoryGetLastError);
  end;
  if mp_MemoryModule <> nil then
    BTMemoryFreeLibrary(mp_MemoryModule);
end;


end.

