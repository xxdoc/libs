library DemoDLL;

uses
  windows;

procedure TestCallstd(f_Text: PChar); stdcall;
begin
  MessageBox(0, f_Text, 'Dll Dialog (stdcall)', 0);
end;

procedure TestCallcdel(f_Text: PChar); cdecl;
begin
  MessageBox(0, f_Text, 'Dll Dialog (cdecl)', 0);
end;

exports
  TestCallstd,
  TestCallcdel;

begin
end.

