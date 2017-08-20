object Form1: TForm1
  Left = 373
  Top = 231
  Width = 330
  Height = 91
  Caption = 'MemoryModule Demo Application'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 24
    Top = 16
    Width = 121
    Height = 25
    Caption = 'File call'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 168
    Top = 16
    Width = 129
    Height = 25
    Caption = 'Memory Call'
    TabOrder = 1
    OnClick = Button2Click
  end
end
