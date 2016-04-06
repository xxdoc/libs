VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4320
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   4035
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   180
      Width           =   6075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l As New clsLicense
Dim c As New clsHashBrowns


Private Sub Command1_Click()

    Dim ret() As String
    Dim usr As String
    Dim pass As String
    Dim i As Integer
    
    For i = 0 To 400
        usr = c.MD5(RandomNum())
        usr = Mid(usr, 7, 8)
        pass = l.CalcPass(usr)
        pass = StrReverse(pass)
        
        push ret(), "User: " & usr & "    Pass: " & pass
    Next
    
    Text1 = Join(ret(), vbCrLf)
    
        
       ' l.Register usr, pass
    
    
End Sub


Function RandomNum()
    Dim tmp
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub Command2_Click()
    Dim t
    t = InputBox("enter test string")
    If Len(t) = 0 Then Exit Sub
    t = Replace(t, "User: ", "")
    t = Replace(t, "    Pass: ", ":")
    t = Trim(t)
    t = Split(t, ":")
    
    l.Register t(0), t(1)
    Command2_Click
    
End Sub
