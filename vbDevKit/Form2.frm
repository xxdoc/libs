VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4200
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   4035
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   60
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Command1_Click()
    Dim c As New clsFileStream
    Dim ret() As String
    
    c.fOpen App.path & "\clsStrings.cls", otreading
    
    While Not c.EndOfFile
        tmp = Trim(c.ReadLine)
        If Len(tmp) = 0 Then GoTo nextone
        s = InStr(tmp, " ")
        If s > 0 Then r = Mid(tmp, 1, s - 1)
        
        Select Case LCase(r)
            Case "property": push ret(), tmp
            Case "function": push ret(), tmp
            Case "sub": push ret(), tmp
            Case "public": push ret(), tmp
        End Select
        
nextone:
    Wend
    
    c.fClose
    
    Text1 = Join(ret, "")
    
    
End Sub


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
