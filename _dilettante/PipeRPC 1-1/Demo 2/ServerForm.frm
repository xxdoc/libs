VERSION 5.00
Begin VB.Form ServerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculation Server"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "ServerForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin Server.PipeRPC pipeCalculate 
      Left            =   1020
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      MaxRequest      =   200
      MaxResponse     =   200
      PipeName        =   "Calc Server Pipe"
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pipeCalculate.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pipeCalculate.ClosePipe
End Sub

Private Sub pipeCalculate_Called(ByVal Pipe As Long, Request() As Byte, Response() As Byte)
    Dim ReqParts() As String
    
    ReqParts = Split(Request, "|")
    On Error Resume Next
    Select Case ReqParts(1)
        Case "+"
            Response = CStr(CDbl(ReqParts(0)) + CDbl(ReqParts(2)))
        Case "-"
            Response = CStr(CDbl(ReqParts(0)) - CDbl(ReqParts(2)))
        Case "*"
            Response = CStr(CDbl(ReqParts(0)) * CDbl(ReqParts(2)))
        Case "÷"
            Response = CStr(CDbl(ReqParts(0)) / CDbl(ReqParts(2)))
    End Select
    If Err Then Response = Err.Description
End Sub
