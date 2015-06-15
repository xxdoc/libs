VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   3780
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   2475
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   510
      Left            =   810
      TabIndex        =   0
      Top             =   1530
      Width           =   2850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long


Dim mRGN As Long, R As RECT, x As Long, y As Long

Dim hr As Long

'must set for scalemode as pixels..
'An application moves a region by calling the OffsetRgn function.
'The given offsets along the x-axis and y-axis determine the number of logical units to move left or right and up or down.

Private Sub Form_Load()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    
    Rectangle p.hdc, 20, 20, 50, 50
    
    
    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Set the rectangle's values
    SetRect R, 20, 20, 50, 50
    'Create an elliptical region
    mRGN = CreateEllipticRgnIndirect(R)
    For x = R.Left To R.Right
        For y = R.Top To R.Bottom
            'If the point is in the region, draw a green pixel
            If PtInRegion(mRGN, x, y) <> 0 Then
                'Draw a green pixel
                SetPixelV Me.hdc, x, y, vbGreen
            ElseIf PtInRect(R, x, y) <> 0 Then
                'Draw a red pixel
                SetPixelV Me.hdc, x, y, vbRed
            End If
        Next y
    Next x
    'delete our region
    'DeleteObject mRGN
    
    'MsgBox Me.HasDC
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If PtInRegion(mRGN, x, y) <> 0 Then
        Me.Caption = "in"
        'InvertRgn Me.hdc, mRGN
    Else
        Me.Caption = "out"
        'InvertRgn Me.hdc, mRGN
    End If
        
    'If PtInRect(R, X, Y) <> 0 Then Me.Caption = "in" Else Me.Caption = "out"
    Label1.Caption = x & " " & y
End Sub
