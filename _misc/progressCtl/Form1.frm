VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Lite weight progress bar demo"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLeftJustify 
      Caption         =   "left justify"
      Height          =   240
      Left            =   4500
      TabIndex        =   8
      Top             =   675
      Width           =   1320
   End
   Begin VB.CommandButton Command5 
      Caption         =   "With count"
      Height          =   465
      Left            =   6480
      TabIndex        =   7
      Top             =   990
      Width           =   1680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "With Captions"
      Height          =   465
      Left            =   4275
      TabIndex        =   6
      Top             =   990
      Width           =   1815
   End
   Begin VB.CheckBox chkVerticalLines 
      Caption         =   "Vertical lines"
      Height          =   195
      Left            =   2700
      TabIndex        =   5
      Top             =   660
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkShowPercent 
      Caption         =   "Show Percentage"
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   660
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "half"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "abort"
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "loop"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin Project1.ucProgress ucProgress1 
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   7155
      _extentx        =   12621
      _extenty        =   556
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim abort As Boolean

Private Sub chkLeftJustify_Click()
    ucProgress1.LeftJustify = (chkLeftJustify.Value = 1)
End Sub

Private Sub chkVerticalLines_Click()
    ucProgress1.FillStyle = IIf(chkVerticalLines.Value = 1, 3, 5)
End Sub

Private Sub Command1_Click()
    
    With ucProgress1
        .Max = 100
        abort = False
        
        .ShowPercent = (chkShowPercent.Value = 1)
        
        If .ShowPercent Then
            .fontSize = 8
            .AssumeMinHeight
        End If

        For i = 0 To 100
            .Value = i
            Sleep 10
            DoEvents
            If abort Then Exit For
        Next
        
        .reset
        
    End With
    
    
End Sub

Private Sub Command2_Click()
    abort = True
End Sub

Private Sub Command3_Click()

    With ucProgress1
        .Max = 100
        .fontSize = 12
        .ShowPercent = True
        .AssumeMinHeight
        .setPercent 48
    End With
    
End Sub

Private Sub Command4_Click()

     With ucProgress1
        
        .Max = 10
        .fontSize = 12
        .AssumeMinHeight
        abort = False
        
        For i = 0 To 10
            If abort Then Exit For
            .caption = "Stage " & i
            .Value = i
            
            For j = 0 To 10
                If abort Then Exit For
                Sleep 10
                DoEvents
            Next
            
        Next
     
        .reset
        
    End With
    
End Sub

Private Sub Command5_Click()

    With ucProgress1
        
        .Max = 10
        .fontSize = 12
        .AssumeMinHeight
        .ShowCount = True
        abort = False
        
        For i = 0 To 10
            If abort Then Exit For
            
            .Value = i
            
            For j = 0 To 10
                If abort Then Exit For
                Sleep 10
                DoEvents
            Next
            
        Next
     
        .reset
        
    End With
    
End Sub
