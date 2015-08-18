VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JsonBag Testbed - IDE testing: Be sure ""Break on unhandled errors"" is set!"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetItemJSON 
      Caption         =   "Get ItemJSON"
      Height          =   495
      Left            =   1740
      TabIndex        =   13
      Top             =   7200
      Width           =   1515
   End
   Begin VB.CommandButton cmdLetItemJSON 
      Caption         =   "Let ItemJSON"
      Height          =   495
      Left            =   60
      TabIndex        =   12
      Top             =   7200
      Width           =   1515
   End
   Begin VB.CommandButton cmdCloneJsonBag 
      Caption         =   "Clone JsonBag"
      Height          =   495
      Left            =   8460
      TabIndex        =   11
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CommandButton cmdExtractJBClone 
      Caption         =   "Extract JB Clone"
      Height          =   495
      Left            =   6780
      TabIndex        =   10
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CommandButton cmdInsertJBClone 
      Caption         =   "Insert JB Clone"
      Height          =   495
      Left            =   5100
      TabIndex        =   9
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CheckBox chkDecimalMode 
      Caption         =   "Decimal Mode"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenSerialize 
      Caption         =   "Gen && Serialize"
      Height          =   495
      Left            =   3420
      TabIndex        =   8
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CheckBox chkWhiteSpace 
      Caption         =   "Add Whitespace"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCutPasteBack 
      Caption         =   "Cut/Paste Back"
      Height          =   495
      Left            =   1740
      TabIndex        =   3
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CommandButton cmdParseReser 
      Caption         =   "Parse/Reser"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   6600
      Width           =   1515
   End
   Begin VB.TextBox txtSerialized 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   420
      Width           =   4515
   End
   Begin VB.TextBox txtOriginal 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   420
      Width           =   4455
   End
   Begin VB.Label lblDeserReser 
      Caption         =   "Parsed/Reserialized, other resutls"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label lblOriginal 
      Caption         =   "Original"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GapHorizontal As Single
Private GapVertical As Single

Private JB As JsonBag

Private Sub chkDecimalMode_Click()
    JB.DecimalMode = chkDecimalMode.Value = vbChecked
End Sub

Private Sub chkWhiteSpace_Click()
    JB.Whitespace = chkWhiteSpace.Value = vbChecked
End Sub

Private Sub cmdCloneJsonBag_Click()
    Dim CloneJB As JsonBag
    
    With JB
        .Clear
        .IsArray = True
        .Item = "Hello"
        .Item = Null
        .Item = "World"
        txtOriginal.Text = .JSON
        
        Set CloneJB = .Clone
        
        .Item(2) = "oops!" 'If we see this we failed to clone properly!
    End With
    
    txtSerialized.Text = CloneJB.JSON
End Sub

Private Sub cmdCutPasteBack_Click()
    txtOriginal.Text = txtSerialized.Text
    txtSerialized.Text = vbNullString
    
    cmdParseReser.SetFocus
End Sub

Private Sub cmdExtractJBClone_Click()
    Dim ExtractedJB1 As JsonBag
    Dim ExtractedJB2 As JsonBag
    
    With JB
        .Clear
        .IsArray = True
        .Item = "Hello"
        With .AddNewObject()
            .Item("x") = 13
            .Item("y") = 14
            With .AddNewArray("Zzz")
                .Item = 100
                .Item = 200
                .Item = 300
            End With
        End With
        .Item = "World"
        txtOriginal.Text = .JSON
        
        Set ExtractedJB1 = .CloneItem(2)
        Set ExtractedJB2 = .Item(2).CloneItem("Zzz")
        
        .Item(2)("y") = "oops!" 'If we see this we failed to clone properly!
        .Item(2)("Zzz")(3) = "oops!" 'If we see this we failed to clone properly!
    End With
    
    txtSerialized.Text = "ExtractedJB1..." & vbNewLine _
                       & ExtractedJB1.JSON & vbNewLine & vbNewLine _
                       & "ExtractedJB2..." & vbNewLine _
                       & ExtractedJB2.JSON
End Sub

Private Sub cmdGenSerialize_Click()
    With JB
        .Clear
        .IsArray = False 'Actually the default after Clear.
        
        ![First] = 1
        ![Second] = Null
        With .AddNewArray("Third")
            .Item = "These"
            .Item = "Add"
            .Item = "One"
            .Item = "After"
            .Item = "The"
            .Item = "Next"
            .Item(1) = "*These*" 'Should overwrite 1st Item, without moving it.
            
            'Add a JSON "object" to this "array" (thus no name supplied):
            With .AddNewObject()
                .Item("A") = True
                !B = False
                !C = 3.14E+16
            End With
        End With
        With .AddNewObject("Fourth")
            .Item("Force Case") = 1 'Use quoted String form to force case of names.
            .Item("force Case") = 2
            .Item("force case") = 3
            
            'This syntax can be risky with case-sensitive JSON since the text is
            'treated like any other VB identifier, i.e. if such a symbol ("Force"
            'or "Case" here) is already defined in the language (VB) or in your
            'code the casing of that symbol will be enforced by the IDE:
            
            ![Force Case] = 666 'Should overwrite matching-case named item, which
                                'also moves it to the end.
            'Safer:
            .Item("Force Case") = 666
        End With
        'Can also use implied (default) property:
        JB("Fifth") = Null
        
        txtSerialized.Text = .JSON
    End With
End Sub

Private Sub cmdGetItemJSON_Click()
    With JB
        .Clear
        .ItemJSON("This") = "{""a"":42,""b"":true}"
        .Item("That") = "A string"
        .ItemJSON("The other thing") = "[1,2,3]"
        With .AddNewArray("More")
            .ItemJSON = "[true,false,null,""\u0020""]"
        End With
        .Item("That's all") = 999
        
        txtOriginal.Text = .JSON
        
        txtSerialized.Text = ".Item(4).ItemIsJSON(1) = " _
                           & CStr(.Item(4).ItemIsJSON(1)) _
                           & vbNewLine & vbNewLine _
                           & ".Item(4).ItemJSON(1) returned:" _
                           & vbNewLine _
                           & .Item(4).ItemJSON(1) _
                           & vbNewLine & vbNewLine _
                           & ".ItemIsJSON(""That"") = " _
                           & CStr(.ItemIsJSON("That")) _
                           & vbNewLine & vbNewLine _
                           & ".ItemIsJSON(""This"") = " _
                           & CStr(.ItemIsJSON("This")) _
                           & vbNewLine & vbNewLine _
                           & ".ItemJSON(""This"") returned:" _
                           & vbNewLine _
                           & .ItemJSON("This")
    End With
End Sub

Private Sub cmdInsertJBClone_Click()
    Dim NewJB As JsonBag
    
    With JB
        .Clear
        .IsArray = True
        .Item = "Hello"
        .Item = Null
        .Item = "World"
        txtOriginal.Text = .JSON
    End With
    
    Set NewJB = New JsonBag
    With NewJB
        .IsArray = True
        .Item(1) = 1
        .Item(2) = True
        .Item(3) = "3"
    End With
    
    JB.CloneItem(2) = NewJB 'This replaces original Item in JB by a clone of NewJB.
    
    Set NewJB = New JsonBag
    With NewJB
        .Item("a") = Timer()
        .Item("b") = Format$(Now(), "YYYYMMDDHHNNSS")
        .Item("c") = 0
    End With
    
    JB.CloneItem = NewJB 'This adds a new Item at the end of JB by cloning NewJB.
    
    NewJB.Item("c") = "oops!" 'If we see this we failed to clone proeprly!
    
    txtSerialized.Text = JB.JSON
End Sub

Private Sub cmdLetItemJSON_Click()
    With JB
        .Clear
        .ItemJSON("This") = "{""a"":42,""b"":true}"
        .Item("That") = "A string"
        .ItemJSON("The other thing") = "[1,2,3,4,5]"
        .ItemJSON("More") = "{""Less"":[true,false,null,""\u0020""]}"
        
        txtOriginal.Text = .JSON
    End With
    
    txtSerialized.Text = vbNullString
End Sub

Private Sub cmdParseReser_Click()
    Dim NewJB As JsonBag
    Dim NewIndex As Long
    Dim ErrDesc As String
    
    'Parse JSON text into JsonBag:
    On Error Resume Next
    JB.JSON = txtOriginal.Text
    If Err Then
        MsgBox "Error " & CStr(Err.Number) & vbNewLine & vbNewLine _
             & Err.Description
        ErrDesc = Err.Description
        On Error GoTo 0
        ErrDesc = Mid$(ErrDesc, InStrRev(ErrDesc, " ") + 1)
        If IsNumeric(ErrDesc) Then
            txtOriginal.SelStart = CLng(ErrDesc) - 1
            txtOriginal.SelLength = 1
            txtOriginal.SetFocus
        End If
        Exit Sub
    End If
    On Error GoTo 0
    
    txtSerialized.Text = JB.JSON
    
    cmdCutPasteBack.SetFocus
End Sub

Private Sub Form_Load()
    GapHorizontal = lblOriginal.Left
    GapVertical = lblOriginal.Top

    Set JB = New JsonBag
    JB.Whitespace = chkWhiteSpace.Value = vbChecked
    JB.WhitespaceIndent = 2
    JB.DecimalMode = chkDecimalMode.Value = vbChecked
End Sub

Private Sub Form_Resize()
    Dim TxtWidth As Single
    Dim TxtHeight As Single
    Dim BtnsLeft1 As Single
    Dim BtnsLeft2 As Single
    
    If WindowState <> vbMinimized Then
        TxtWidth = (ScaleWidth - 3# * GapHorizontal) / 2#
        TxtHeight = ScaleHeight - (5# * GapVertical _
                                     + lblOriginal.Height _
                                     + chkWhiteSpace.Height _
                                     + cmdParseReser.Height _
                                     + cmdLetItemJSON.Height)
        BtnsLeft1 = (ScaleWidth - (cmdParseReser.Width _
                                 + cmdCutPasteBack.Width _
                                 + cmdGenSerialize.Width _
                                 + cmdInsertJBClone.Width _
                                 + cmdExtractJBClone.Width _
                                 + cmdCloneJsonBag.Width _
                                 + 6# * GapHorizontal)) / 2#
        BtnsLeft2 = (ScaleWidth - (cmdLetItemJSON.Width _
                                 + cmdGetItemJSON.Width _
                                 + 2# * GapHorizontal)) / 2#
        With txtOriginal
            .Move GapHorizontal, lblOriginal.Height + GapVertical, TxtWidth, TxtHeight
            txtSerialized.Move 2# * GapHorizontal + .Width, .Top, TxtWidth, TxtHeight
            lblDeserReser.Left = txtSerialized.Left
            chkWhiteSpace.Move BtnsLeft1, .Top + .Height + GapVertical
        End With
        With chkWhiteSpace
            chkDecimalMode.Move .Left + .Width + GapHorizontal, .Top
        End With
        
        With chkDecimalMode
            cmdParseReser.Move BtnsLeft1, .Top + .Height + GapVertical
        End With
        With cmdParseReser
            cmdCutPasteBack.Move .Left + .Width + GapHorizontal, .Top
        End With
        With cmdCutPasteBack
            cmdGenSerialize.Move .Left + .Width + GapHorizontal, .Top
        End With
        With cmdGenSerialize
            cmdInsertJBClone.Move .Left + .Width + GapHorizontal, .Top
        End With
        With cmdInsertJBClone
            cmdExtractJBClone.Move .Left + .Width + GapHorizontal, .Top
        End With
        With cmdExtractJBClone
            cmdCloneJsonBag.Move .Left + .Width + GapHorizontal, .Top
        End With
        
        With cmdParseReser
            cmdLetItemJSON.Move BtnsLeft2, .Top + .Height + GapVertical
        End With
        With cmdLetItemJSON
            cmdGetItemJSON.Move .Left + .Width + GapHorizontal, .Top
        End With
    End If
End Sub
