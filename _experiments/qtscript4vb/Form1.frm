VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QTScript Demo w/Debugger"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      Height          =   375
      Left            =   7965
      TabIndex        =   15
      Top             =   7335
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   810
      TabIndex        =   13
      Top             =   7290
      Width           =   6900
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   900
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   1890
      Width           =   6855
   End
   Begin VB.CheckBox chkSetTimeout 
      Caption         =   "Script Timeout"
      Height          =   285
      Left            =   3195
      TabIndex        =   11
      Top             =   270
      Width           =   1410
   End
   Begin VB.TextBox txtFile 
      Height          =   330
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Text            =   "test.js"
      Top             =   990
      Width           =   4605
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   240
      Left            =   4995
      TabIndex        =   6
      Top             =   315
      Width           =   960
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "4000"
      Top             =   270
      Width           =   555
   End
   Begin VB.CheckBox chkDebugger 
      Caption         =   "Debugger"
      Height          =   285
      Left            =   315
      TabIndex        =   4
      Top             =   270
      Width           =   1680
   End
   Begin VB.TextBox txtOutput 
      Height          =   2535
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4230
      Width           =   6900
   End
   Begin VB.ComboBox cbo1 
      Height          =   315
      Left            =   900
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1530
      Width           =   4605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "eval"
      Height          =   330
      Left            =   7920
      TabIndex        =   1
      Top             =   1935
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AddFile"
      Height          =   330
      Left            =   5625
      TabIndex        =   0
      Top             =   990
      Width           =   1320
   End
   Begin VB.Label Label4 
      Caption         =   "Test COM Object "
      Height          =   240
      Left            =   90
      TabIndex        =   14
      Top             =   6885
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "Output"
      Height          =   375
      Left            =   90
      TabIndex        =   10
      Top             =   4590
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Script Text"
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   1575
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Add File"
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE:
'  to run this we assume you have QT4.8 installed
'  AND that its \bin directory has been added to your PATH envirnoment variable

'the following are for VS2008 development
'
' https://download.qt.io/archive/qt/4.8/4.8.5/
'   qt-win-opensource-4.8.5-vs2008.exe = 234 MB
'
'to compile the dll with VS2008 use the following:
' https://download.qt.io/official_releases/vsaddin/
'   qt-vs-addin-1.1.11-opensource.exe
'
'the runtime distribution requirements are:
'
'release mode: 12.2mb
'  QtCore4.dll
'  QtGui4.dll
'  QtScript4.dll
'  QtScriptTools4.dll
'
'debug mode dll: 26mb
'  QtScriptToolsd4.dll
'  QtCored4.dll
'  QtGuid4.dll
'  QtScriptd4.dll
'
'if we use a scriptAgent instead of the full fledged script debugger ui the release mode
'dependancy is down to just QtScript4,QtScriptTools4 1.8mb, i will experiment with that in
'another project..
'
'
'script console examples:
'  https://searchcode.com/codesearch/view/31490215/

Dim hlib As Long

Private Sub cbo1_Click()
    txtScript = Replace(Replace(cbo1.Text, "\n", vbCrLf), "\t", vbTab)
End Sub

Private Sub chkDebugger_Click()
     QtOp op_setdbg, chkDebugger.Value
End Sub

Private Sub chkSetTimeout_Click()
    On Error Resume Next
    Dim i As Long
    
    If chkSetTimeout.Value = 1 Then
        i = CLng(txtTimeout)
        If Err.Number <> 0 Then
            MsgBox "Invalid timeout number could not convert to long value"
            Exit Sub
        End If
    Else
        i = -1
    End If
    
    QtOp op_setTimeout, i
    
End Sub

Private Sub cmdReset_Click()
    QtOp op_reset
End Sub

Private Sub Command1_Click()
    Dim pth As String
    
    If FileExists(txtFile) Then
        pth = txtFile
    ElseIf FileExists(App.path & "\" & txtFile) Then
        pth = App.path & "\" & txtFile
    Else
        MsgBox "Path not found: " & txtFile
        Exit Sub
    End If
       
    If Not AddFile(pth) Then
        MsgBox "AddFile Error"
    End If
    
End Sub

Private Sub Command2_Click()
    Dim outVal As Variant

    txtOutput = Empty
    
    If Not Eval(txtScript.Text, outVal) Then
        txtOutput = "Eval failed!"
        Exit Sub
    End If
    
    On Error Resume Next
    If IsArray(outVal) Then
        txtOutput = TypeName(outVal) & ": " & Join(outVal, ", ") 'join not safe to use with any isMissing args but this is a quick demo so..
        If Err.Number <> 0 Then
            txtOutput = "Error joining: " & TypeName(outVal) & ": Ubound = " & UBound(outVal) & " isMissing(0) ? " & IsMissing(outVal(LBound(outVal)))
        End If
    Else
        txtOutput = TypeName(outVal) & ": " & outVal
    End If
    
End Sub

Private Sub Command3_Click()
    List1.Clear
End Sub

Private Sub Form_Load()
    
    Dim pth As String
    Dim dll As String
    
    dll = App.path & "\debug\dbg.dll"
    If Not FileExists(dll) Then
        dll = App.path & "\dbg.dll"
        If Not FileExists(dll) Then
            MsgBox "dbg.dll not found: " & dll
            End
        End If
    End If
    
    hlib = LoadLibrary(dll)
       
    If hlib = 0 Then
        MsgBox "Did you remember to add the qt4.8\bin directory to your path?" & vbCrLf & vbCrLf & _
               "LoadLib(dbg.dll) failed (if you just added it restart ide)" & vbCrLf & _
               "lastdllError: " & GetSystemErrorMessageText(Err.LastDllError) & vbCrLf, vbExclamation
        End
    End If
    
    'dont use any dbg.dll imports until after LoadLibrary to be sure it was found..
    '(since in different dir during development)
    QtOp op_setResolverHandler, AddressOf HostResolver
    
    cbo1.AddItem "1+2"
    cbo1.AddItem "'test'+' string'"
    cbo1.AddItem "function x(y){return y}x('fart')"
    cbo1.AddItem "[1,2,3,4,5]"
    cbo1.AddItem "new Array(5)" 'uninitilized values
    cbo1.AddItem "new Array(1 ,2,3,4,'test')" 'if any element is empty join will fail..
    cbo1.AddItem "v = nativeToUpper('test');v"
    cbo1.AddItem "while(1){} //danger use timeout.."
    cbo1.AddItem "for(i=0;i<10;i++)resolver('list1.additem','test_'+i);"
    cbo1.AddItem "alert(resolver('txtFile.text'))"
    cbo1.ListIndex = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    QtOp op_qtShutdown
    FreeLibrary hlib
End Sub



Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtFile = Data.Files(1)
End Sub
