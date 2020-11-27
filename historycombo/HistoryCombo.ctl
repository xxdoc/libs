VERSION 5.00
Begin VB.UserControl HistoryCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "HistoryCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim history() As String
Dim m_file As String

Event KeyPress(KeyAscii As Integer)

Private Const CB_LIMITTEXT = &H141
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub UserControl_Initialize()
    'usual limit is 255 on paste?!, 2800 on programatic set .text still has limits though
    'paste might not even work on a 50kb post
    SendMessage cbo.hwnd, CB_LIMITTEXT, 0, ByVal 0&
End Sub

Sub SetFont(name As String, size As Long)
    On Error Resume Next
    cbo.Font.name = name
    cbo.Font.size = size
End Sub

Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = cbo.Text
End Property

Property Let Text(ByVal v As String)
    cbo.Text = v
End Property

Property Get Combo() As ComboBox
    Set Combo = cbo
End Property

Sub ClearHistory()
    On Error Resume Next
    If FileExists(m_file) Then Kill m_file
    cbo.Clear
    Erase history
End Sub

Sub LoadHistory(localFile As String)

    On Error Resume Next
    Dim entries() As String, x, f As String
    
    cbo.Clear
    m_file = localFile
    If Not FileExists(m_file) Then Exit Sub
    
    f = ReadFile(m_file)
    entries = Split(f, vbCrLf)
    For i = 0 To UBound(entries)
        x = entries(i)
        If Len(x) > 0 Then
            push history, x
            cbo.AddItem x
        End If
    Next
    
    cbo.Text = cbo.List(0) 'last one used
    
    
End Sub

'add new enties. if its an old entry, we move it to the top of the list..
Function RecordIfNew()

    On Error Resume Next
    Dim i, entries() As String, arg As String
    
    arg = cbo.Text
    If Len(arg) = 0 Then Exit Function
    
    If Not AryIsEmpty(history) Then
        If history(0) = arg Then Exit Function 'same command as last used..
    End If
    
    push entries, arg 'add new entry to beginning
    
    If Not AryIsEmpty(history) Then
        For i = 0 To UBound(history)
             If history(i) <> arg Then push entries, history(i)
        Next
    End If
    
    history = entries
    cbo.Clear
    
    For i = 0 To UBound(history)
         cbo.AddItem history(i)
    Next
    
    cbo.Text = arg
    
End Function

Function SaveHistory(Optional maxEntries As Long = 50)

    On Error Resume Next
    Dim f As String, entries() As String

    If AryIsEmpty(history) Then Exit Function
    
    If UBound(history) > maxEntries Then
        For i = 0 To maxEntries
            push entries, history(i)
        Next
        history = entries
    End If
    
    writeFile m_file, Join(history, vbCrLf)
    
End Function

Private Sub cbo_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    cbo.Width = UserControl.Width
End Sub

Private Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function


Private Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  f = FreeFile
  Dim b() As Byte
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Private Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error Resume Next
    Dim b() As Byte
    If FileExists(path) Then Kill path
    Err.Clear
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = IIf(Err.Number = 0, True, False)
End Function

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

