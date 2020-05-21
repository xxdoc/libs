VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private Const LB_GETCOUNT = &H18B
Private Const LB_SETTOPINDEX = &H197
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Function lbCount(hListBox As Long) As Long
   lbCount = SendMessage(hListBox, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
End Function

Public Sub lbScrollTo(hListBox As Long, Optional ByVal index As Long = -1)
    If index = -1 Then index = lbCount(hListBox)
    SendMessage hListBox, LB_SETTOPINDEX, ByVal index - 1, ByVal 0
End Sub

Function lbCopy(lstBox As Object) As String
    
    Dim i As Long
    Dim tmp() As String
    
    For i = 0 To lstBox.ListCount
        push tmp, lstBox.List(i)
    Next
    
    lbCopy = Join(tmp, vbCrLf)
    
End Function

Sub lbAppend(lstBox As Object, txt, Optional ByVal index = -1)
    On Error GoTo hell
    If index = -1 Then index = lstBox.ListCount - 1
    If index < 0 Then Exit Sub
    lstBox.List(index) = lstBox.List(index) & txt
hell:
End Sub

Sub lbUpdate(lstBox As Object, txt, Optional ByVal index = -1)
    On Error GoTo hell
    If index = -1 Then index = lstBox.ListCount - 1
    If index < 0 Then Exit Sub
    lstBox.List(index) = txt
hell:
End Sub


Public Sub lvColumnSort(ListViewControl As Object, Column As Object)
    On Error Resume Next
    Const lvwAscending As Long = 0
    Const lvwDescending As Long = 1
     
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .sorted = -1
    End With
End Sub

Function lvGetAllElements(lv As Object) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next
    
    push ret, tmp
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next
    
    lvGetAllElements = Join(ret, vbCrLf)
    
End Function

Function lvGetColumn(lv As Object, colIndex As Long) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    If colIndex > lv.ColumnHeaders.count Then Exit Function
    If colIndex < 0 Then Exit Function
    
    'push ret, lv.ColumnHeaders(colIndex + 1).Text
    'push ret, String(50, "-")
        
    For Each li In lv.ListItems
        If colIndex = 0 Then
            push ret, li.Text
        Else
            push ret, li.SubItems(colIndex)
        End If
    Next
    
    lvGetColumn = Join(ret, vbCrLf)
    
End Function

Function lvSelCount(lv As Object) As Long
    On Error Resume Next
    Dim i As Long, li As Object
    For Each li In lv.ListItems
        If li.selected Then
            If Err.Number = 0 Then i = i + 1
        End If
    Next
    lvSelCount = i
End Function

Function cboFromFile(cbo As Object, fpath) As Long
    Dim fs As New CFileStream
    On Error GoTo hell
    
    cbo.Clear
    fs.Open_ fpath, otRead
    While Not fs.eof
        cbo.AddItem fs.ReadLine
    Wend
    
hell:
    On Error Resume Next
    fs.Close_
    cboFromFile = cbo.ListCount
    
End Function


Public Function fgFill(FlexGrid As Object, rs As Object) As Long
'*******************************************************
'PURPOSE: Populate MSFlexGrid with data from an
'         ADO Recordset
'PARAMETERS:  FlexGrid: MsFlexGrid to Populate
'             RS: Open ADO Recordset
'RETURNS:     True if successful, false otherwise
'REQUIRES:    -- Reference to Microsoft Active Data Objects
'             -- Recordset should be open with cursor set at
'                first row when passed and must
'                support recordcount property
'             -- FlexGrid should be empty when passed
'EXAMPLE:
'Dim conn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'Dim sConnString As String
'
'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDatabase.mdb"
'conn.Open sConnString
'rs.Open " SELECT * FROM MyTable", oConn, adOpenKeyset, adLockOptimistic
'PopulateFlexGrid MSFlexGrid1, rs
'
'rs.Close
'conn.Close
'***********************************************************
'**********************************************
            
    On Error GoTo ErrorHandler

    If TypeName(FlexGrid) <> "MSFlexGrid" Then Exit Function
    If TypeName(rs) <> "Recordset" Then Exit Function
    'If Not TypeOf rs Is ADODB.Recordset Then Exit Function
    
    Dim i As Integer
    Dim J As Integer
    
    initFlexGrid FlexGrid
    
    With FlexGrid
  
        If Not rs.eof Then
        
            '.Rows = rs.RecordCount + 1
            .Cols = rs.fields.count
            .AddItem ""
            
            For i = 0 To rs.fields.count - 1
                .TextMatrix(0, i) = rs.fields(i).Name
            Next
        
            i = 1
            Do While Not rs.eof
            
                .AddItem ""
                For J = 0 To rs.fields.count - 1
                    If Not IsNull(rs.fields(J).value) Then
                        .TextMatrix(i, J) = trim(rs.fields(J).value)
                    End If
                Next
        
                i = i + 1
                rs.MoveNext
            Loop
        
        
        End If
        fgFill = .rows
    End With
    
    Exit Function
ErrorHandler:
    fgFill = -1
    
End Function

Private Function initFlexGrid(fg As Object)
    
    On Error Resume Next
    Dim i As Long
    
    With fg
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        For i = .rows - 1 To 0 Step -1
            .RemoveItem i
        Next
    End With
    
End Function

'copies all the data from a flexgrid into well formatted rows for text display
'use delimiter to skip padding and have in csv type format,
'delimiter not checked for in data
Public Sub fgCopy(FlexGrid As Object, Optional delimiter)

   Dim colWidths() As Long
   Dim rows() As String
   Dim fields() As String
   Dim d As String
   Dim c As Long, r As Long, maxwidth As Long
   
   If TypeName(FlexGrid) <> "MSFlexGrid" Then Exit Sub
   
   With FlexGrid
   
        ReDim colWidths(.Cols)
        
        If Len(delimiter) = 0 Then
            For c = 0 To .Cols - 1
                For r = 0 To .rows - 1
                    maxwidth = Len(.TextMatrix(r, c))
                    If maxwidth > colWidths(c) Then colWidths(c) = maxwidth
                Next
            Next
        End If
        
        For r = 0 To .rows - 1
            Erase fields
            For c = 0 To .Cols - 1
                d = .TextMatrix(r, c)
                If Len(delimiter) = 0 Then
                    d = d & Space(colWidths(c) - Len(d) + 2)
                Else
                    d = d & delimiter
                End If
                push fields, d
            Next
            push rows, Join(fields, "")
        Next
    End With
    
    Clipboard.Clear
    Clipboard.SetText Join(rows, vbCrLf)
   
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
       
    x = UBound(ary)
    ReDim Preserve ary(x + 1)
    
    If IsObject(value) Then
        Set ary(x + 1) = value
    Else
        ary(x + 1) = value
    End If
    
    Exit Sub
init:
    ReDim ary(0)
    If IsObject(value) Then
        Set ary(0) = value
    Else
        ary(0) = value
    End If
End Sub
