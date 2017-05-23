VERSION 5.00
Begin VB.Form GossRESTMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GossREST server"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "GossRESTMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin GossREST.Gossamer Gossamer1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      VDir            =   "VDir"
   End
End
Attribute VB_Name = "GossRESTMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialized during Form_Load via SanitizeInit, used during Sanitize:
Private Const FROM_CHARS As String = "[]\[\]\'\;\%\_\#"
Private Const TO_CHARS As String = "[[]]\[[]\[]]\['']\\[%]\[_]\[#]"
Private FromChars() As String
Private ToChars() As String

Private CN As ADODB.Connection
Attribute CN.VB_VarHelpID = -1
Private RS As ADODB.Recordset
Private STM As ADODB.Stream
Private LogFile As Integer

Private Sub SanitizeInit()
    FromChars = Split(FROM_CHARS, "\")
    ToChars = Split(TO_CHARS, "\")
End Sub

Private Function Sanitize(ByVal Arg As String) As String
    'Attempt to "sanitize" argument Arg against SQL Injection errors.
    Dim I As Long
    
    For I = 0 To UBound(FromChars)
        Arg = Replace$(Arg, FromChars(I), ToChars(I))
    Next
    Sanitize = Arg
End Function

Private Function Query(ByVal QueryText As String) As Byte()
    Const QUERY_TEMPLATE As String = _
            "SELECT * FROM [Movies] " _
          & "WHERE " _
          & "[Title] = '$1$' OR " _
          & "[Title] LIKE '$1$' & ' %' OR " _
          & "[Title] LIKE '% ' & '$1$' OR " _
          & "[Title] LIKE '% ' & '$1$' & '%' OR " _
          & "[Title] LIKE '%' & '$1$' & ' %' OR " _
          & "[Title] LIKE '%' & '$1$' & ',%' OR " _
          & "[Title] LIKE '%' & '$1$' & '.%' OR " _
          & "[Title] LIKE '%' & '$1$' & ':%' OR " _
          & "[Initials1] = '$2$' OR " _
          & "[Initials2] = '$2$' " _
          & "ORDER BY [Title] ASC"
    Dim ActualQuery As String
    Dim Root As JNode
    Dim RecordIndex As Long
    Dim FieldIndex As Long
    
    If Len(QueryText) = 0 Then
        'Shouldn't match anything, just get an empty Recordset so
        'we have the field names for headings:
        ActualQuery = Replace$(Replace$(QUERY_TEMPLATE, _
                                        "$2$", _
                                        vbFormFeed), _
                               "$1$", _
                               vbFormFeed)
    Else
        ActualQuery = Replace$(Replace$(QUERY_TEMPLATE, _
                                        "$2$", _
                                        GossRESTDB.AsInitials(QueryText)), _
                               "$1$", _
                               Sanitize(QueryText))
    End If
    With RS
        .Open ActualQuery, , , , adCmdText
        Set Root = New JNode
        Root("RecordCount") = RS.RecordCount
        If RS.RecordCount > 0 Then
            Set Root("Rows") = New JNode
            Root("Rows").MakeArray
            Do Until .EOF
                Set Root("Rows")(RecordIndex) = New JNode
                Root("Rows")(RecordIndex).MakeObject
                For FieldIndex = 0 To .Fields.Count - 1
                    With .Fields.Item(FieldIndex)
                        Root("Rows")(RecordIndex).Value(.Name) = .Value
                    End With
                Next
                .MoveNext
                RecordIndex = RecordIndex + 1
            Loop
        End If
        .Close
    End With
    With STM
        .Open
        .Type = adTypeText
        .CharSet = "utf-8"
        .WriteText Root.JSON, adWriteChar
        .Position = 0
        .Type = adTypeBinary
        Query = .Read(adReadAll)
        .Close
    End With
End Function

Private Sub Form_Initialize()
    GossRESTDB.InitializeDB
End Sub

Private Sub Form_Load()
    SanitizeInit
    Set CN = New ADODB.Connection
    CN.Open GossRESTDB.ConnectionString
    Set RS = New ADODB.Recordset
    'Set up as a high-performance server-side cursor with the ability to
    'be data-bound to an MSHFlexGrid:
    With RS
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .ActiveConnection = CN
        .Properties("IRowsetIdentity") = True
    End With
    Set STM = New ADODB.Stream
    
    LogFile = FreeFile(0)
    Open "log.txt" For Append As #LogFile
    Gossamer1.StartListening
    
    Show
    WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Gossamer1.StopListening
    Close #LogFile
    CN.Close
End Sub

Private Sub Gossamer1_DynamicRequest( _
    ByVal Method As String, _
    ByVal URI As String, _
    ByVal Params As String, _
    ByVal ReqHeaders As Collection, _
    ByRef RespStatus As Single, _
    ByRef RespStatusText As String, _
    ByRef RespMIME As String, _
    ByRef RespExtraHeaders As String, _
    ByRef RespBody() As Byte, _
    ByVal ClientIndex As Integer)
    
    Dim ErrNumber As Long
    Dim ErrDescription As String
    
    If Method = "GET" Then
        'We'll assume URI = "/query" but we'll take any as "/query" in this program.
        On Error Resume Next
        RespBody = Query(Gossamer1.URLDecode(Params))
        If Err Then
            ErrNumber = Err.Number
            ErrDescription = Err.Description
            On Error GoTo 0
            RespStatus = 500
            RespStatusText = "Internal Server Error"
            Print #LogFile, _
                  "Error "; _
                  CStr(ErrNumber); _
                  " (&H"; Right$("0000000" & Hex$(ErrNumber), 8); ") "; _
                  ErrDescription
            Exit Sub
        End If
        On Error GoTo 0
        RespStatus = 200
        RespStatusText = "Ok"
        RespMIME = "application/json; charset=utf-8"
    Else
        RespStatus = 405
        RespStatusText = "Method Not Allowed"
        RespExtraHeaders = "Allow: GET" & vbCrLf
    End If
End Sub

Private Sub Gossamer1_LogEvent(ByVal GossEvent As GossEvent, ByVal ClientIndex As Integer)
    With GossEvent
        Print #LogFile, _
              Format$(.Timestamp, "YYYY-MM-DD HH:NN:SS, "); _
              CStr(ClientIndex); ", "; _
              .IP; ", "; _
              CStr(.EventType); ", "; _
              CStr(.EventSubtype); ", "; _
              .Method; ", "; _
              .Text
    End With
End Sub
