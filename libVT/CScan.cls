VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'scans{} block can come first or last...
'{"permalink": "https:\/\/www.virustotal.com\/file\/856f1c37472ec67c72aaa76b03da19e50c725c9222cd9aee1360df656d0862c7\/analysis\/1303390551\/",
'"sha1": "fec37a0cf9208fdb732bd8bb2ef628480a56a22c",
'"resource": "0A2BA6D0D8C7220D0DB6D8A8F1EB5890",
'"response_code": 1,
'"scan_date": "2011-04-21 12:55:51",
'"scan_id": "856f1c37472ec67c72aaa76b03da19e50c725c9222cd9aee1360df656d0862c7-1303390551",
'"verbose_msg": "Scan finished, scan information embedded in this object",
'"sha256": "856f1c37472ec67c72aaa76b03da19e50c725c9222cd9aee1360df656d0862c7",
'"positives": 16,
'"total": 41,
'"md5": "0a2ba6d0d8c7220d0db6d8a8f1eb5890",
'"scans": {

'Example response from a submission:
'{
'    "scan_id": "e5212347fddae865d3b09df23b3701b607912d914598ed889a3d74b142b60c28-1383901546",
'    "sha1": "66f3eeceb1e6750d922c8ea786fc833d79214abe",
'    "resource": "e5212347fddae865d3b09df23b3701b607912d914598ed889a3d74b142b60c28",
'    "response_code": 1,
'    "sha256": "e5212347fddae865d3b09df23b3701b607912d914598ed889a3d74b142b60c28",
'    "permalink": "https://www.virustotal.com/file/e5212347fddae865d3b09df23b3701b607912d914598ed889a3d74b142b60c28/analysis/1383901546/",
'    "md5": "0800344e45b1aada2917a6f90df0e6e3",
'    "verbose_msg": "Scan request successfully queued, come back later for the report"
'}


Public permalink
Public sha1
Public resource
Public response_code
Public scan_date
Public scan_id
Public verbose_msg
Public sha256
Public positives
Public total
Public md5
Public first_seen

Public LocalFilePath 'not part of VT response, we add manually for UI

Public RawJson As String
Public HTTP_ResponseCode As Long
Public HTTP_StatusMessage As String

Public results As New Collection
Public HadError As Boolean
Public Errors As New Collection

Public Function GetReport() As String
    Dim r As CResult
    Dim tmp As String
    
    If response_code <> 1 Then
        If Len(LocalFilePath) > 0 Then tmp = tmp & "File:        " & FileNameFromPath(LocalFilePath) & vbCrLf & vbCrLf
        tmp = verbose_msg
    Else

        tmp = "Scan Date:   " & Me.scan_date & vbCrLf
        If Len(first_seen) > 0 Then tmp = tmp & "First Seen:  " & first_seen & vbCrLf
        
        tmp = tmp & "Hash:        " & Me.md5 & vbCrLf
        tmp = tmp & "Detections:  " & Me.positives & "/" & Me.total & vbCrLf
        
        If Len(LocalFilePath) > 0 Then tmp = tmp & "File:        " & FileNameFromPath(LocalFilePath) & vbCrLf
        
        If positives = 0 Then
            tmp = tmp & vbCrLf & "File had no detections" & vbCrLf
        Else
            tmp = tmp & String(45, "-") & vbCrLf
            For Each r In results
                If r.detected Then
                    tmp = tmp & r.PaddedScanner() & r.result & vbCrLf
                End If
            Next
        End If
            
    End If
    
    GetReport = tmp & vbCrLf
    
End Function

Public Function BriefReport() As String
    Dim tmp As String
    
    If response_code <> 1 Then
        tmp = verbose_msg
    Else
        tmp = "Scan Date:   " & Me.scan_date & vbCrLf
        If Len(first_seen) > 0 Then tmp = tmp & "First Seen:  " & first_seen & vbCrLf
        tmp = tmp & "Detections:  " & Me.positives & "/" & Me.total
    End If
    
    BriefReport = tmp
    
End Function

Public Sub VisitPage()
    On Error Resume Next
    If Len(permalink) > 0 And LCase(VBA.Left(permalink, 4)) = "http" Then
        Shell "cmd /c start " & permalink, vbHide
    End If
End Sub

Private Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

Private Function setbyname(name As Variant, value As Variant)
    
    If InStr(name, "{") > 0 Then name = Replace(name, "{", Empty)
    If InStr(name, "}") > 0 Then name = Replace(name, "}", Empty)
    name = Trim(Replace(name, """", Empty))
    value = Trim(Replace(value, """", Empty))
    value = Replace(value, "\\", Chr(5))
    value = Replace(value, "\", Empty)
    value = Replace(value, Chr(5), "\\")
    
    Select Case LCase(name)
        Case "permalink": permalink = value
        Case "sha1": sha1 = value
        Case "resource": resource = value
        Case "response_code": response_code = value
        Case "scan_date": scan_date = value
        Case "scan_id": scan_id = value
        Case "verbose_msg": verbose_msg = value
        Case "sha256": sha256 = value
        Case "positives": positives = value
        Case "total": total = value
        Case "md5": md5 = value
        Case "first_seen": first_seen = value
    End Select
    
End Function

Friend Sub ParseResponse(json As String)

    On Error Resume Next
    Dim a As Long, b As Long
    Dim main As String, name As String, value As String, scans As String
    
    json = Replace(json, vbCr, Empty)
    json = Replace(json, vbLf, Empty)
    RawJson = json
    
    a = InStr(2, json, "scans")
    b = InStr(2, json, "}},") 'end of scans...(I will not count braces)
    If a > 0 And b > a Then
        scans = Mid(json, a, b - a)
        main = Replace(json, scans, Empty)
    Else
        main = json 'probably an error message
    End If
        
    'main = Mid(json, 2, a)
   
    tmp = Split(main, ",") 'small bug to simple..verbose_msg can contains commas...complexity of fix not worth it..
    For Each entry In tmp
        entry = Trim(entry)
        If Len(entry) = 0 Then GoTo nextone
        b = InStr(1, entry, ":")
        If b < 2 Then GoTo nextone
        name = Mid(entry, 1, b - 1)
        value = Mid(entry, b + 1)
        setbyname name, value
nextone:
    Next
    
    If Right(Me.md5, 1) = "}" Then Me.md5 = Left(Me.md5, Len(Me.md5) - 1)
    
    DoEvents
    
    Dim r As CResult
    
    b = InStr(a, json, "{") + 1
    rest = Mid(json, b)
    tmp = Split(rest, "}")
    For Each entry In tmp
        entry = Trim(entry)
        If VBA.Left(entry, 1) = "," Then entry = Trim(Mid(entry, 2))
        If Len(entry) = 0 Then GoTo nextone2
        b = InStr(1, entry, "{")
        If b < 2 Then GoTo nextone2
        
        Set r = New CResult
        
        scanner = Trim(Replace(Replace(Mid(entry, 1, b - 1), ":", Empty), """", Empty))
        rest = Mid(entry, b + 1)
        r.ParseResponse scanner, rest
        results.Add r
        
nextone2:
    Next
    
    If Err.Number <> 0 Then
        HadError = True
        Errors.Add "CScan.Parse Error: " & Err.Description
    End If
    
    '"scans": {"CAT-QuickHeal": {"detected": false, "version": "11.00", "result": null, "update": "20110421"},
    


End Sub

Private Sub Class_Initialize()
    positives = 0
End Sub
