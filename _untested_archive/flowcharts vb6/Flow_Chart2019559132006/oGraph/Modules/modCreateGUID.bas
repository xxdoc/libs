Attribute VB_Name = "modCreateGUID"
' modCreateGUID - Create a Globally Unique Identifier (GUID)
' Samples:
'   {3201047B-FA1C-11D0-B3F9-004445535400}
'   {0547C3D5-FA24-11D0-B3F9-004445535400}

    Option Explicit
    DefLng A-Z

    ' The following is from Topic: Windows Conferencing API, GUID, MSDN April 1997
    ' typedef struct _GUID {
    '    unsigned long Data1;
    '    unsigned short Data2;
    '    unsigned short Data3;
    '    unsigned char Data4[8];
    '} GUID;
    '
    'Holds a globally unique identifier (GUID), which identifies a particular _
     object class and interface. This identifier is a 128-bit value.
    '
    'For more information about GUIDs, see the Remote Procedure Call (RPC) _
     documentation or the OLE Programmer's Reference.
    '
    'Use the guidgen.exe utility to generate new values.
    'See also CONFDEST, CONFGUID, CONFNOTIFY
    '© 1997 Microsoft Corporation

    Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As String * 1
    End Type
    
    Private Declare Function CoCreateGuid Lib "ole32.dll" (tGUIDStructure As GUID) As Long
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
    '

Public Function CreateGUID() As String
    Dim sGUID   As String       'store result here
    Dim tGUID   As GUID         'get into this structure
    Dim bGuid() As Byte         'get formatted string here
    Dim lRtn    As Long
    Const clLen As Long = 50
    
    If CoCreateGuid(tGUID) = 0 Then                             'use API to get the GUID
        bGuid = String(clLen, 0)
        lRtn = StringFromGUID2(tGUID, VarPtr(bGuid(0)), clLen)  'use API to format it
        If lRtn > 0 Then                                        'truncate nulls
            sGUID = Mid$(bGuid, 1, lRtn - 1)
        End If
        CreateGUID = sGUID
    End If
End Function

Public Function IsValidGUID(GUID As Variant) As Boolean
'  Function added by Larry Rebich using the DELL8500 while in Fort McDowell, AZ
'  Make an Alias, always doing it the wrong.
'  GUID input is now variant
    IsValidGUID = IsGUIDValid(GUID)
End Function

Public Function IsGUIDValid(GUID As Variant) As Boolean
'  Function created by Larry Rebich while in Grangeville, Idaho.
'  GUID input is now variant

' Samples:
'   {3201047B-FA1C-11D0-B3F9-004445535400}
'   {0547C3D5-FA24-11D0-B3F9-004445535400}

    Const sSample = "{0547C3D5-FA24-11D0-B3F9-004445535400}"
    Dim ary() As String
    Dim sTemp As String
    Dim iPos As Integer
    
    If IsNull(GUID) Then Exit Function      ' Added
    If IsEmpty(GUID) Then Exit Function     ' Added
    
    sTemp = CStr(GUID)      'convert to string
    sTemp = Trim$(sTemp)    ' Make sure no extra spaces
    
    If Len(sTemp) < (Len(sSample) - 2) Then Exit Function 'can't be less than min with out braces
    
    ' Strip off prefix, if any
    If Len(sTemp) > 0 Then
        If Right(sTemp, 1) = "}" Then       ' Prefix in this form supported
            iPos = InStr(sTemp, "{")        'in first position?
            If iPos > 1 Then
                sTemp = Mid$(sTemp, iPos)
            End If
        End If
    End If
    
    ' Add braces if none
    If Len(sTemp) = Len(sSample) - 2 Then       'maybe no braces
        If Left$(sTemp, 1) <> "{" Then
            sTemp = "{" & sTemp
        End If
        If Right$(sTemp, 1) <> "}" Then
            sTemp = sTemp & "}"
        End If
    End If
    
    If Len(sTemp) = Len(sSample) Then                           'correct length
        If Left$(sTemp, 1) = "{" And Right$(sTemp, 1) = "}" Then    'has braces
            If InStr(sTemp, "-") Then                           'Must have at least one dash
                ary() = Split(sTemp, "-")                       'right number of dashes
                If UBound(ary) = 4 Then                         'must be this
                    If Len(ary(0)) = Len("{0547C3D5") Then      'correct lengths
                        If Len(ary(1)) = 4 Then
                            If Len(ary(2)) = 4 Then
                                If Len(ary(3)) = 4 Then
                                    If Len(ary(4)) = Len("004445535400}") Then
                                        IsGUIDValid = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function CreateGUIDWithPrefix(sPrefix As String) As String

    Dim GUID As String
    
    GUID = CreateGUID()
    GUID = sPrefix & GUID
    CreateGUIDWithPrefix = GUID

End Function

Public Function LenGUID() As Integer

    LenGUID = Len(CreateGUID)
End Function




