VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Created & released by KSY, 06/14/2003
'
#Const DEBUG_VER = 0

Private m_colMatches As New Collection

Public Property Get Matches() As Collection
   Set Matches = m_colMatches
End Property

Public Function ParseMenu( _
      SourceText As String, _
      Optional BeginDelimiter As String = ksBegin_VB_Menu, _
      Optional EndDelimiter As String = ksEnd & vbCrLf) As VBMenus
   
   Set ParseMenu = New VBMenus
   Set m_colMatches = ParseMenu.Items
   If InStr(1, SourceText, ksBegin_VB_Menu) <= 0 Then
      Exit Function
   End If
   
   Call Parse(SourceText, BeginDelimiter, EndDelimiter)
End Function


Public Function Parse( _
      SourceText As String, _
      Optional BeginDelimiter As String = ksBegin & vbSpace, _
      Optional EndDelimiter As String = ksEnd & vbCrLf) As Collection
  
   Dim strItemText As String
   Dim pos As Long, posBegin As Long, posEnd As Long
   Dim nLevel As Long, posChild As Long
   Dim posPopupEnd As Long, strPopupEndCheck As String
   Dim lTextLen As Long, lEndTextLen As Long
   Dim lCount As Long
   
   lTextLen = Len(SourceText)
   lCount = 0
   nLevel = 0
   pos = 1
   
   lEndTextLen = Len(SourceText) - 1
   Do Until pos >= lEndTextLen
      
      posBegin = InStr(pos, SourceText, BeginDelimiter)
      If posBegin = 0 Then
         Exit Do '>---> Loop
      End If
     
      pos = posBegin + 1
      posEnd = InStr(posBegin + Len(BeginDelimiter), SourceText, EndDelimiter)
      
      ''Debug.Assert posEnd
      If posEnd = 0 Then
         Exit Do
      End If
      strItemText = Mid$(SourceText, posBegin, posEnd - posBegin + Len(EndDelimiter))
      posChild = InStr(Len(BeginDelimiter), strItemText, BeginDelimiter)
      lCount = lCount + 1
      
      If posChild Then
         'IsParent=True
         strItemText = Mid$(strItemText, 1, posChild - 1)
         Call AddToCollection(strItemText, nLevel, True, lCount, posBegin)
         nLevel = nLevel + 1
      Else
         Call AddToCollection(strItemText, nLevel, False, lCount, posBegin)
         posPopupEnd = InStr(posEnd + Len(EndDelimiter), SourceText, BeginDelimiter)
         If posPopupEnd Then
            strPopupEndCheck = Mid$(SourceText, posEnd + Len(EndDelimiter), posPopupEnd - posEnd - Len(EndDelimiter))
            nLevel = nLevel - UBound(VBA.Split(strPopupEndCheck, EndDelimiter, , vbBinaryCompare))
         End If
      End If
DoNextLoop:
   Loop

   Set Parse = m_colMatches
End Function

Private Sub AddToCollection(Text As String, Level As Long, IsParent As Boolean, _
                                    FoundSeq As Long, Position As Long)
   Dim CtrlName As String
   Dim CtrlTypeName As String
   Dim oMenu As VBMenu
   Dim strFirstLine As String
   
   strFirstLine = Before(Text, vbCrLf, vbBinaryCompare, , True)
   If IsVBMenu(strFirstLine, CtrlName, CtrlTypeName) Then
       Set oMenu = New VBMenu
       With oMenu
          .Text = Text
          .Name = CtrlName
          .Level = Level
          .IsParent = IsParent
          .Position = Position
          .Length = Len(Text)
          Call AddCtrlProperties(oMenu, Text)
          If .Index = -1 Then
             .Key = CtrlName
          Else
             .Key = CtrlName & "(" & .Index & ")"
          End If
          'm_colMatches.Add oMenu, .Key
          m_colMatches.Add oMenu
       End With
   End If
End Sub

Private Function IsVBMenu(FirstLine As String, _
         ByRef CtrlName As String, ByRef CtrlTypeName As String) As Boolean
   Dim strTypeName As String
   CtrlName = AfterRev(RTrim$(FirstLine), vbSpace)
   CtrlTypeName = Between(FirstLine, ksBegin & vbSpace, vbSpace)
   IsVBMenu = (CtrlTypeName = ksVBMenu)
End Function

Private Sub AddCtrlProperties(ByVal oMenu As VBMenu, ByVal strPropertyText As String)
   
   ''Debug.Assert LenB(strPropertyText) > 0
   If LenB(strPropertyText) = 0 Then
      Exit Sub
   End If
   
   Dim asProp() As String, i As Long, strPropName As String, strValue As String
   asProp = Split(strPropertyText, vbCrLf)
         
   With oMenu
      For i = 0 To UBound(asProp)
         If LenB(asProp(i)) Then
            ''Debug.Print asProp(i)
            strPropName = Trim$(Before(asProp(i), "="))
            ''Debug.Assert InStr(1, asProp(i), "Caption") = 0
            If LenB(strPropName) > 0 Then
               strValue = Before(TrimCrLfTab(After(asProp(i), "=")), vbSQ, vbBinaryCompare, , True)
               ''Debug.Assert strValue <> vbQ & 0 & vbQ
               If IsNumeric(strValue) Then
                  .Properties(strPropName) = Val(strValue)
               ElseIf strValue = ksTrue Then
                  .Properties(strPropName) = True
              ElseIf strValue = ksFalse Then
                  .Properties(strPropName) = False
              Else
                  strValue = StripQuotes(strValue, True)
                  .Properties(strPropName) = strValue
               End If
               ''Debug.Print strPropName & "===" & strValue
            End If
         End If
      Next i
   End With
End Sub



