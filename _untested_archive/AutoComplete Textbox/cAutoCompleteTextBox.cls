VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoCompleteTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTextBox As RichTextBox
Attribute mTextBox.VB_VarHelpID = -1
Private mRS As ADODB.Recordset
Private mBannerText As String, mCueText As String
Private mSearchString As String
Private mIDForMatch As String, mTextForMatch As String
Private mIDColumnName As String, mSearchColumnName As String
Private mRestoreLastPick As Boolean 'create a property for this, if you like, and if the textbox gets focus again, it will show the last value the user chose

Event MatchSelected(MatchedID As String, MatchedText As String) 'raised upon enter key-press or focus-loss
Public Sub Init(TextBox As RichTextBox, RS As ADODB.Recordset, IDColumnName As String, SearchColumnName As String, Optional BannerText As String, Optional CueText As String)
   Set mTextBox = TextBox
   Set mRS = RS
   mIDColumnName = IDColumnName
   mSearchColumnName = SearchColumnName
   If Len(BannerText) Then mBannerText = BannerText
   If Len(CueText) Then mCueText = CueText
   mTextBox.Text = mBannerText
   RefreshTextbox
End Sub
Private Sub mTextBox_GotFocus()
   If mRestoreLastPick Then
      If mTextBox.Text = mBannerText Then mTextBox.Text = mCueText
   Else
      mTextBox.Text = mCueText
      mSearchString = vbNullString
      mIDForMatch = vbNullString
      mTextForMatch = vbNullString
   End If
   RefreshTextbox
End Sub
Private Sub mTextBox_LostFocus()
   If mTextBox.Text = mCueText Then mTextBox.Text = mBannerText
   RefreshTextbox True
   RaiseEvent MatchSelected(mIDForMatch, mTextForMatch)
End Sub
Private Sub mTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyEscape
         mSearchString = vbNullString
         FindFirst
      Case vbKeyBack
         If Len(mSearchString) > 0 Then mSearchString = Left$(mSearchString, Len(mSearchString) - 1)
         FindFirst
      Case vbKeyHome
         FindFirst
      Case vbKeyEnd
         FindLast
      Case vbKeyUp
         If Len(mIDForMatch) Then FindPrevious Else FindLast
       Case vbKeyDown
         If Len(mIDForMatch) Then FindNext Else FindFirst
      Case vbKeyLeft, vbKeyRight
         'User can't use these arrow keys - cursor must remain where we put it!
      Case vbKeyReturn
         If mIDForMatch = vbNullString Then
            Beep
         Else
            RaiseEvent MatchSelected(mIDForMatch, mTextForMatch)
         End If
      Case Else
         Exit Sub
   End Select
   KeyCode = 0
End Sub
Private Sub mTextBox_KeyPress(KeyAscii As Integer)
   If (KeyAscii >= 32 And KeyAscii <= 126) Then
      mSearchString = mSearchString & Chr$(KeyAscii)
      FindFirst
      If Len(mIDForMatch) = 0 Then
         mSearchString = Left$(mSearchString, Len(mSearchString) - 1)
         FindFirst
         Beep
      End If
   End If
   KeyAscii = 0
End Sub
Private Sub RefreshTextbox(Optional LostFocus As Boolean)
Dim AllItalic As Boolean, AllGrey As Boolean
   With mTextBox
      AllItalic = (.Text = mBannerText) Or (.Text = mCueText)
      AllGrey = AllItalic Or (Not LostFocus And Len(mIDForMatch) > 0)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelItalic = AllItalic
      .SelUnderline = False
      .SelColor = IIf(AllGrey, vbGrayText, vbWindowText)
   
      If Not LostFocus And Len(mIDForMatch) > 0 Then
         .SelStart = InStr(UCase$(mTextForMatch), UCase$(mSearchString)) - 1
         .SelLength = Len(mSearchString)
         .SelUnderline = True
         .SelColor = vbHighlight
         .SelStart = .SelStart + .SelLength
      Else
         .SelStart = 0
      End If
   End With
End Sub
Private Sub mTextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If mTextBox.Text = mCueText Then
      mTextBox.SelStart = 0
   Else
      RefreshTextbox
   End If
End Sub
Private Sub FindFirst()
   DoSearch 1, adSearchForward
End Sub
Private Sub FindLast()
   DoSearch mRS.RecordCount, adSearchBackward
End Sub
Private Sub FindNext()
   If mRS.AbsolutePosition = mRS.RecordCount Or mRS.EOF Then
      FindFirst
   Else
      If Not DoSearch(mRS.AbsolutePosition + 1, adSearchForward) Then
         FindFirst
      End If
   End If
End Sub
Private Sub FindPrevious()
   If mRS.AbsolutePosition = 1 Or mRS.BOF Then
      FindLast
   Else
      If Not DoSearch(mRS.AbsolutePosition - 1, adSearchBackward) Then
         FindLast
      End If
   End If
End Sub
Private Function DoSearch(StartRow As Long, Direction As SearchDirectionEnum) As Boolean
Dim SearchString As String, PreviousRowPos As Long

   If mRS.RecordCount = 0 Then Exit Function
   
   If Len(mSearchString) > 0 Then
      SearchString = mSearchColumnName & " LIKE '%" & mSearchString & "%'"
   Else
      SearchString = mSearchColumnName & " > ''"
   End If

   PreviousRowPos = mRS.AbsolutePosition
   mRS.AbsolutePosition = StartRow
   mRS.Find SearchString, , Direction
   
   If Not mRS.EOF And Not mRS.BOF Then
      mTextForMatch = mRS.Fields(mSearchColumnName)
      mIDForMatch = mRS.Fields(mIDColumnName)
      mTextBox.Text = mTextForMatch
      DoSearch = True
   Else
      mRS.AbsolutePosition = PreviousRowPos
      mTextForMatch = vbNullString
      mIDForMatch = vbNullString
   End If
   
   RefreshTextbox
End Function
Private Sub Class_Initialize()
   mCueText = "Type something"
   mBannerText = mCueText
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mTextBox = Nothing
   Set mRS = Nothing
End Sub
