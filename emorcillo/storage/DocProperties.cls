VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' OLEDocumentProperties class
'
' This class reads and writes properties to OLE documents
' including Office documents.
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 07/31/1999
' Last Updated:
'
'        08/03/1999. * pvRead* and pvWrite* were removed.
'                      Properties now use GetPropertyByName
'                      and SetPropertyByName.
'                    * GetPropertyByName and SetPropertyByName
'                      Name parameter was changed to
'                      Variant to pass the property ID
'                      in numeric format.
'                    * Added support for strings array
'                      properties.
'                    * Added BindToStorage method
'                    * Load was renamed to BindToFile
'
'        07/31/1999. * Added User properties support.
'                    * GetPropertyByName bug (always read
'                      PIDSI_CHARCOUNT) corrected.
'
'*********************************************************************************************

Option Explicit

Dim m_Storage As olelib.IStorage
Dim m_PropSet As IPropertySetStorage
    
Enum odpErrors
    odpInvalidFile = vbObjectError
    odpFileIsClosed
    odpPropertySetDoesntExist
End Enum

Const odpFileIsCloseStr = "The file is closed."
Const odpPropertySetDoesntExistStr = "The property set doesn't exist."

Enum odpPropertySets
   odpSummary
   odpDocSummary
   odpUserProperties
End Enum

Public Property Get Application() As String
   Application = GetPropertyByName(odpSummary, PIDSI_APPNAME)
End Property

Public Property Let Application(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_APPNAME, New_Value
End Property

Public Property Get Author() As String
   Author = GetPropertyByName(odpSummary, PIDSI_AUTHOR)
End Property

Public Property Let Author(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_AUTHOR, New_Value
End Property

Public Sub BindToFile(ByVal FileName As String)
Dim lRes As Long

   ' Check if the file is a storage file
   lRes = StgIsStorageFile(FileName)
   
   If lRes <> 0 Then
      If (Err And &H80000000) <> 0 Then
          Err.Raise lRes
      Else
          Err.Raise odpInvalidFile, , FileName & " isn't a structured storage file."
      End If
   End If
   
   ' Open the storage file
   Set m_Storage = StgOpenStorage(FileName, Nothing, STGM_READWRITE Or STGM_TRANSACTED Or STGM_SHARE_DENY_WRITE, vbNullString)
   
   ' Get the IPropertySetStorage
   ' interface
   Set m_PropSet = m_Storage
        
End Sub

Public Sub BindToStorage(ByVal Stg As Object)

    ' Open the storage file
    Set m_Storage = Stg.Storage
    
    ' Get IPropertySetStorage
    Set m_PropSet = m_Storage
        
End Sub

Public Property Get Bytes() As Long
   Bytes = GetPropertyByName(odpDocSummary, PID_BYTECOUNT)
End Property

Public Property Let Bytes(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_BYTECOUNT, NewValue
End Property

Public Property Get Category() As String
   Category = GetPropertyByName(odpDocSummary, PID_CATEGORY)
End Property

Public Property Let Category(ByVal New_Value As String)
   SetPropertyByName odpDocSummary, PID_CATEGORY, New_Value
End Property

Public Property Get Chars() As Long
   Chars = GetPropertyByName(odpSummary, PIDSI_CHARCOUNT)
End Property

Public Property Let Chars(ByVal NewValue As Long)
   SetPropertyByName odpSummary, PIDSI_CHARCOUNT, NewValue
End Property

Public Sub CloseFile()

   ' Release the objects
   Set m_PropSet = Nothing
   Set m_Storage = Nothing
   
End Sub

Public Property Get Comments() As String
   Comments = GetPropertyByName(odpSummary, PIDSI_COMMENTS)
End Property

Public Property Let Comments(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_COMMENTS, New_Value
End Property

Public Property Get Company() As String
   Company = GetPropertyByName(odpDocSummary, PID_COMPANY)
End Property

Public Property Let Company(ByVal NewValue As String)
   SetPropertyByName odpDocSummary, PID_COMPANY, NewValue
End Property

Public Property Get CreationDate() As Date
   CreationDate = GetPropertyByName(odpSummary, PIDSI_CREATE_DTM)
End Property

Public Property Let CreationDate(ByVal NewValue As Date)
   SetPropertyByName odpSummary, PIDSI_CREATE_DTM, NewValue
End Property

Public Property Get EditingTime() As Date
   EditingTime = GetPropertyByName(odpSummary, PIDSI_EDITTIME)
End Property

Public Property Let EditingTime(ByVal NewValue As Date)
   SetPropertyByName odpSummary, PIDSI_EDITTIME, NewValue
End Property

Public Sub EnumProperties(ByVal PropSet As odpPropertySets, PropNames() As Variant)
Dim uSTATPS As STATPROPSTG
Dim oEnum As IEnumSTATPROPSTG
Dim oPropStorage As IPropertyStorage
Dim lCount As Long

   On Error Resume Next
   
   ' Open the property set
   Select Case PropSet
       Case odpSummary
           Set oPropStorage = m_PropSet.Open(FMTID_SummaryInformation, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
       Case odpDocSummary
           Set oPropStorage = m_PropSet.Open(FMTID_DocSummaryInformation, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
       Case odpUserProperties
           Set oPropStorage = m_PropSet.Open(FMTID_UserProperties, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
   End Select
   
   ' If the set can't be opened
   ' exit the sub
   If oPropStorage Is Nothing Then Exit Sub
   
   ' Get the enumeator object
   Set oEnum = oPropStorage.Enum

   ' Erase the array
   Erase PropNames
   
   ' Enumerate all properties
   Do While oEnum.Next(1, uSTATPS) = 0
       
      ReDim Preserve PropNames(0 To lCount)
      
      If uSTATPS.lpwstrName = 0 Then
          PropNames(lCount) = uSTATPS.propid
      Else
          PropNames(lCount) = Ptr2Str(uSTATPS.lpwstrName, True, True)
      End If
      
      ' Increment the counter
      lCount = lCount + 1
       
   Loop
   
   ' Release the objects
   Set oEnum = Nothing
   Set oPropStorage = Nothing
   
End Sub

Public Function GetPropertyByName( _
      ByVal PropSet As odpPropertySets, _
      ByVal Name As Variant) As Variant
Dim uPROP As PROPSPEC
Dim oIPS As IPropertyStorage
Dim vValue As Variant
Dim sPropName As String
Dim iVarType As Integer

   If m_PropSet Is Nothing Then Err.Raise odpFileIsClosed, , odpFileIsCloseStr
   
   On Error Resume Next
   
   ' Open the property set
   Select Case PropSet
      Case odpSummary
         Set oIPS = m_PropSet.Open(FMTID_SummaryInformation, _
                                   STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
      Case odpDocSummary
         Set oIPS = m_PropSet.Open(FMTID_DocSummaryInformation, _
                                   STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
      Case odpUserProperties
         Set oIPS = m_PropSet.Open(FMTID_UserProperties, _
                                   STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
   End Select
   
   If oIPS Is Nothing Then
      On Error GoTo 0
      Err.Raise odpPropertySetDoesntExist, , odpPropertySetDoesntExistStr
   End If
   
   ' Set property ID/Name
   If VarType(Name) = vbString Then
      
      sPropName = Name
      
      uPROP.ulKind = PRSPEC_LPWSTR
      uPROP.ID_or_LPWSTR = StrPtr(sPropName)
       
   Else
   
      uPROP.ulKind = PRSPEC_PROPID
      uPROP.ID_or_LPWSTR = CLng(Name)
       
   End If
   
   ' Get the property
   oIPS.ReadMultiple 1, uPROP, vValue

   ' Check vValue type and convert if necesary
   If Not IsEmpty(vValue) Then
       
      iVarType = VarType(vValue)
      
      Select Case iVarType
          
         Case VT_LPSTR, VT_LPWSTR
            GetPropertyByName = ToBSTR(vValue)
                      
         Case VT_FILETIME
            GetPropertyByName = ToDate(vValue)
          
         Case vbByte, vbBoolean, vbInteger, vbLong, vbSingle, _
              vbDouble, vbCurrency, vbDecimal, vbDate, vbString
               
            GetPropertyByName = vValue
                          
         Case VT_CF
          
         Case VT_VECTOR Or VT_LPSTR
            GetPropertyByName = ToBSTRArray(vValue)
         
         Case VT_VECTOR Or VT_LPWSTR
            GetPropertyByName = ToBSTRArray(vValue, True)
            
      End Select
      
   End If
   
   ' Clear variant vValue
   PropVariantClear vValue
   
   ' Release the property storage
   Set oIPS = Nothing
    
End Function

Public Property Get HiddenSlides() As Long
   HiddenSlides = GetPropertyByName(odpDocSummary, PID_HIDDENCOUNT)
End Property

Public Property Let HiddenSlides(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_HIDDENCOUNT, NewValue
End Property

Public Property Get Keywords() As String
   Keywords = GetPropertyByName(odpSummary, PIDSI_KEYWORDS)
End Property

Public Property Let Keywords(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_KEYWORDS, New_Value
End Property

Public Property Get LastSaveDate() As Date
    LastSaveDate = GetPropertyByName(odpSummary, PIDSI_LASTSAVE_DTM)
End Property

Public Property Let LastSaveDate(ByVal NewValue As Date)
    SetPropertyByName odpSummary, PIDSI_LASTSAVE_DTM, NewValue
End Property

Public Property Get LastSavedBy() As String
   LastSavedBy = GetPropertyByName(odpSummary, PIDSI_LASTAUTHOR)
End Property

Public Property Let LastSavedBy(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_LASTAUTHOR, New_Value
End Property

Public Property Get Lines() As Long
   Lines = GetPropertyByName(odpDocSummary, PID_LINECOUNT)
End Property

Public Property Let Lines(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_LINECOUNT, NewValue
End Property

Public Property Get MMClips() As Long
   MMClips = GetPropertyByName(odpDocSummary, PID_MMCLIPCOUNT)
End Property

Public Property Let MMClips(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_MMCLIPCOUNT, NewValue
End Property

Public Property Get Notes() As Long
   Notes = GetPropertyByName(odpDocSummary, PID_NOTECOUNT)
End Property

Public Property Let Notes(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_NOTECOUNT, NewValue
End Property

Public Property Get Pages() As Long
   Pages = GetPropertyByName(odpSummary, PIDSI_PAGECOUNT)
End Property

Public Property Let Pages(ByVal NewValue As Long)
   SetPropertyByName odpSummary, PIDSI_PAGECOUNT, NewValue
End Property

Public Property Get Paragraphs() As Long
   Paragraphs = GetPropertyByName(odpDocSummary, PID_PARCOUNT)
End Property

Public Property Let Paragraphs(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_PARCOUNT, NewValue
End Property

Public Property Get PresentationTarget() As String
   PresentationTarget = GetPropertyByName(odpDocSummary, PID_PRESFORMAT)
End Property

Public Property Let PresentationTarget(ByVal NewValue As String)
   SetPropertyByName odpDocSummary, PID_PRESFORMAT, NewValue
End Property

Public Property Get PrintDate() As Date
   PrintDate = GetPropertyByName(odpSummary, PIDSI_LASTPRINTED)
End Property

Public Property Let PrintDate(ByVal NewValue As Date)
   SetPropertyByName odpSummary, PIDSI_LASTPRINTED, NewValue
End Property

Public Sub Revert()
   m_Storage.Revert
End Sub

Public Property Get Revision() As String
   Revision = GetPropertyByName(odpSummary, PIDSI_REVNUMBER)
End Property

Public Property Let Revision(ByVal NewValue As String)
   SetPropertyByName odpSummary, PIDSI_REVNUMBER, NewValue
End Property

Public Sub Save()
   m_Storage.Commit STGC_DEFAULT
End Sub

Public Sub SetPropertyByName(ByVal PropSet As odpPropertySets, ByVal Name As Variant, ByVal NewValue As Variant)
Dim uPROP As PROPSPEC, Value As Variant, PropName As String
Dim IPS As IPropertyStorage, vt As Integer, CLSID As UUID

    If m_PropSet Is Nothing Then Err.Raise odpFileIsClosed, , odpFileIsCloseStr
        
    Select Case PropSet
        
        Case odpSummary
            
            Set IPS = m_PropSet.Open(FMTID_SummaryInformation, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
        
        Case odpDocSummary
            
            Set IPS = m_PropSet.Open(FMTID_DocSummaryInformation, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
        
        Case odpUserProperties
            
            On Error Resume Next

            Set IPS = m_PropSet.Open(FMTID_UserProperties, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
            
            If Err.Number <> 0 Then
                Set IPS = m_PropSet.Create(FMTID_UserProperties, CLSID, PROPSETFLAG_ANSI, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
                m_Storage.Commit STGC_DEFAULT
            End If
            
            On Error GoTo 0
    
    End Select

    If IsNumeric(Name) Then
        
        uPROP.ulKind = PRSPEC_PROPID
        uPROP.ID_or_LPWSTR = CLng(Name)
    
    Else
        
        PropName = CStr(Name)
        
        uPROP.ulKind = PRSPEC_LPWSTR
        uPROP.ID_or_LPWSTR = StrPtr(PropName)
        
    End If

    vt = VarType(NewValue)
    
    Select Case vt
        
        Case VT_BSTR
            ToLPSTR NewValue, Value
        
        Case VT_DATE
            ToFILETIME NewValue, Value
        
        Case VT_BOOL
            Value = NewValue
        
        Case Else
            If vt = VT_ARRAY Or VT_BSTR Then
                ToLPSTRArray NewValue, Value
            Else
                Value = CLng(NewValue)
            End If
            
    End Select
    
    ' Write the property
    IPS.WriteMultiple 1, uPROP, Value
    
    ' Save property storage
    IPS.Commit STGC_DEFAULT
    
    ' Clear Value
    PropVariantClear Value
    Set IPS = Nothing
    
End Sub

Public Property Get Slides() As Long
   Slides = GetPropertyByName(odpDocSummary, PID_SLIDECOUNT)
End Property

Public Property Let Slides(ByVal NewValue As Long)
   SetPropertyByName odpDocSummary, PID_SLIDECOUNT, NewValue
End Property

Public Property Get Subject() As String
   Subject = GetPropertyByName(odpSummary, PIDSI_SUBJECT)
End Property

Public Property Let Subject(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_SUBJECT, New_Value
End Property

Public Property Get Template() As String
   Template = GetPropertyByName(odpSummary, PIDSI_TEMPLATE)
End Property

Public Property Let Template(ByVal NewValue As String)
   SetPropertyByName odpSummary, PIDSI_TEMPLATE, NewValue
End Property

Public Property Get Title() As String
   Title = GetPropertyByName(odpSummary, PIDSI_TITLE)
End Property

Public Property Let Title(ByVal New_Value As String)
   SetPropertyByName odpSummary, PIDSI_TITLE, New_Value
End Property

Public Property Get Words() As Long
   Words = GetPropertyByName(odpSummary, PIDSI_WORDCOUNT)
End Property

Public Property Let Words(ByVal NewValue As Long)
   SetPropertyByName odpSummary, PIDSI_WORDCOUNT, NewValue
End Property
