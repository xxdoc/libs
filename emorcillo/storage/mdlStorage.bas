Attribute VB_Name = "mdlStorage"
'*********************************************************************************************
'
' DocumentProperties/Storage
'
' Support functions and declarations module
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http:'www.domaindlx.com/e_morcillo
'
' Created: 07/31/1999
' Updates:
'           08/12/1999. The comments were revised and enhaced.
'           08/02/1999. IsValidVariant was removed.
'           12/13/1999. Added 1 parameter to CreateFileStorage
'           02/17/2000. Added ErrorMessage function.
'*********************************************************************************************

Option Explicit

Public FMTID_SummaryInformation As olelib.UUID
Public FMTID_DocSummaryInformation As olelib.UUID
Public FMTID_UserProperties As olelib.UUID
Public IID_IUnknown As UUID
Public IID_Null As UUID

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Function ErrorMessage(ByVal ErrNum As Long) As String
Dim MessageLen As Long

    ErrorMessage = String(256, 0)
    
    MessageLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_MAX_WIDTH_MASK, 0, ErrNum, 0&, ErrorMessage, Len(ErrorMessage), ByVal 0)
    
    ErrorMessage = Left$(ErrorMessage, MessageLen)
    
End Function

'*********************************************************************************************
'
' Initializes FMTIDs and IIDs
'
'*********************************************************************************************
Public Sub Main()

    olelib.CLSIDFromString FMTIDSTR_SummaryInformation, FMTID_SummaryInformation
    olelib.CLSIDFromString FMTIDSTR_DocSummaryInformation, FMTID_DocSummaryInformation
    olelib.CLSIDFromString FMTIDSTR_UserProperties, FMTID_UserProperties
    
    olelib.CLSIDFromString IIDSTR_IUnknown, IID_IUnknown

End Sub

'*********************************************************************************************
'
' Returns a String from a LPxSTR pointer
'
' Parameters:
'
' Ptr: pointer to the string
' FreeSource: If True the source string pointer if freed.
' Unicode: Indicates if the source string is Unicode or ANSI. Default is ANSI.
'
'*********************************************************************************************
Public Function Ptr2Str(Ptr As Long, Optional FreeSource As Boolean, Optional ByVal Unicode As Boolean) As String

    If Unicode Then
        
        ' The string is Unicode
        
        ' Create a BSTR from the pointer
        Ptr2Str = SysAllocString(Ptr)
    
    Else
        
        ' Get string length to initialize
        ' the string.
        Ptr2Str = String$(lstrlenA(Ptr), 0)
        
        ' Copy the string
        lstrcpyA Ptr2Str, ByVal Ptr
        
    End If

    If FreeSource Then CoTaskMemFree Ptr: Ptr = 0
    
End Function
'*********************************************************************************************
'
' Converts a LPSTR or LPWSTR variant to String.
'
' Parameters:
'
' Var: source variant
'
'*********************************************************************************************
Public Function ToBSTR(Var As Variant) As String
Dim VType As Integer, Ptr As Long

   ' Get variant type
   VType = VarType(Var)
            
   ' Get string pointer
   MoveMemory Ptr, ByVal VarPtr(Var) + 8, 4
    
   If VType = VT_LPSTR Then ' ANSI String
       ToBSTR = Ptr2Str(Ptr, , False)
   ElseIf VType = VT_LPWSTR Then ' Unicode String
       ToBSTR = SysAllocString(Ptr)
   End If
    
   ' Clear the variant
   PropVariantClear Var
    
End Function

'*********************************************************************************************
'
' Converts a FILETIME variant to Date
'
' Parameters:
'
' Var: source variant
'
'*********************************************************************************************
Public Function ToDate(Var As Variant) As Date
Dim FT As Currency, ST As SYSTEMTIME, LocalFT As Currency
Dim Serial As Double

    ' Get FILETIME from variant
    MoveMemory FT, ByVal VarPtr(Var) + 8, Len(FT)
    
    ' Date properties are in UTC. Convert to
    ' Local time.
    FileTimeToLocalFileTime FT, LocalFT
    
    ' Convert FILETIME to SYSTEMTIME
    FileTimeToSystemTime LocalFT, ST
    
    ' Convert SYSTEMTIME to Date
    SystemTimeToVariantTime ST, Serial
    
    ' Set the return value
    ToDate = Serial
    
    ' Clear source variant
    PropVariantClear Var
    
End Function

'*********************************************************************************************
'
' Converts a Date to FILETIME variant
'
' Parameters:
'
' Value: source date
' Var: destination variant
'
'*********************************************************************************************
Public Sub ToFILETIME(ByVal Value As Date, Var As Variant)
Dim ST As SYSTEMTIME, FT As Currency

    ' Convert Date to SYSTEMTIME
    VariantTimeToSystemTime Value, ST
    
    ' Convert SYSTEMTIME to FILETIME
    SystemTimeToFileTime ST, FT
    
    ' Convert Local FILETIME to UTC FILETIME.
    ' Date properties must be saved in UTC.
    LocalFileTimeToFileTime FT, FT

    ' Clear any previous content
    PropVariantClear Var
    
    ' Set the variant type
    MoveMemory ByVal VarPtr(Var), VT_FILETIME, 2
    
    ' Copy the FILETIME to the variant
    MoveMemory ByVal VarPtr(Var) + 8, FT, Len(FT)
    
End Sub

'*********************************************************************************************
'
' Creates a LPSTR or LPWSTR variant from VB string
'
' Parameters:
'
' BSTR: Source string
' Var: destination variant
' Unicode: indicates if the result string must be ANSI or Unicode. Default is ANSI.
'
'*********************************************************************************************
Public Sub ToLPSTR(ByVal BSTR As String, Var As Variant, Optional ByVal Unicode As Boolean)
Dim VarType As Integer, Ptr As Long

    ' Set the string type
    If Unicode Then
        VarType = VT_LPWSTR ' Unicode
    Else
        VarType = VT_LPSTR  ' ANSI
    End If
    
    ' Add null char at the end of the
    ' string.
    BSTR = BSTR & vbNullChar
    
    If Unicode Then
        
        ' Allocate memory for the new string
        Ptr = CoTaskMemAlloc(Len(BSTR) * 2)
        
        ' Copy string from BSTR to the
        ' allocated memory
        MoveMemory ByVal Ptr, ByVal BSTR, Len(BSTR) * 2
    
    Else
    
        ' Allocate memory for the new string
        Ptr = CoTaskMemAlloc(Len(BSTR))
        
        ' Copy string from BSTR to the
        ' allocated memory
        lstrcpyA Ptr, ByVal BSTR
    
    End If
    
    ' Clear any previuos content from
    ' the variant
    PropVariantClear Var
    
    ' Write variant type
    MoveMemory Var, VarType, 2
 
    ' Write pointer
    MoveMemory ByVal VarPtr(Var) + 8, Ptr, 4

End Sub

'*********************************************************************************************
'
' Creates a array of strings from a variant containing a counted array of LPWSTR or LPSTR
'
' Parameters:
'
' Var: source variant
' Unicode: indicates if the source is ANSI or Unicode. Default is ANSI.
'
'*********************************************************************************************
Public Function ToBSTRArray(Var As Variant, Optional ByVal Unicode As Boolean) As Variant
Dim A() As String, Cnt As Long, PtrElem As Long
Dim PtrStr As Long

    ' Get element count from variant
    MoveMemory Cnt, ByVal VarPtr(Var) + 8, 4
    
    ' Get pointer to first element
    MoveMemory PtrElem, ByVal VarPtr(Var) + 12, 4

    ' Reallocate the VB array
    ReDim A(0 To Cnt - 1)
    
    For Cnt = 0 To Cnt - 1
        
        ' Get pointer to the string
        MoveMemory PtrStr, ByVal PtrElem, 4
        
        ' Copy the string from the pointer
        If Unicode Then
        
            A(Cnt) = Space$(lstrlenW(PtrStr))
            MoveMemory ByVal StrPtr(A(Cnt)), ByVal PtrStr, Len(A(Cnt)) * 2
            
        Else
        
            A(Cnt) = Space$(lstrlenA(PtrStr))
            lstrcpyA A(Cnt), ByVal PtrStr
            
        End If
        
        ' Move to next element
        PtrElem = PtrElem + 4
        
    Next
    
    ' Clear the source variant
    PropVariantClear Var
        
    ' Return the VB array
    ToBSTRArray = A
    
End Function

'*********************************************************************************************
'
' Creates a counted array of LPSTR from a VB array of strings
'
' Parameters:
'
' Value: source array
' Var: destination variant
'
'*********************************************************************************************
Public Sub ToLPSTRArray(Value As Variant, Var As Variant)
Dim ArrPtr As Long, ElemPtr As Long, PtrStr As Long
Dim Cnt As Long, I As Long, TmpStr As String

    ' Get element count
    Cnt = UBound(Value) - LBound(Value) + 1

    ' Alloc memory for the array. We
    ' must save each string pointer
    ' in the array. Each pointer have
    ' 4 bytes.
    ArrPtr = CoTaskMemAlloc(Cnt * 4)
    
    ' Set pointer to first element
    ElemPtr = ArrPtr
        
    For I = LBound(Value) To UBound(Value)
    
        ' Alloc memory for the string
        PtrStr = CoTaskMemAlloc(Len(Value(I)) + 1)
        
        ' Copy string pointer to array element
        MoveMemory ByVal ElemPtr, PtrStr, 4
        
        ' Copy string to string pointer
        TmpStr = Value(I) & vbNullChar
        lstrcpyA PtrStr, ByVal TmpStr
                
        ' Move element pointer to next element
        ElemPtr = ElemPtr + 4
        
    Next
    
    ' Set variant type
    MoveMemory ByVal VarPtr(Var), VT_VECTOR Or VT_LPSTR, 2
    
    ' Set variant element count
    MoveMemory ByVal VarPtr(Var) + 8, Cnt, 4
    
    ' Set Array pointer
    MoveMemory ByVal VarPtr(Var) + 12, ArrPtr, 4
    
End Sub



