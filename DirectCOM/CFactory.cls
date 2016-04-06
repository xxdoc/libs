VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Is there any way to instantiate an activex class from a specified .DLL
'file,without registering that file?
'
'Yes, I've written a Class for this:
'
'usage:
'Dim CF as CFactory,O as Object
'Set CF = New CFactory
'
'Set O = CF.GetInstance("C:\MyPath\scrrun.dll", "Dictionary")
'Set O = CF.GetInstance("C:\MyPath\scrrun.dll", "FileSystemObject")
'or this way
'Set O = CF.GetInstance("C:\MyPath\MyLib.dll", "MyClass")
'
'Here the Factory-Code:
'****Into a Class named CFactory (set a reference to EdanMos OleLib.tlb)


Option Explicit

'Direct Instanciation of COM-Objects (bypassing the registry)
'Olaf Schmidt - o...@datenhaus.de (2001)
'Uses Eduardo Morcillos OleLib.Tlb wich you will find here:
'http://www.domaindlx.com/e_morcillo/scripts/type/default.asp?page=ole

Private Declare Function LoadLibrary& Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName$)
Private Declare Function FreeLibrary& Lib "kernel32" (ByVal hLibModule&)
Private Declare Function GetProcAddress& Lib "kernel32" (ByVal hModule&, ByVallpProcName$)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal MSG&, ByVal wParam&, ByVallParam&)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal bLength&)

Private IIDCF As UUID, IIU As UUID, pIUNull As IUnknown
Private Libs As Collection, CFs As Collection, ASM() As Byte

Private Sub Class_Initialize()
  'Get IID for IClassFactory
  CLSIDFromString "{00000001-0000-0000-C000-000000000046}", IIDCF
  'same for IUnknown
  CLSIDFromString "{00000000-0000-0000-C000-000000000046}", IIU
  Set Libs = New Collection
  Set CFs = New Collection
  'Initialize ASM-Code for StdCall (principle found on VBDotCom)
  ASM = StrConv("XYYYYPh    h    h    Ë    √Ã", vbFromUnicode)
End Sub

Public Function GetInstance(sFile$, sClass$) As Object
Dim Result&, Lib&, pGetClass&, RefIU As IUnknown
Dim TLI As ITypeLib, CID As UUID, pICF As IClassFactory, pp&
  On Error Resume Next
  Lib = Libs(sFile)
  If Lib = 0 Then
    Lib = LoadLibrary(sFile)
  Else
    Set pICF = CFs(sFile & sClass)
    If Not pICF Is Nothing Then
      pICF.CreateInstance pIUNull, IIU, RefIU
      Set GetInstance = RefIU
      Exit Function
    End If
  End If
  If Lib = 0 Then Err.Clear: Exit Function
  Libs.Add Lib, sFile 'Cache the Library-Handle
  Set TLI = LoadTypeLibEx(sFile, REGKIND_NONE)
  If TLI Is Nothing Then Err.Clear: Exit Function

  If Not ClassExists(TLI, sClass, CID) Then Err.Clear: Exit Function

  pGetClass = GetProcAddress(Lib, "DllGetClassObject")
  If pGetClass = 0 Then Err.Clear: Exit Function

  RtlMoveMemory ASM(7), VarPtr(pICF), 4 'Param3
  RtlMoveMemory ASM(12), VarPtr(IIDCF), 4 'Param2
  RtlMoveMemory ASM(17), VarPtr(CID), 4 'Param1
  RtlMoveMemory ASM(22), pGetClass - VarPtr(ASM(22)) - 4, 4
  Result = CallWindowProc(VarPtr(ASM(0)), 0, 0, 0, 0)
  If Result = &H80040111 Then Err.Clear: Exit Function 'Bad ClassID
  If pICF Is Nothing Then Err.Clear: Exit Function 'couldn't get IClassFactory
  CFs.Add pICF, sFile & sClass
  pICF.CreateInstance pIUNull, IIU, RefIU
  Set GetInstance = RefIU
  Err.Clear
End Function

Private Function ClassExists(TLI As ITypeLib, sClass$, CID As UUID) As Boolean
Dim i&, sName$, Obj As Object, pAttr&
Dim TI As ITypeInfo, TA As TYPEATTR
  On Error Resume Next
  For i = 0 To TLI.GetTypeInfoCount - 1
    If TLI.GetTypeInfoType(i) <> TKIND_COCLASS Then GoTo nxt
    Set TI = TLI.GetTypeInfo(i)
    TI.GetDocumentation DISPID_UNKNOWN, sName, "", 0, ""
    If UCase(sName) <> UCase(sClass) Then GoTo nxt
    pAttr = TI.GetTypeAttr
    MoveMemory TA, ByVal pAttr, Len(TA)
    TI.ReleaseTypeAttr pAttr
    If TA.wTypeFlags Then CID = TA.iid: ClassExists = True: Exit For
nxt: Next i
  Err.Clear
End Function

Private Sub Class_Terminate()
Dim Lib
  Set CFs = Nothing
  For Each Lib In Libs: FreeLibrary Lib: Next
End Sub
