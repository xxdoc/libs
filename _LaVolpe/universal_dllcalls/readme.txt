I was recently made aware that a function I've used from time to time for calling virtual functions of COM objects was perfectly adept at calling functions from just about any standard DLL out there. So, I whipped up a 'generic' class that can call both standard DLL functions & COM VTable functions. No thunks are used, just a couple of supporting API calls in the main class, including the low-level core API function: DispCallFunc

What does this mean for you? Well, it does allow you to call DLL functions from nearly 10 different calling conventions, including the two most common: StdCall & CDecl. It also allows you to call virtual functions from COM objects. And if you wish, it means you do not have to declare a single API function declaration in your VB project. Though, personally, I'd use it for calling DLL conventions other than StdCall.

I'd consider this topic advanced for one reason only. This is very low level. If you provide incorrect parameter information to the class, your project is likely to crash. For advanced coders, we have no problem doing the research to understand what parameter information is required, be it variable type, a pointer, a pointer to a pointer, function return types, etc, etc. Not-so-advanced coders just want to plug in values & play, but when playing at such a low level, that usually results in crashes, frustration.

The attachment includes very simple examples of calling DLL functions and calling a COM virtual function. You will notice that the form has no API function declarations, though several DLL functions are called & executed correctly. A sample call to the class might look like:
Code:

Debug.Print myClass.CallFunction_DLL("user32.dll", "IsWindowUnicode", STR_NONE, CR_LONG, _
                                        CC_STDCALL, Me.hWnd)

For DLL calls, the class takes the DLL name and function name to be called. Technically, you aren't passing the function pointer to the class. However, the class does make the call to the pointer, not via declared API functions. Just thought I'd throw this comment in, should someone suggest we aren't really calling functions by pointer. The class is, the user calling the class is not, but can be if inclined to modify the code a bit.

Limitations: Callbacks from non-StdCall DLLs/functions
If whatever function you are calling requires a callback pointer, then stack corruption is likely from all calling conventions where you pass a VB function pointer as the callback address. The exceptions are stdCall DLLs and also CDecl calls, if the thunk/patch option in the class is used.

Tip: If you really like this class, you may want to instantiate one for each DLL you will be calling quite often. This could speed things up a bit when making subsequent calls. As is, the class will load the requested DLL into memory if it isn't already. Once class is called again, for a different DLL, then the previous DLL is unloaded if needed & the new DLL loaded as needed. So, if you created cUser32, cShell32, cKernel32 instances, less code is executed in the class if it doesn't have to drop & load DLLs.
vb Code:

    ' top of form
    Private cUser32 As cUniversalDLLCalls
    Private cKernel32 As cUniversalDLLCalls
    Private cShell32 As cUniversalDLLCalls
     
    ' in form load
    Set cUser32 = New cUniversalDLLCalls
    Set cKernel32 = New cUniversalDLLCalls
    Set cShell32 = New cUniversalDLLCalls
    ' now use cUser32 for all user32.dll calls, cKernel32 for kernel32, cShell32 for shell32, etc

Tip: When using the STR_ANSI flag to indicate the passed parameters include string values destined for ANSI functions, the class will convert the passed string to ANSI before calling the function. Doing so, default Locale is used for string conversion. If this is a problem, you should ensure you convert the string(s) to ANSI before passing it to the class. If you do this conversion, use STR_NONE & pass the string via StrPtr(). FYI: strings used strictly as a buffer for return values should always be passed via StrPtr() and the flag STR_NONE used; regardless if destined for ANSI or unicode functions. ANSI strings are never passed to COM interfaces. Always use StrPtr(theString) for any string parameters to those COM methods.
vb Code:

    ' how to have a VB string contain ANSI vs Unicode
    myString = StrConv(myString, vbFromUnicode, [Locale ID])
    ' how to convert the returned ANSI string to a proper VB string
    myString = StrConv(myString, vbUnicode, [Locale ID])

Tip: If you ever need to call a private COM interface function by its pointer/address, post #24 below shows how that can be done. A slight modification to the attached class is required.

Change History
- 28 Nov 2014. Added thunk/patch/workaround for passing a VB function address to a CDECL dll that expects a CDECL callback address. Sample found in post #10 below


This post will be dedicated to linking to known good posts/URLs of interfaces that not only provide their GUID, but the VTable Order. Will update this post from time to time as others add information to this thread.

So if you researched an Interface and want to share it, please post it & include at least these 3 basic pieces of information: GUID, VTable order & link describing the virtual functions. An example would be appreciated also.

You cannot assume that the listing of functions provided on MSDN pages is the actual VTable order. It used to be, but no longer is reliable. Order is extremely important, because virtual functions are called relative to their offset from the inherited IUnknown interface. VTable entries are in multiples of four.

A starter page. Layout of the COM object

So, you have a string GUID, how do you get it to a Long value for passing to appropriate functions? Simple:
Code:

Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, piid As Any) As Long
' sample of getting the IDataObject GUID
Dim aGUID(0 To 3) As Long
Call IIDFromString(StrPtr("{0000010e-0000-0000-C000-000000000046}"), ByVal VarPtr(aGUID(0)))

How do we know if an object supports a specific interface? We ask the object...
Code:

Dim IID_IPicture As Long, aGUID(0 To 3) As Long, sGUID As String
Dim c As cUniversalDLLCalls
Const IUnknownQueryInterface As Long = 0&   ' IUnknown vTable offset to Query implemented interfaces
Const IUnknownRelease As Long = 8&          ' IUnkownn vTable offset to decrement reference count

    ' ask if Me.Icon picture object supports IPicture
    sGuid = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
    Set c = New cUniversalDLLCalls
    c.CallFunction_DLL "ole32.dll", "IIDFromString", STR_NONE, CR_LONG, CC_STDCALL, StrPtr(sGUID), VarPtr(aGuid(0))
    c.CallFunction_COM ObjPtr(Me.Icon), IUnknownQueryInterface, CR_LONG, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IID_IPicture)
    If IID_IPicture <> 0& Then
        ' do stuff
        ' Release the IPicture interface at some point. QueryInterface calls AddRef internally
        c.CallFunction_COM IID_IPicture, IUnknownRelease, CR_LONG, CC_STDCALL
    End If

Here's a few interfaces to start this thread out...

IUnknown: GUID {00000000-0000-0000-C000-000000000046}
VTable Order: QueryInterface, AddRef, Release

IPicture: GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
VTable Order: GetHandle, GetHPal, GetType, GetWidth, GetHeight, Render, SetHPal, GetCurDC,
SelectPicture, GetKeepOriginalFormat, SetKeepOriginalFormat, PictureChanged, SaveAsFile, GetAttributes

IDataObject: GUID {0000010e-0000-0000-C000-000000000046}
VTable Order: GetData, GetDataHere, QueryGetData, GetCanonicalFormatEtc, SetData,
EnumFormatEtc, DAdvise, DUnadvise, EnumDAdvise

Tip #1. Get the IDataObject from the Data parameter of VB's OLEDrag[...] events
Code:

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

Dim IID_DataObject As Long
CopyMemory IID_DataObject, ByVal ObjPtr(Data) + 16&, 4&
' you now have an unreferenced pointer to the IDataObject

Tip #2. Get IDataObject of the clipboard
Code:

Private Declare Function OleGetClipboard Lib "ole32.dll" (ByRef ppDataObj As Long) As Long

Dim IID_DataObject As Long
OleGetClipboard IID_DataObject
' if IID_DataObject is non-null, you have a referenced pointer to the IDataObject
' Referenced pointers must call IUnknown.Release

IOLEObject: GUID {00000112-0000-0000-C000-000000000046}
VTable Order: SetClientSite, GetClientSite, SetHostNames, Close, SetMoniker, GetMoniker,
InitFromData, GetClipboardData, DoVerb, EnumVerbs, Update, IsUpToDate, GetUserClassID, GetUserType, SetExtent, GetExtent, Advise, EnumAdvise, GetMiscStatus, SetColorScheme

IStream: inherits IUnknown:ISequentialStream. GUID {0000000C-0000-0000-C000-000000000046}
VTable Order: Read [from ISequentialStream], Write [from ISequentialStream], Seek, SetSize,
CopyTo, Commit, Revert, LockRegion, UnlockRegion, Stat, Clone

ITypeLib: GUID {00020402-0000-0000-C000-000000000046}
VTable Order: GetTypeInfoCount, GetTypeInfo, GetTypeInfoType, GetLibAttr,
GetTypeComp, GetDocumentation, IsName, FindName, ReleaseTLibAttr

ITypeInfo: GUID {00020401-0000-0000-C000-000000000046}
VTable Order: GetTypeAttr, GetTypeComp, GetFuncDesc, GetVarDesc, GetNames,
GetRefTypeOfImplType, GetImplTypeFlags, GetIDsOfNames, Invoke, GetDocumentation, GetDLLEntry, GetRefTypeInfo, AddressOfMember, CreateInstance, GetMops, GetContainingTypeLib, ReleaseTypeAttr, ReleaseFuncDesc, RelaseVarDesc


---------------------------------------------------------------------------------------------------

Low-level helper for Invoke that provides machine independence for customized Invoke.
Syntax
C++


HRESULT DispCallFunc(
   void       *pvInstance,
   ULONG_PTR  oVft,
   CALLCONV   cc,
   VARTYPE    vtReturn,
   UINT       cActuals,
   VARTYPE    *prgvt,
   VARIANTARG **prgpvarg,
   VARIANT    *pvargResult
);

Parameters

pvInstance

    An instance of the interface described by this type description.
oVft

    For FUNC_VIRTUAL functions, specifies the offset in the VTBL.
cc

    The calling convention. One of the CALLCONV values, such as CC_STDCALL.
vtReturn

    The variant type of the function return value. Use VT_EMPTY to represent void.
cActuals

    The number of function parameters.
prgvt

    An array of variant types of the function parameters.
prgpvarg

    The function parameters.
pvargResult

    The function result.

Return value

If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
Requirements

Header
	

OleAuto.h

Library
	

OleAut32.lib

DLL
	

OleAut32.dll


