VERSION 5.00
Begin VB.UserControl uCallback 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "uCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* uCallback - UserControl generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

Option Explicit

'-Callback declarations---------------------------------------------------------------------------
Private z_CbMem   As Long                                                   'Callback allocated memory address
Private z_Cb()    As Long                                                   'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

Private Sub UserControl_Terminate()
  cb_Terminate
End Sub

'-Callback code-----------------------------------------------------------------------------------
Private Function cb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  cb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Terminate the callback thunks
Private Sub cb_Terminate()
Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag

  If z_CbMem <> 0 Then                                                      'If memory allocated
    If VirtualFree(z_CbMem, 0, MEM_RELEASE) <> 0 Then                       'Release
      z_CbMem = 0                                                           'Indicate memory released
    End If
  End If
End Sub

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************

'Callback ordinal 3
'Private Function Ordinal3(...) As Long
'End Function

'Callback ordinal 3
'Private Function Ordinal2(...) As Long
'End Function

'Callback ordinal 1
'Private Function Ordinal1(...) As Long
'End Function

