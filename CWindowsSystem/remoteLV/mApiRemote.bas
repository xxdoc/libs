Attribute VB_Name = "Module1"
'========mApiRemote.bas==========
'*****************************************************************
' Module to inject API call in remote process.
' Written by Arkadiy Olovyannikov (ark@msun.ru)
' Copyright 2005 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'*****************************************************************
Option Explicit

Public Enum ARG_FLAG
   arg_Value
   arg_Pointer
End Enum

'Structure for passing parameters in remote API calls
Public Type API_DATA
   lpData       As Long      'Pointer to data or real data
   dwDataLength As Long      'Data length
   argType      As ARG_FLAG  'ByVal or ByRef?
   bOut         As Boolean   'Is this argument [OUT]? If True,
                             'lpData will be filled with [out] data
End Type

Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Const INFINITE = -1&
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_READWRITE = &H4&

'Variables to store main kernel functions addresses
'This allow call GetProcAddress for kernel32 only once
Dim hKernel           As Long
Dim lpGetModuleHandle As Long
Dim lpLoadLibrary     As Long
Dim lpFreeLibrary     As Long
Dim lpGetProcAddress  As Long
Dim bKernelInit       As Boolean

Dim abAsm() As Byte 'buffer for assembly code
Dim lCP As Long     'used to keep track of latest byte added to assembly code

'********************************************************************************
'Public calling function. Prepare some data (retrieve function address)
'and call private CallFunctionRemote, which do the job.
'Input values are self-descriptive:
'hProcess  - handle to remote process
'LibName   - API library name  (e.g. "user32")
'FuncName  - API function name (e.g. "GetWindowTextA").
'*****Note: FuncName is CaSeSeNsItIvE!*****
'nParams   - # of function params (according normal API call)
'data()    - an array of input params with description (see API_DATA structure)
'dwTimeOut - timeout to wait API return from remote process.
'*****Note: In case of some incorrect call INFINITE timeout can hang your and/or
'remote app, so when debugging, use some FINITE value in milliseconds
'(e.g. 5000 means 5 sec)*****
'Return value - same as standard API call.
'*********************************************************************************
Public Function CallAPIRemote(ByVal hProcess As Long, ByVal LibName As String, _
                             ByVal FuncName As String, ByVal nParams As Long, _
                             data() As API_DATA, _
                             Optional ByVal dwTimeOut As Long = INFINITE) As Long
   
   If hProcess = GetCurrentProcess Then
'*****TODO*****:
'You can get my sample to CallAPIByName (see http://www.freevbcode.com/ShowCode.Asp?ID=1863)
'and use this call for current process instead of poppin message.
      MsgBox "Unfortunatelly, VB is single thread application." & vbCrLf & "You have to call standard APi in your process address space", vbCritical, "Remote API error"
      Exit Function
   End If
   
   Dim hLib As Long, fnAddress As Long
   Dim bNeedUnload As Boolean
   Dim locData(1) As API_DATA
     
   hLib = GetModuleHandleRemote(hProcess, LibName)
   If hLib = 0 Then
      hLib = LoadLibraryRemote(hProcess, LibName)
      If hLib = 0 Then Exit Function
      bNeedUnload = True
   End If
   
   fnAddress = GetProcAddressRemote(hProcess, hLib, FuncName)
   If fnAddress Then
      CallAPIRemote = CallFunctionRemote(hProcess, fnAddress, nParams, data, dwTimeOut)
   End If
   If bNeedUnload Then Call FreeLibraryRemote(hProcess, hLib)
'*****TODO: API set last error in remote process!
'Use ErrNum = CallFunctionRemoteOneParam(hProcess,lpGetLastError,0,0,0,0)
'Where lpLastError = GetProcAddress(hKernel,"GetLastError")
'and, probably, SetLastError ErrNum to set same error in your process.
End Function

'*****************************************************************
'Main function which do the job.
'Parameters are same as in above function, except of func_address -
'function address in remote process.
'*****************************************************************
Private Function CallFunctionRemote(ByVal hProcess As Long, ByVal func_addr As Long, _
                                    ByVal nParams As Long, data() As API_DATA, _
                                    Optional ByVal dwTimeOut As Long = INFINITE) As Long
   Dim hThread As Long, ThreadId As Long
   Dim addr As Long, ret As Long, h As Long, i As Long
   Dim codeStart As Long
   Dim param_addr() As Long
   
   If nParams = 0 Then
      CallFunctionRemote = CallFunctionRemoteOneParam(hProcess, func_addr, 0, 0, 0, 0)
   ElseIf nParams = 1 Then
      CallFunctionRemote = CallFunctionRemoteOneParam(hProcess, func_addr, 1, _
                           data(0).lpData, data(0).dwDataLength, data(0).argType, _
                           data(0).bOut)
   End If
   
   ReDim abAsm(50 + 6 * nParams)
   ReDim param_addr(nParams - 1)
   lCP = 0
   addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal UBound(abAsm) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   
   codeStart = GetAlignedCodeStart(addr)
   lCP = codeStart - addr
   For i = 0 To lCP - 1
       abAsm(i) = &HCC
   Next
   PrepareStack 1 'remove ThreadFunc lpParam
   Dim s As String
   s = "MessageBoxA" & Chr(0)
   For i = nParams To 1 Step -1
       AddByteToCode &H68 'push wwxxyyzz
       If data(i - 1).argType = arg_Value Then
          If data(i - 1).dwDataLength > 4 Then
             MsgBox "Arguments passing as Value should not exeed 4 bytes (long)", vbCritical
             GoTo CleanUp
          End If
          AddLongToCode data(i - 1).lpData
       Else
          param_addr(i - 1) = VirtualAllocEx(ByVal hProcess, ByVal 0&, _
                              ByVal data(i - 1).dwDataLength, MEM_RESERVE Or MEM_COMMIT, _
                              PAGE_READWRITE)
          If param_addr(i - 1) = 0 Then GoTo CleanUp
          If WriteProcessMemory(hProcess, ByVal param_addr(i - 1), ByVal data(i - 1).lpData, _
                               data(i - 1).dwDataLength, ret) = 0 Then GoTo CleanUp
          AddLongToCode param_addr(i - 1)
       End If
   Next
   AddCallToCode func_addr, addr + VarPtr(abAsm(lCP)) - VarPtr(abAsm(0))
   AddByteToCode &HC3
   AddByteToCode &HCC
   If WriteProcessMemory(hProcess, ByVal addr, abAsm(0), UBound(abAsm) + 1, ret) = 0 Then GoTo CleanUp
   hThread = CreateRemoteThread(hProcess, 0, 0, ByVal codeStart, data(0).lpData, 0&, ThreadId)
   If hThread Then
      ret = WaitForSingleObject(hThread, dwTimeOut)
      If ret = 0 Then ret = GetExitCodeThread(hThread, h)
   End If
   CallFunctionRemote = h
   For i = 0 To nParams - 1
       If param_addr(i) <> 0 Then
          If data(i).bOut Then
             ReadProcessMemory hProcess, ByVal param_addr(i), ByVal data(i).lpData, data(i).dwDataLength, ret
          End If
       End If
   Next i
CleanUp:
   VirtualFreeEx hProcess, ByVal addr, 0, MEM_RELEASE
   For i = 0 To nParams - 1
       If param_addr(i) <> 0 Then VirtualFreeEx hProcess, ByVal param_addr(i), 0, MEM_RELEASE
   Next i
End Function

'******************************************************************************
'Simplified version of above function - one parameter doesn't require asm code.
'******************************************************************************
Private Function CallFunctionRemoteOneParam(ByVal hProcess As Long, ByVal func_addr As Long, _
                                    ByVal nParams As Long, ByVal lngVal As Long, _
                                    ByVal dwSize As Long, ByVal argType As ARG_FLAG, _
                                    Optional ByVal bReturn As Boolean) As Long
   Dim hThread As Long, ThreadId As Long
   Dim addr As Long, ret As Long, h As Long, i As Long
   Dim lngTemp As Long
   If nParams = 0 Then
      bReturn = False
   Else
      If argType = arg_Pointer Then
          addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal dwSize, _
                                MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
          If addr = 0 Then Exit Function
          Call WriteProcessMemory(hProcess, ByVal addr, ByVal lngVal, dwSize, ret)
          lngTemp = addr
      Else
          lngTemp = lngVal
      End If
   End If
   hThread = CreateRemoteThread(hProcess, 0, 0, ByVal func_addr, lngTemp, 0&, ThreadId)
   If hThread Then
      ret = WaitForSingleObject(hThread, 1000)
      If ret = 0 Then ret = GetExitCodeThread(hThread, h)
      CallFunctionRemoteOneParam = h
      CloseHandle hThread
   End If
   If bReturn Then
      If addr <> 0 Then
         ReadProcessMemory hProcess, ByVal addr, ByVal lngVal, dwSize, ret
         VirtualFreeEx hProcess, ByVal addr, 0, MEM_RELEASE
      End If
   End If
End Function

'*****************************************************************
'Some usefull Public functions for loading/unloading libraries.
'[in]/[out] parameters are same as in appropriate API cals
'except of remote process handle (hProcess).
'*****************************************************************
Public Function GetModuleHandleRemote(ByVal hProcess As Long, ByVal LibName As String) As Long
   If Not InitKernel Then Exit Function
   If GetModuleHandle(LibName) = hKernel Then
      GetModuleHandleRemote = hKernel
      Exit Function
   End If

   Dim hThread As Long, ThreadId As Long
   Dim addr As Long, ret As Long, h As Long
   addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal Len(LibName) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   If addr = 0 Then Exit Function
   If WriteProcessMemory(hProcess, ByVal addr, ByVal LibName, Len(LibName), ret) Then
      hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpGetModuleHandle, addr, 0&, ThreadId)
      If hThread Then
         ret = WaitForSingleObject(hThread, 500)
         If ret = 0 Then ret = GetExitCodeThread(hThread, h)
      End If
   End If
   VirtualFreeEx hProcess, ByVal addr, 0, MEM_RELEASE
   CloseHandle hThread
   GetModuleHandleRemote = h
End Function

Public Function LoadLibraryRemote(ByVal hProcess As Long, ByVal LibName As String) As Long
   If Not InitKernel Then Exit Function
   If GetModuleHandle(LibName) = hKernel Then
      LoadLibraryRemote = hKernel
      Exit Function
   End If
   
   Dim hThread As Long, ThreadId As Long
   Dim addr As Long, ret As Long, h As Long
   addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal Len(LibName) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   If addr = 0 Then Exit Function
   If WriteProcessMemory(hProcess, ByVal addr, ByVal LibName, Len(LibName), ret) Then
      hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpLoadLibrary, addr, 0&, ThreadId)
      If hThread Then
         ret = WaitForSingleObject(hThread, 500)
         If ret = 0 Then ret = GetExitCodeThread(hThread, h)
      End If
   End If
   LoadLibraryRemote = h
End Function

Public Function GetProcAddressRemote(ByVal hProcess As Long, ByVal hLib As Long, ByVal fnName As String) As Long
   If Not InitKernel Then Exit Function
   
   If hLib = hKernel Then
      GetProcAddressRemote = GetProcAddress(hKernel, fnName)
      Exit Function
   End If
   Dim localData(1) As API_DATA
   Dim abName() As Byte
   With localData(0)
      .lpData = hLib
      .dwDataLength = 4
      .argType = arg_Value
   End With
   fnName = fnName & Chr(0)
   abName = StrConv(fnName, vbFromUnicode)
   With localData(1)
      .lpData = VarPtr(abName(0))
      .dwDataLength = UBound(abName) + 1
      .argType = arg_Pointer
   End With
   GetProcAddressRemote = CallFunctionRemote(hProcess, lpGetProcAddress, 2, localData)
End Function

Public Function FreeLibraryRemote(ByVal hProcess As Long, ByVal hLib As Long) As Long
   If Not InitKernel Then Exit Function
   If hLib = hKernel Then
      FreeLibraryRemote = True
      Exit Function
   End If
   
   Dim hThread As Long, ThreadId As Long, h As Long, ret As Long
   hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpFreeLibrary, hLib, 0&, ThreadId)
   If hThread Then
      ret = WaitForSingleObject(hThread, 500)
      If ret = 0 Then ret = GetExitCodeThread(hThread, h)
   End If
   CloseHandle hThread
   FreeLibraryRemote = h
End Function

'============Private routines to prepare asm (op)code===========
Private Sub AddCallToCode(ByVal dwAddress As Long, ByVal BaseAddr As Long)
    AddByteToCode &HE8
    AddLongToCode dwAddress - BaseAddr - 5
End Sub

Private Sub AddLongToCode(ByVal lng As Long)
    Dim i As Integer
    Dim byt(3) As Byte
    CopyMemory byt(0), lng, 4
    For i = 0 To 3
        AddByteToCode byt(i)
    Next
End Sub

Private Sub AddByteToCode(ByVal byt As Byte)
    abAsm(lCP) = byt
    lCP = lCP + 1
End Sub

Private Function GetAlignedCodeStart(ByVal dwAddress As Long) As Long
    GetAlignedCodeStart = dwAddress + (15 - (dwAddress - 1) Mod 16)
    If (15 - (dwAddress - 1) Mod 16) = 0 Then GetAlignedCodeStart = GetAlignedCodeStart + 16
End Function

Private Sub PrepareStack(ByVal numParamsToRemove As Long)
    If numParamsToRemove = 0 Then Exit Sub
    Dim i As Long
    AddByteToCode &H58     'pop eax -  pop return address
    For i = 1 To numParamsToRemove
        AddByteToCode &H59 'pop ecx -  kill param
    Next i
    AddByteToCode &H50     'push eax - put return address back
End Sub

Private Sub ClearStack(ByVal nParams As Long)
   Dim i As Long
   For i = 1 To nParams
       AddByteToCode &H59  'pop ecx - remove params from stack
   Next
End Sub

'==========Get main kernel32 functions addresses=========
Private Function InitKernel() As Boolean
   If bKernelInit Then
      InitKernel = True
      Exit Function
   End If
   hKernel = GetModuleHandle("kernel32")
   If hKernel = 0 Then Exit Function
   lpGetProcAddress = GetProcAddress(hKernel, "GetProcAddress")
   lpGetModuleHandle = GetProcAddress(hKernel, "GetModuleHandleA")
   lpLoadLibrary = GetProcAddress(hKernel, "LoadLibraryA")
   lpFreeLibrary = GetProcAddress(hKernel, "FreeLibrary")
   InitKernel = True
   bKernelInit = True
End Function

