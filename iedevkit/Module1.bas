Attribute VB_Name = "modGlobals"
Option Explicit
   
Public Declare Function StringFromGUID2 Lib "OLE32" (ByRef lpGUID As UUID, _
                                                      ByVal lpszBuff As Long, _
                                                      ByVal intcb As Long) As Long
   
Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function sGuid(gd As UUID) As String
    sGuid = String(38, 0)
    StringFromGUID2 gd, StrPtr(sGuid), 39
End Function
