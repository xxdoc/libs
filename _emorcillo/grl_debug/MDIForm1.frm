VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "VB Debugger"
   ClientHeight    =   5835
   ClientLeft      =   1485
   ClientTop       =   4185
   ClientWidth     =   8865
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuODbgChild 
         Caption         =   "Debug &Child Processes"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange &Icons"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Debugging processes with VB
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98 + VB5
'
' History:
'           03/09/2000 - This code was released
'
'*********************************************************************************************
Option Explicit

Dim m_bExitLoop As Boolean

'
' FindDebugForm
'
' Finds the debug window for applicacion with
' process ID = PID
'
Function FindDebugForm(ByVal PID As Long) As frmDebugger
Dim F As Form

   For Each F In Forms
      If TypeOf F Is frmDebugger Then
         If F.lPID = PID Then
            Set FindDebugForm = F
            Exit For
         End If
      End If
   Next

End Function


Private Sub MDIForm_Load()

   Show
   
   EnterDebugLoop
   
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

   If Me.ActiveForm Is Nothing Then
      m_bExitLoop = True
   End If
   
End Sub


Private Sub mnuFOpen_Click()

   On Error Resume Next
   
   With CommonDialog1
      .CancelError = True
      .DefaultExt = "exe"
      .Filter = "Programs|*.exe"
      .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
   
      .ShowOpen
   
      If Err.Number = 0 Then
         DebugProcess .filename, InputBox("Type the command line arguments", .filename)
      End If
      
   End With
   
End Sub
'
' CreateProcHandler
'
' Creates a new debugger form for
' the new process
'
Private Sub CreateProcHandler(DE As DEBUG_EVENT)
Dim oDbgForm As frmDebugger
Dim ProcInfo As CREATE_PROCESS_DEBUG_INFO

   ' Create the debug window
   Set oDbgForm = New frmDebugger
   Load oDbgForm
   
   ' Get the CREATE_PROCESS_DEBUG_INFO
   ' structure from the byte array
   MoveMemory ProcInfo, DE.DEBUG_INFO(0), Len(ProcInfo)

   With oDbgForm.lvwThreads.ListItems
      With .Add(, "#" & DE.dwThreadId, Hex1(DE.dwThreadId, 8))
         .SubItems(1) = Hex1(ProcInfo.lpStartAddress, 8)
      End With
   End With
   
   ' Change the window caption
   oDbgForm.Caption = StrFromPtrPtrPID(ProcInfo.lpImageName, DE.dwProcessId, ProcInfo.fUnicode)
   
   ' Set the flags
   oDbgForm.lPID = DE.dwProcessId
   oDbgForm.bRunning = True
 
   ' Show the form
   oDbgForm.Show
   
End Sub

Private Sub CreateThreadHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim ThreadInfo As CREATE_THREAD_DEBUG_INFO
   
   ' Get the CREATE_THREAD_DEBUG_INFO
   ' structure from the byte array
   MoveMemory ThreadInfo, DE.DEBUG_INFO(0), Len(ThreadInfo)
   
   With DbgForm.lvwThreads.ListItems
      With .Add(, "#" & DE.dwThreadId, Hex1(DE.dwThreadId, 8))
         .SubItems(1) = Hex1(ThreadInfo.lpStartAddress, 8)
      End With
   End With

End Sub

Private Sub DebugProcess(ByVal ExeName As String, ByVal CmdLine As String)
Dim SI As STARTUPINFO
Dim PI As PROCESS_INFORMATION
  
   ' Launch the process
   If CreateProcess(vbNullString, _
         ExeName & " " & CmdLine, _
         ByVal 0&, _
         ByVal 0&, _
         False, _
         DEBUG_PROCESS Or NORMAL_PRIORITY_CLASS Or (-DEBUG_ONLY_THIS_PROCESS * (Not mnuODbgChild.Checked)), _
         ByVal 0&, _
         vbNullString, SI, PI) Then
   
      ' Close the handles
      ' returned by CreateProcess
      CloseHandle PI.hThread
      CloseHandle PI.hProcess
      
   Else
      
      MsgBox "The program cannot be lauched. GetLastError = " & Err.LastDllError, vbCritical, ExeName
      
   End If
   
End Sub

'
' EnterDebugLoop
'
' Waits for debug events and
' processes them.
'
Private Sub EnterDebugLoop()
Dim DE As DEBUG_EVENT
Dim lRes As Long
Dim lStatus As Long
Dim oDbgForm As frmDebugger

   Do Until m_bExitLoop
           
      ' Wait 1ms for a debug event
      lRes = WaitForDebugEvent(DE, 1)
      
      ' If lRes <> 0 there's a
      ' debug event in DE
      If lRes Then
      
         ' Get the form
         Set oDbgForm = FindDebugForm(DE.dwProcessId)
         
         ' Set the default status
         lStatus = DBG_CONTINUE
      
         Select Case DE.dwDebugEventCode
            
            Case CREATE_PROCESS_DEBUG_EVENT
            
               ' This is the first debug
               ' event. It's received after
               ' the process is created.
            
               CreateProcHandler DE
               
            Case CREATE_THREAD_DEBUG_EVENT
               
               ' A thread was started
               
               CreateThreadHandler DE, oDbgForm
               
            Case EXIT_THREAD_DEBUG_EVENT
               
               ' A thread has finished
               
               ExitThreadHandler DE, oDbgForm
            
            Case EXIT_PROCESS_DEBUG_EVENT
            
               ' The process has ended.
            
               ExitProcHandler DE, oDbgForm
               
            Case LOAD_DLL_DEBUG_EVENT
            
               ' A DLL was loaded
            
               LoadDLLHandler DE, oDbgForm
               
            Case UNLOAD_DLL_DEBUG_EVENT
               
               ' A DLL was unloaded
               
               UnloadDLLHandler DE, oDbgForm
               
            Case EXCEPTION_DEBUG_EVENT
            
               ' An exception was raised
               ' in the debugged process
               
               lStatus = ExceptionHandler(DE, oDbgForm)
                                
            Case OUTPUT_DEBUG_STRING_EVENT
          
               ' The debugged process has
               ' called OutputDebugString
          
               OutputHandler DE, oDbgForm
               
         End Select
          
         ' Continue the process execution
         ContinueDebugEvent DE.dwProcessId, DE.dwThreadId, lStatus
      
      End If
      
      ' Allow this process
      ' to process messages
      DoEvents
            
   Loop
   
End Sub


Private Function ExceptionHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger) As Long
Dim ExcepInfo As EXCEPTION_DEBUG_INFO
Dim Str As String
   
   ' Get EXCEPTION_DEBUG_INFO structure
   ' from byte array
   MoveMemory ExcepInfo, DE.DEBUG_INFO(0), Len(ExcepInfo)
   
   Select Case ExcepInfo.pExceptionRecord.ExceptionCode
      
      Case EXCEPTION_SINGLE_STEP
         ExceptionHandler = DBG_CONTINUE
         
      Case EXCEPTION_BREAKPOINT
      
         ' The process has set a breakpoint.
         
         With DbgForm.txtOutput
            .SelStart = Len(.Text)
            .SelText = "Breakpoint at address " & Hex1(ExcepInfo.pExceptionRecord.ExceptionAddress, 8) & vbCrLf
         End With
         
         ExceptionHandler = DBG_CONTINUE
      
      Case Else
      
         ' An error has occurred.
         ' Make ContinueDebugEvent to
         ' call the process exception
         ' handler.
         
         Str = "Exception error:" & vbCrLf & vbCrLf
         
         Select Case ExcepInfo.pExceptionRecord.ExceptionCode
            Case EXCEPTION_ACCESS_VIOLATION
               Str = Str & "Access violation"
            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
               Str = Str & "Array bounds exceeded"
            Case EXCEPTION_DATATYPE_MISALIGNMENT
               Str = Str & "Data type misalignment"
            Case EXCEPTION_FLT_DIVIDE_BY_ZERO, EXCEPTION_INT_DIVIDE_BY_ZERO
               Str = Str & "Division by zero"
            Case EXCEPTION_FLT_OVERFLOW, EXCEPTION_INT_OVERFLOW
               Str = Str & "Overflow"
            Case EXCEPTION_FLT_DENORMAL_OPERAND
               Str = Str & "Denormal Operand"
            Case EXCEPTION_FLT_INEXACT_RESULT
               Str = Str & "Float inexact result"
            Case EXCEPTION_FLT_STACK_CHECK
               Str = Str & "Float stack check"
            Case EXCEPTION_FLT_UNDERFLOW
               Str = Str & "Float underflow"
            Case EXCEPTION_ILLEGAL_INSTRUCTION
               Str = Str & "Illegal instruction"
            Case EXCEPTION_IN_PAGE_ERROR
               Str = Str & "In page error"
            Case EXCEPTION_INVALID_DISPOSITION
               Str = Str & "Invalid disposition"
            Case EXCEPTION_NONCONTINUABLE_EXCEPTION
               Str = Str & "Non-continuable exception"
            Case EXCEPTION_PRIV_INSTRUCTION
               Str = Str & "Priv instruction"
            Case EXCEPTION_STACK_OVERFLOW
               Str = Str & "Stack overflow"
            Case Else
               Str = Str & "Unknown"
         End Select
         
         Str = Str & " at address " & Hex1(ExcepInfo.pExceptionRecord.ExceptionAddress, 8)
         Str = Str & vbCrLf & vbCrLf & "Do you want to continue execution"
         If ExcepInfo.pExceptionRecord.ExceptionFlags = EXCEPTION_NONCONTINUABLE Then
            Str = Str & " (non-continuable exception)"
         End If
         Str = Str & "?"
         
         If MsgBox(Str, vbYesNo Or vbCritical, DbgForm.Caption) = vbYes Then
            ExceptionHandler = DBG_CONTINUE
         Else
            ExceptionHandler = DBG_EXCEPTION_NOT_HANDLED
         End If
                                 
   End Select
   
End Function


Private Sub ExitProcHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim ExitInfo As EXIT_PROCESS_DEBUG_INFO
   
   ' Get EXIT_PROCESS_DEBUG_INFO structure
   ' from the byte array
   MoveMemory ExitInfo, DE.DEBUG_INFO(0), Len(ExitInfo)
      
   ' Change the window caption
   DbgForm.Caption = DbgForm.Caption & " - Exit Code: " & ExitInfo.dwExitCode
   
   DbgForm.bRunning = False
   
End Sub

Private Sub ExitThreadHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim EThreadInfo As EXIT_THREAD_DEBUG_INFO
   
   ' Get the CREATE_THREAD_DEBUG_INFO
   ' structure from the byte array
   MoveMemory EThreadInfo, DE.DEBUG_INFO(0), Len(EThreadInfo)
   
   DbgForm.lvwThreads.ListItems.Remove "#" & DE.dwThreadId
   
End Sub

Private Function GetPath(ByVal File As String) As String
Dim lIdx As Long

   For lIdx = Len(File) To 1 Step -1
      If Mid$(File, lIdx, 1) = "\" Then
         GetPath = Left$(File, lIdx - 1)
         Exit Function
      End If
   Next
   
End Function

Private Function GetFilename(ByVal File As String) As String
Dim lIdx As Long

   For lIdx = Len(File) To 1 Step -1
      If Mid$(File, lIdx, 1) = "\" Then
         GetFilename = Mid$(File, lIdx + 1)
         Exit Function
      End If
   Next
   
End Function


Function Hex1(ByVal Val As Long, ByVal cc As Long) As String

   Hex1 = Hex$(Val)
   
   Do While Len(Hex1) < cc
      Hex1 = "0" & Hex1
   Loop
   
   Hex1 = "&H" & Hex1
   
End Function


Private Sub LoadDLLHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim DLLInfo As LOAD_DLL_DEBUG_INFO
Dim Str As String

   ' Get the LOAD_DLL_DEBUG_INFO structure
   ' from the byte array
   MoveMemory DLLInfo, DE.DEBUG_INFO(0), Len(DLLInfo)
   
   ' Add the DLL path to the DLL listbox
   With DbgForm.lvwDLLs.ListItems
      
      Str = StrFromPtrPtrPID(DLLInfo.lpImageName, DE.dwProcessId, DLLInfo.fUnicode)
      
      On Error Resume Next
      
      With .Add(, "#" & CStr(DLLInfo.lpBaseOfDll), GetFilename(Str))
         .SubItems(1) = Hex1(DLLInfo.lpBaseOfDll, 8)
         .SubItems(2) = Hex1(DE.dwThreadId, 8)
         .SubItems(3) = GetPath(Str)
      End With
      
   End With

End Sub

Private Sub OutputHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim ODSInfo As OUTPUT_DEBUG_STRING_INFO

   ' Get OUTPUT_DEBUG_STRING_INFO structure
   ' from byte array
   MoveMemory ODSInfo, DE.DEBUG_INFO(0), Len(ODSInfo)
   
   If ODSInfo.nDebugStringLength > 0 Then
      DbgForm.txtOutput.SelStart = Len(DbgForm.txtOutput.Text)
      DbgForm.txtOutput.SelText = StrFromPtrPID(ODSInfo.lpDebugStringData, DE.dwProcessId, ODSInfo.fUnicode, ODSInfo.nDebugStringLength)
   End If

End Sub

Private Sub UnloadDLLHandler(DE As DEBUG_EVENT, ByVal DbgForm As frmDebugger)
Dim UDLLInfo As UNLOAD_DLL_DEBUG_INFO
   
   ' Get the UNLOAD_DLL_DEBUG_INFO structure
   ' from the byte array
   MoveMemory UDLLInfo, DE.DEBUG_INFO(0), Len(UDLLInfo)
   
   DbgForm.lvwDLLs.ListItems.Remove "#" & CStr(UDLLInfo.lpBaseOfDll)

End Sub

Private Sub mnuFExit_Click()

   Unload Me

End Sub


Private Sub mnuODbgChild_Click()

   mnuODbgChild.Checked = Not mnuODbgChild.Checked

End Sub

Private Sub mnuWArrange_Click()

   Arrange vbArrangeIcons

End Sub

Private Sub mnuWCascade_Click()

   Arrange vbCascade

End Sub

Private Sub mnuWTileH_Click()

   Arrange vbTileHorizontal
   
End Sub


Private Sub mnuWTileV_Click()

   Arrange vbTileVertical

End Sub


