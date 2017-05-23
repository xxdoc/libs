VERSION 5.00
Begin VB.Form fCallback 
   BackColor       =   &H00F0F0F0&
   Caption         =   "Callback sample"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuCallback 
      Caption         =   "&Callback"
      Begin VB.Menu mnuItem 
         Caption         =   "&EnumWindows"
         Index           =   0
      End
      Begin VB.Menu mnuItem 
         Caption         =   "EnumFontFamilies"
         Index           =   1
      End
      Begin VB.Menu mnuItem 
         Caption         =   "SetTimer"
         Index           =   2
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Lightweight Subclassing"
         Index           =   3
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Subclassing and SetTimer simulaneously"
         Index           =   4
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuItem 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "fCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* fCallback - Generic callback sample
'*
'* Warning: some of the included callback thunks aren't 'End' safe. The Enumumeration API's
'*  should be bullet proof, but anything that continues until stopped, like subclassing and
'*  timers, will cause problems if the user clicks or executes End in the IDE. So what does
'*  the bIdeSafety parameter provide (I ask rhetorically)? It supplies a degree of crash
'*  protection. Suppose you've set a timer running and then click End. The callback thunk is
'*  still alive in allocated memory and being called periodically by the timer code in the OS.
'*  The thunk only knows that the IDE has stopped, and that it can't callback to the specified
'*  callback procedure. The thunk can't stop the timer, it has no knowledge about timers, it's
'*  a generic, general-pupose thunk. Subsequently, if the user runs the app in the IDE again,
'*  the original thunk will detect that the IDE is running and resume callback. So you could
'*  find yourself with timer events when you hadn't asked for any, or two timers running
'*  simultaneously.
'*
'* Bottom line: use freely with enumeration api's, both OS provided and third-party (I'm
'*  thinking of zip dll's and the like, where it's common to receive progress updates via a
'*  callback), but be very careful with facilities that need to be explicitly stopped/killed.
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

Private Const LF_FACESIZE As Long = 32

Private Type LOGFONT
  lfHeight                As Long
  lfWidth                 As Long
  lfEscapement            As Long
  lfOrientation           As Long
  lfWeight                As Long
  lfItalic                As Byte
  lfUnderline             As Byte
  lfStrikeOut             As Byte
  lfCharSet               As Byte
  lfOutPrecision          As Byte
  lfClipPrecision         As Byte
  lfQuality               As Byte
  lfPitchAndFamily        As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
  tmHeight                As Long
  tmAscent                As Long
  tmDescent               As Long
  tmInternalLeading       As Long
  tmExternalLeading       As Long
  tmAveCharWidth          As Long
  tmMaxCharWidth          As Long
  tmWeight                As Long
  tmOverhang              As Long
  tmDigitizedAspectX      As Long
  tmDigitizedAspectY      As Long
  tmFirstChar             As Byte
  tmLastChar              As Byte
  tmDefaultChar           As Byte
  tmBreakChar             As Byte
  tmItalic                As Byte
  tmUnderlined            As Byte
  tmStruckOut             As Byte
  tmPitchAndFamily        As Byte
  tmCharSet               As Byte
  ntmFlags                As Long
  ntmSizeEM               As Long
  ntmCellHeight           As Long
  ntmAveWidth             As Long
End Type

Private Type RECT
  Left                    As Long
  Top                     As Long
  Right                   As Long
  Bottom                  As Long
End Type

Private nOriginalWndProc  As Long                                           'SetWindowLong nOriginalWndProc
Private nTimerID          As Long                                           'SetTimer nTimerID
Private nLineHeight       As Long                                           'Height of a line of text
Private rc                As RECT                                           'Scrolling rectangle

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnumFontFamiliesA Lib "gdi32" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Form_Load()
  Me.AutoRedraw = False
  nLineHeight = Me.TextHeight("My") 'Get the height in pixels of a line of text
End Sub

Private Sub Form_Resize()
  rc.Right = Me.ScaleWidth
  rc.Bottom = Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If nOriginalWndProc <> 0 And nTimerID <> 0 Then
    MsgBox "Cancel '" & mnuItem(4).Caption & "' from the 'Callback' menu before closing", vbExclamation
    Cancel = True
    
  ElseIf nOriginalWndProc <> 0 Then
    MsgBox "Cancel '" & mnuItem(3).Caption & "' from the 'Callback' menu before closing", vbExclamation
    Cancel = True
  
  ElseIf nTimerID <> 0 Then
    MsgBox "Cancel '" & mnuItem(2).Caption & "' from the 'Callback' menu before closing", vbExclamation
    Cancel = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cb_Terminate
End Sub

Private Sub mnuItem_Click(Index As Integer)
  Dim nCallback1 As Long
  Dim nCallback2 As Long
  
  Select Case Index
  Case 0
'EnumWindows
    nCallback1 = cb_AddressOf(1, 2) 'Callback ordinal #1 with 2 parameters
    If nCallback1 <> 0 Then
      Me.Cls
      EnumWindows nCallback1, 123
    End If
    
  Case 1
'EnumFontFamilies
    nCallback1 = cb_AddressOf(2, 4) 'Callback ordinal #2 with 4 parameters
    If nCallback1 <> 0 Then
      Me.Cls
      EnumFontFamiliesA Me.hDC, vbNullString, nCallback1, 0
      With Me
        .FontName = "Courier New"
        .FontBold = False
        .FontItalic = False
        .FontSize = 9
        .FontStrikethru = False
        .FontUnderline = False
      End With
      nLineHeight = Me.TextHeight("My")
    End If
    
  Case 2
'Timer
    If mnuItem(2).Checked Then
      KillTimer 0, nTimerID
      nTimerID = 0
      mnuItem(0).Enabled = True: mnuItem(1).Enabled = True: mnuItem(3).Enabled = True: mnuItem(4).Enabled = True
      mnuItem(2).Checked = False
    Else
      nCallback1 = cb_AddressOf(3, 4) 'Callback ordinal #3 with 4 parameters
      If nCallback1 <> 0 Then
        MsgBox "Warning: the Timer callback isn't 'End' safe.", vbInformation
        Me.Cls
        nTimerID = SetTimer(0, 0, 100, nCallback1)
        mnuItem(0).Enabled = False: mnuItem(1).Enabled = False: mnuItem(3).Enabled = False: mnuItem(4).Enabled = False
        mnuItem(2).Checked = True
      End If
    End If
    
  Case 3
'Lightweight subclassing
    If mnuItem(3).Checked Then
      SetWindowLong Me.hWnd, -4, nOriginalWndProc
      nOriginalWndProc = 0
      mnuItem(0).Enabled = True: mnuItem(1).Enabled = True: mnuItem(2).Enabled = True: mnuItem(4).Enabled = True
      mnuItem(3).Checked = False
    Else
      nCallback1 = cb_AddressOf(4, 4) 'Callback ordinal #4 with 4 parameters
      If nCallback1 <> 0 Then
        MsgBox "Warning: the 'Lightweight subclassing' callback isn't 'End' safe.", vbInformation
        Me.Cls
        nOriginalWndProc = SetWindowLong(Me.hWnd, -4, nCallback1)
        mnuItem(0).Enabled = False: mnuItem(1).Enabled = False: mnuItem(2).Enabled = False: mnuItem(4).Enabled = False
        mnuItem(3).Checked = True
      End If
    End If
  
  Case 4
'Subclassing and SetTimer simultaneously
    If mnuItem(4).Checked Then
      SetWindowLong Me.hWnd, -4, nOriginalWndProc
      KillTimer 0, nTimerID
      nTimerID = 0
      nOriginalWndProc = 0
      mnuItem(0).Enabled = True: mnuItem(1).Enabled = True: mnuItem(2).Enabled = True: mnuItem(3).Enabled = True
      mnuItem(4).Checked = False
    Else
      nCallback1 = cb_AddressOf(4, 4, 0) 'Callback ordinal #4 with 4 parameters, thunk #0
      nCallback2 = cb_AddressOf(3, 4, 1) 'Callback ordinal #3 with 4 parameters, thunk #1
      If nCallback1 <> 0 And nCallback2 Then
        MsgBox "Warning: the 'Subclassing and SetTimer simultaneously' callback isn't 'End' safe.", vbInformation
        Me.Cls
        nOriginalWndProc = SetWindowLong(Me.hWnd, -4, nCallback1)
        nTimerID = SetTimer(0, 0, 100, nCallback2)
        mnuItem(0).Enabled = False: mnuItem(1).Enabled = False: mnuItem(2).Enabled = False: mnuItem(3).Enabled = False
        mnuItem(4).Checked = True
      End If
    End If

  Case 6
    Unload Me
    
  End Select
End Sub

Private Sub Display(ByVal sText As String)
  Const SW_INVALIDATE As Long = &H2
  
  ScrollWindowEx Me.hWnd, 0, -nLineHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE
  UpdateWindow Me.hWnd
  Me.CurrentY = Me.ScaleHeight - nLineHeight
  Print sText
End Sub

Private Function Hfmt(ByVal nValue As Long) As String
  Hfmt = Right$("0000000" & Hex$(nValue), 8)
End Function

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

'Callback ordinal 4
Private Function WndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_PAINT = &HF
  
  WndProc = CallWindowProcA(nOriginalWndProc, lng_hWnd, uMsg, wParam, lParam)
  
  If uMsg <> WM_PAINT Then
    Display "hWnd: " & Hfmt(lng_hWnd) & ", " & _
            "uMsg: " & Hfmt(uMsg) & ", " & _
            "wParam: " & Hfmt(wParam) & ", " & _
            "lParam: " & Hfmt(lParam) & ", " & _
            "Return: " & Hfmt(WndProc)
  End If
End Function

'Callback ordinal 3
Private Function TimerProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
  Display "Timer: GetTicks = " & dwTime
End Function

'Callback ordinal 2
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, ByVal lParam As Long) As Long
  Dim FaceName As String
  
  FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  FaceName = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
  
  With Me
    .FontName = FaceName
    .FontBold = False
    .FontItalic = False
    .FontSize = 10
    .FontStrikethru = False
    .FontUnderline = False
  End With
  
  nLineHeight = Me.TextHeight("My")
  Display FaceName
  
  EnumFontFamProc = 1 'Continue
End Function

'Callback ordinal 1
Private Function EnumWindowsProc(ByVal lng_hWnd As Long, ByVal lParam As Long) As Long
  Dim nLen     As Long
  Dim sCaption As String
  
  sCaption = Space$(256)
  nLen = GetWindowTextA(lng_hWnd, sCaption, 255)
  
  If nLen > 0 Then
    Display Left$(sCaption, nLen)
  End If
  
  EnumWindowsProc = 1 'Continue
End Function
