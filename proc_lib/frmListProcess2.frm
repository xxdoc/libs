VERSION 5.00
Begin VB.Form frmListProcess2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Choose Process"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9150
   LinkTopic       =   "frmListProcess"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1140
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   9015
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   3330
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   315
      Left            =   7980
      TabIndex        =   0
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label lblRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Search: "
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   3360
      Width           =   555
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuShowDlls 
         Caption         =   "Show Dlls"
      End
      Begin VB.Menu mnuDumpProcess 
         Caption         =   "Dump Process"
      End
   End
End
Attribute VB_Name = "frmListProcess2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         disassembler functionality provided by olly.dll which
'         is a modified version of the OllyDbg GPL source from
'         Oleh Yuschuk Copyright (C) 2001 - http://ollydbg.de
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'this is an older select process form which only uses a listbox instead of the mscomctl.ocx
'this is the fallback in case the first one fails due to control not found.

Option Explicit

Dim selPid As Long
Dim cpi As New CProcessLib
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Dim baseCaption As String
Dim baseCol As Collection

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Private Function LoadProccesses(c As Collection)

    Dim p As CProcess
    Dim cc As Long, cs As String
    Dim tmp As String
    
    List1.Clear
    
    For Each p In c
        If p.pid > 4 Then
            tmp = rpad(p.pid)
            tmp = tmp & IIf(p.is64Bit, "*64 ", "    ")
        
            cs = InStr(p.User, " ")
            If cs > 0 Then p.User = Mid(p.User, 1, cs)
            
            cc = InStr(p.User, ":")
            If cc > 0 Then p.User = Mid(p.User, cc + 1)
                
            If Len(p.User) > 10 Then p.User = Mid(p.User, 1, 9) & "~"
            tmp = tmp & rpad(p.User, 10)
            
            tmp = tmp & p.fullpath  'can fail on win7?
            List1.AddItem tmp
        End If
    Next
    
End Function

Function SelectProcess(c As Collection) As CProcess
   
    selPid = 0
    Set baseCol = c
    LoadProccesses c
    
    On Error Resume Next
    Me.Show 1
    
    If selPid < 1 Then Exit Function
    Set SelectProcess = GetProcess(selPid)
    Unload Me
    
End Function
 
Private Function GetProcess(pid As Long) As CProcess
     Dim p As CProcess
     
     For Each p In baseCol
        If p.pid = pid Then
            Set GetProcess = p
            Exit Function
        End If
     Next
     
    Set GetProcess = New CProcess
    
End Function

Private Sub Form_Load()
    Dim User As String
    On Error Resume Next
    
    User = cpi.GetProcessUser(GetCurrentProcessId())
    If InStr(User, ":") > 0 Then
        User = Mid(User, InStr(User, ":") + 1)
    End If
    
    If Len(User) > 0 Then Me.Caption = Me.Caption & "   -   Running As: " & User
    Me.Caption = Me.Caption & "   -   SeDebug?: " & cpi.SeDebugEnabled
    baseCaption = Me.Caption
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 500 - Command1.Height
    List2.Move List1.Left, List1.Top, List1.Width, List1.Height
    Command1.Top = Me.Height - Command1.Height - 400
    Command1.Left = Me.Width - Command1.Width - 400
    lblRefresh.Left = Command1.Left - lblRefresh.Width - 400
    lblRefresh.Top = Command1.Top
    txtSearch.Top = Command1.Top
    Label1.Top = Command1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    selPid = 0
End Sub

Private Sub lblRefresh_Click()
     LoadProccesses cpi.GetRunningProcesses
     If List2.Visible Then txtSearch_Change
End Sub

Private Sub List1_Click()
    On Error Resume Next
    Dim str, p As CProcess, a As Long
    selPid = 0
    str = List1.List(List1.ListIndex)
    If Len(str) > 0 Then
        a = InStr(str, " ")
        If a > 0 Then
            selPid = CLng(Mid(str, 1, a)) 'first number
            Set p = GetProcess(selPid)
            Me.Caption = baseCaption & "  - cmdline " & p.cmdLine
        End If
    End If
End Sub

Private Sub List2_Click()
    On Error Resume Next
    Dim str, p As CProcess, a As Long
    selPid = 0
    str = List2.List(List2.ListIndex)
    If Len(str) > 0 Then
        a = InStr(str, " ")
        If a > 0 Then
            selPid = CLng(Mid(str, 1, a)) 'first number
            Set p = GetProcess(selPid)
            Me.Caption = baseCaption & "  - cmdline " & p.cmdLine
        End If
    End If
End Sub

Private Sub List1_DblClick()
    Me.Visible = False
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub List2_DblClick()
    Me.Visible = False
End Sub

Private Sub mnuDumpProcess_Click()
    Dim p As String
    Dim b As Boolean
    Dim cp As CProcess
    
    Set cp = GetProcess(selPid)
    If cp.pid = 0 Then Exit Sub
    
    p = App.path & "\mem.dmp"
    b = cpi.DumpProcess(cp.pid, p)

    MsgBox "saved memory dump? " & b & vbCrLf & vbCrLf & "Output file: " & p

End Sub

Private Sub mnuShowDlls_Click()
    Dim cp As CProcess
    Set cp = GetProcess(selPid)
    If cp.pid = 0 Then Exit Sub
    'frmDlls.ShowDllsFor (cp.pid), Me
End Sub

 

'for listview sorting...
Private Function rpad(v, Optional l As Long = 5)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < l Then
        rpad = v & String(l - X, " ")
    Else
hell:
        rpad = v
    End If
End Function

Private Sub txtSearch_Change()
    Dim i As Long
    
    If Len(txtSearch) = 0 Then
        List2.Visible = False
        Exit Sub
    End If
    
    List2.Clear
    List2.Visible = True
    
    For i = 0 To List1.ListCount - 1
        If InStr(1, List1.List(i), txtSearch, vbTextCompare) > 0 Then
            List2.AddItem List1.List(i)
        End If
    Next
    
End Sub







