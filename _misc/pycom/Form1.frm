VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2985
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3780
      Width           =   10140
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   10050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const wMsgBox As Boolean = False

Function dbg(x)
    If wMsgBox Then MsgBox x
    List1.AddItem x
End Function

Function reload(moduleName As String) As String
    On Error Resume Next
    Dim cmd As String
    cmd = Replace("import %1;reload(%1)", "%1", moduleName)
    Set o = CreateObject("PythonDemos.Utilities")
    o.exec cmd
    Text1 = Replace(Err.Description, vbLf, vbCrLf)
End Function

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Public Sub setText(x)
    Text1.Text = x
End Sub

Private Sub Form_Load()

    'notes:
    '       run the PythonDemos.py once normally as admin to register the COM objects in registry,
    '
    '       had to update the following reg key manually:
    '           HKLM\SOFTWARE\Classes\CLSID\{41E24E95-D45A-11D2-852C-204C4F4F5020}\InprocServer32
    '           HKLM\SOFTWARE\Classes\CLSID\{41E24E95-D45A-11D2-852C-204C4F4F5021}\InprocServer32
    '           default = C:\Python27\Lib\site-packages\pywin32_system32\pythoncomloader27.dll
    '           it was just the dll name and was giving module not found  in createobject call..
    '
    '       FIXED:
    '         once a python com object is loaded into memory changes to the file
    '         dont seem to take..you have to close the vb6 ide and restart it or just compile/run exe always!
    '           'https://mail.python.org/pipermail/python-win32/2003-May/001013.html
                'https://github.com/SublimeText/Pywin32/blob/master/lib/x32/win32com/servers/interp.py
                'https://stackoverflow.com/questions/6946376/how-to-reload-a-class-in-python-shell
                'https://docs.python.org/2/library/functions.html
    '
    '       any errors in your python script or callback it just silently dies
    '       you have to debug everything blindly? this includes byval/byref etc not helpful
    '       necessitates msgbox debugigng everything :-\
    '
    '       ref: http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/ch12.htm
    
    On Error GoTo hell
    Dim o As Object, pyobj As Object
    
    If IsIde() And GetModuleHandle("python27.dll") <> 0 Then
        dbg "found python in memory still..trying to reload module in case there were changes..."
        reload "PythonDemos"
    End If
    
    Set pyobj = CreateObject("PythonDemos.Utilities")
    dbg "CreateObj(PythonDemos.Utilities) = " & Hex(ObjPtr(pyobj))
    
    'test a property get/set
    dbg "pyObj.Name = " & pyobj.Name
    pyobj.Name = "Good snake"
    dbg "pyObj.Name = " & pyobj.Name
    
    'call python method and receive string array - working
    y = pyobj.SplitString("hello from vb")
    dbg "Received type: " & TypeName(y)
    dbg Join(y, ", ")
    
    'python triggers vb6 callback - working
    rv = pyobj.CallVB(AddressOf myCallBack)
    dbg "Back in vb6 retval = " & rv
    
    'return another python com object class from a python method
    Set o = pyobj.getPyObj
    dbg "pyobj.getPyObj returns type = " & TypeName(o)
    dbg "newObj.test() = " & o.test()
    
    'pass in a VB COM object to python, have it set our lower text box
    'through Sub Me.setText(x) then Msgbox(me.Text1.Text)
    Call pyobj.useVbObj(Me)
    
    dbg "tests complete"
    Set pyobj = Nothing
    
    Exit Sub
hell:
    If wMsgBox Then MsgBox Err.Description
    Text1 = Replace(Err.Description, vbLf, vbCrLf)
    If Not pyobj Is Nothing Then Set pyobj = Nothing
    'End
    
End Sub

