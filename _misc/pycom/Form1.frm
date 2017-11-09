VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
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

Private Sub Form_Load()

    'notes:
    '       run the python script once normally as admin to register the COM objects in registry,
    '
    '       had to update the following reg key manually:
    '           HKLM\SOFTWARE\Classes\CLSID\{41E24E95-D45A-11D2-852C-204C4F4F5020}\InprocServer32
    '           HKLM\SOFTWARE\Classes\CLSID\{41E24E95-D45A-11D2-852C-204C4F4F5021}\InprocServer32
    '           default = C:\Python27\Lib\site-packages\pywin32_system32\pythoncomloader27.dll
    '           it was just the dll name and was giving module not found  in createobject call..
    '
    '       once a python com object is loaded into memory changes to the file
    '       dont seem to take..you have to close the vb6 ide and restart it or just compile/run exe always!
    '
    '       any errors in your python script or callback it just silently dies
    '       you have to debug everything blindly? this includes byval/byref etc not helpful
    '       necessitates msgbox debugigng everything :-\
    
    On Error GoTo hell
        
    Set pyobj = CreateObject("PythonDemos.Utilities")
    dbg "CreateObj(PythonDemos.Utilities) = " & Hex(ObjPtr(pyobj))
    
    'call python method and reveive string array - working
    y = pyobj.SplitString("hello from vb")
    dbg "Received type: " & TypeName(y)
    dbg Join(y, ", ")
    
    'python triggers vb6 callback - working
    rv = pyobj.CallVB(AddressOf myCallBack)
    dbg "Back in vb6 retval = " & rv
    
    'return another python com object class from a python method
    Dim o As Object
    Set o = pyobj.getPyObj
    dbg "pyobj.getPyObj returns type = " & TypeName(o)
    dbg "newObj.test() = " & o.test()
    
    
    dbg "tests complete"
    Set pyobj = Nothing
    
    Exit Sub
hell:
    MsgBox "Caught error: " & Err.Description
    If Not pyobj Is Nothing Then Set pyobj = Nothing
    End
    
End Sub

