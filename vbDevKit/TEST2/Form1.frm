VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim fso As New CFileSystem3
'    Dim c As Collection
'
'    Set c = fso.GetSubFolders("C:\Program Files", , True)
'
'
'    For Each f In c
'        Debug.Print f
'    Next
'    Debug.Print c.Count
    
    'Debug.Print fso.GetShortName("C:\Documents and Settings\david\Desktop\doesnotexistyet.txt")
    
    'Debug.Print fso.dlg.FolderDialog2()
    
    'If fso.FolderExists("c:\windows") Then MsgBox 1
    
    Set c = fso.dlg.OpenMulti()
    For Each f In c
        Debug.Print f
    Next
    
    End
    
End Sub
