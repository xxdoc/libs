VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
     
    Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
    Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
    Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
     
    Private m_lCookie As Long
     
    Private Sub Form_Load()
        List1.AddItem "test"
        List1.AddItem Now
        m_lCookie = PutObject(Me, "MySpecialProject.Form1")
        Shell App.Path & "\2.exe"
    End Sub
     
    Private Sub Form_Unload(Cancel As Integer)
        RevokeObject m_lCookie
    End Sub
     
    Public Function PutObject(oObj As Object, sPathName As String, Optional ByVal Flags As Long) As Long
        Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
        Const IDX_REGISTER  As Long = 3
        Dim hResult         As Long
        Dim pROT            As IUnknown
        Dim pMoniker        As IUnknown
        
        hResult = GetRunningObjectTable(0, pROT)
        If hResult < 0 Then
            Err.Raise hResult, "GetRunningObjectTable"
        End If
        hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
        If hResult < 0 Then
            Err.Raise hResult, "CreateFileMoniker"
        End If
        DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(PutObject)
    End Function
     
    Public Sub RevokeObject(ByVal lCookie As Long)
        Const IDX_REVOKE    As Long = 4
        Dim hResult         As Long
        Dim pROT            As IUnknown
        
        hResult = GetRunningObjectTable(0, pROT)
        If hResult < 0 Then
            Err.Raise hResult, "GetRunningObjectTable"
        End If
        DispCallByVtbl pROT, IDX_REVOKE, lCookie
    End Sub
     
    Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
        Const CC_STDCALL    As Long = 4
        Dim lIdx            As Long
        Dim vParam()        As Variant
        Dim vType(0 To 63)  As Integer
        Dim vPtr(0 To 63)   As Long
        Dim hResult         As Long
        
        vParam = A
        For lIdx = 0 To UBound(vParam)
            vType(lIdx) = VarType(vParam(lIdx))
            vPtr(lIdx) = VarPtr(vParam(lIdx))
        Next
        hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
        If hResult < 0 Then
            Err.Raise hResult, "DispCallFunc"
        End If
    End Function
