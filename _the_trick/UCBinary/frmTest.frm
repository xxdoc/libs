VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ctlBinaryData test by The trick"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5040
      Width           =   615
   End
   Begin VB.PictureBox picPreview 
      Height          =   4875
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      Top             =   120
      Width           =   5115
   End
   Begin UCBinData.ctlBinData ctlResources 
      Left            =   720
      Top             =   5100
      _ExtentX        =   1455
      _ExtentY        =   529
      Content         =   "frmTest.frx":0000
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type UUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(0 To 7)       As Byte
End Type

Private Declare Function OleLoadPicture Lib "OleAut32" ( _
                         ByVal pStream As IUnknown, _
                         ByVal lSize As Long, _
                         ByVal fRunMode As Long, _
                         ByRef riid As UUID, _
                         ByRef lplpvObj As Any) As Long
Private Declare Function SHCreateMemStream Lib "Shlwapi" _
                         Alias "#12" ( _
                         ByRef pInit As Any, _
                         ByVal cbInit As Long) As IUnknown

Private m_lCurFile  As Long

Private Sub cmdNext_Click()

    m_lCurFile = (m_lCurFile + 1) Mod ctlResources.FilesCount
    Update
    
End Sub

Private Sub cmdPrev_Click()

    m_lCurFile = m_lCurFile - 1
    If m_lCurFile < 0 Then m_lCurFile = ctlResources.FilesCount - 1
    Update
    
End Sub

Private Sub Update()
    Dim bData()     As Byte
    Dim cStm        As IUnknown
    Dim cPic        As IPicture
    Dim sFiles()    As String
    
    sFiles = ctlResources.FilesList()
    bData = ctlResources.File(sFiles(m_lCurFile))
    Set cStm = SHCreateMemStream(bData(0), UBound(bData) + 1)
    
    If OleLoadPicture(cStm, 0, 0, IID_IPicture, cPic) < 0 Then
        MsgBox "Unable to load picture", vbCritical
    Else
        Set picPreview.Picture = cPic
    End If
    
End Sub

Private Function IID_IPicture() As UUID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
End Function

Private Sub Form_Load()
    Update
End Sub
