VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Storage sample"
   ClientHeight    =   4230
   ClientLeft      =   2400
   ClientTop       =   3165
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   Begin VB.PictureBox picImage 
      Height          =   2055
      Index           =   1
      Left            =   3735
      ScaleHeight     =   1995
      ScaleWidth      =   3645
      TabIndex        =   2
      ToolTipText     =   "Double click to load a picture file"
      Top             =   2130
      Width           =   3705
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2910
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picImage 
      Height          =   2055
      Index           =   0
      Left            =   3750
      ScaleHeight     =   1995
      ScaleWidth      =   3645
      TabIndex        =   1
      ToolTipText     =   "Double click to load a picture file"
      Top             =   15
      Width           =   3705
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Type some text."
      Top             =   45
      Width           =   3705
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVFont 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Structured Storage Sample Program
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 07/31/1999
' Updates:
'           08/12/1999. ReadPict now uses IPersistStream to read the picture.
'
'*********************************************************************************************

Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim m_FileName As String
Dim m_Changed As Boolean
Private Sub OpenFile(ByVal FileName As String)
Dim File As Storage, Data As Stream, Stg As Storage
        
    ' Clear previous text and image
    txtText.Text = ""
    Set picImage(0).Picture = Nothing
    Set picImage(1).Picture = Nothing
    
    ' Open the structured storage file
    Set File = OpenStorageFile(FileName)
    
    ' Open TextBox stream
    Set Data = File.OpenStream("TextBox")
    
    ' Read the text
    txtText.Text = Data.ReadData(vbString)
    
    ' Read the font
    Set txtText.Font = Data.ReadObject
    
    ' Get form data
    Set Data = File.OpenStream("Form")

    ' Save form data
    Me.WindowState = 0
    Me.Move Data.ReadData(vbSingle), Data.ReadData(vbSingle), Data.ReadData(vbSingle), Data.ReadData(vbSingle)

    ' Open Pictures storage
    Set Stg = File.OpenStorage("Pictures")
     
    Set Data = Stg.OpenStream("Index 0")
    Set picImage(0).Picture = Data.ReadObject
    
    Set Data = Stg.OpenStream("Index 1")
    Set picImage(1).Picture = Data.ReadObject

    m_Changed = False
    m_FileName = FileName
    
End Sub

Private Sub SaveFile(ByVal FileName As String)
Dim File As Storage, FileProps As DocProperties
Dim Data As Stream, Stg As Storage
Dim UN As String * 260
             
   On Error Resume Next
    
   ' Create storage file
   Set File = CreateStorageFile(FileName)
    
   If Err.Number = 58 Then
      
      Err.Clear
      
      Set File = OpenStorageFile(FileName)
      
      If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical
         Exit Sub
      End If
      
   End If
   
    ' Create a new DocProperties object
    Set FileProps = New DocProperties

    ' Bind properties to storage file
    FileProps.BindToStorage File
    
    ' Get the current logged user name
    GetUserName UN, Len(UN)
    
    ' Write properties
    With FileProps
        .Application = "Edanmo's VB Structured Storage Sample Application"
        If .Author = "" Then .Author = Left$(UN, InStr(UN, vbNullChar))
        .Title = "Storage File Sample"
        .LastSavedBy = Left$(UN, InStr(UN, vbNullChar))
        .Comments = "This sample file contains text and graphics."
        .Revision = CStr(Val(.Revision) + 1)
        '.SetPropertyByName odpDocSummary, PID_DOCPARTS, Array("Text", "Picture")
    End With
    
    ' Create a storage to store the textbox
    ' text and font
    Set Data = File.CreateStream("TextBox", sfCreate Or sfReadWrite Or sfShareExclusive)
    
    ' Create a stream to save the text
    ' within the TextBox storage
    
    ' Save the text
    Data.WriteData txtText.Text
    
    ' Save the font
    Data.WriteObject txtText.Font
    
    ' Create a storage to store the
    ' pictures
    Set Stg = File.CreateStorage("Pictures", sfCreate Or sfReadWrite Or sfShareExclusive)
     
    ' Create a stream within "Picture" storage
    ' and let the Picture property save the
    ' image
    Set Data = Stg.CreateStream("Index 0", sfCreate Or sfReadWrite Or sfShareExclusive)
    Data.WriteObject picImage(0)
    
    Set Data = Stg.CreateStream("Index 1", sfCreate Or sfReadWrite Or sfShareExclusive)
    Data.WriteObject picImage(1)

    ' Create another stream
    Set Data = File.CreateStream("Form", sfCreate Or sfReadWrite Or sfShareExclusive)

    ' Save form data
    Data.WriteData Me.Left
    Data.WriteData Me.Top
    Data.WriteData Me.Width
    Data.WriteData Me.Height
    
    Data.WriteData "Hi Mom!"
    
    Dim S As String
    S = "Hello"
    Data.WriteData S
    ' Force storage object
    ' to write changes
    File.Commit
    
    m_Changed = False

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    txtText.Move 0, 0, ScaleWidth / 2 - 2, ScaleHeight
    picImage(0).Move ScaleWidth / 2 + 1, 0, txtText.Width, ScaleHeight / 2 - 2
    picImage(1).Move picImage(0).Left, ScaleHeight / 2 + 1, txtText.Width, picImage(0).Height
    
End Sub


Private Sub mnuFExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFNew_Click()

    If m_Changed Then
        If MsgBox("Do you want to save the changes?", vbYesNo Or vbQuestion) = vbYes Then
            mnuFSave_Click
        End If
    End If
    
    m_FileName = ""
    m_Changed = False
    
    ' Clear the text
    txtText.Text = ""
    
    ' Reset the font
    With txtText.Font
        .Bold = False
        .Italic = False
        .Name = "Courier New"
        .Strikethrough = False
        .Underline = False
        .Size = 10
    End With
    
    ' Clear the pictures
    Set picImage(0).Picture = Nothing
    Set picImage(1).Picture = Nothing
    
End Sub


Private Sub mnuFOpen_Click()

    If m_Changed Then
        If MsgBox("Do you want to save the changes?", vbYesNo Or vbQuestion) = vbYes Then
            mnuFSave_Click
        End If
    End If
    
    On Error Resume Next
    
    With CommonDialog1
        .DialogTitle = "Open storage file"
        .Filter = "Storage Files|*.stg"
        .DefaultExt = "stg"
        .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
        .ShowOpen
    End With
    
   If Err.Number = 0 Then
      
      OpenFile CommonDialog1.FileName
      
      If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical
      End If
      
   End If

End Sub

Private Sub mnuFSave_Click()

    On Error Resume Next
    
    If m_FileName = "" Then
    
        With CommonDialog1
            .DialogTitle = "Save storage file"
            .Filter = "Storage Files|*.stg"
            .DefaultExt = "stg"
            .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist
            .ShowSave
        End With
    
        If Err.Number = 0 Then
            m_FileName = CommonDialog1.FileName
        Else
            Exit Sub
        End If
        
    End If
    
    On Error GoTo 0
    
    SaveFile m_FileName
    
End Sub



Private Sub mnuFSaveAs_Click()
    
    m_FileName = ""
    mnuFSave_Click
    
End Sub

Private Sub mnuVFont_Click()

    On Error Resume Next
    
    With CommonDialog1
        .Flags = cdlCFPrinterFonts Or cdlCFScreenFonts
        .ShowFont
    
        If Err.Number = 0 Then
            txtText.FontName = .FontName
            txtText.FontSize = .FontSize
            txtText.FontItalic = .FontItalic
            txtText.FontStrikethru = .FontStrikethru
            txtText.FontUnderline = .FontUnderline
        End If
        
    End With
    
End Sub

Private Sub picImage_DblClick(Index As Integer)

    On Error Resume Next
    
    With CommonDialog1
        .DefaultExt = "bmp"
        .DialogTitle = "Open image"
        .Filter = "Images|*.bmp;*.wmf;*.ico;*.gif;*.jpg"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .FileName = ""
        .ShowOpen
    End With
        
    If Err.Number = 0 Then
       
        Set picImage(Index).Picture = Stg.LoadPicture(CommonDialog1.FileName)
                
        m_Changed = True
        
    End If
    
End Sub


Private Sub txtText_Change()

    m_Changed = True
    
End Sub


