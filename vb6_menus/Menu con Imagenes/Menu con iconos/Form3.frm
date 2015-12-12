VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Ejemplo Con Apis"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form3"
   ScaleHeight     =   3330
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Click en el form"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ------------------------------------------------------
' Autor:    Leandro I. Ascierto
' Fecha:    17 de Julio de 2010
' Web:      www.leandroascierto.com.ar
' ------------------------------------------------------

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal HWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Const MF_CHECKED = &H8&
Const MF_APPEND = &H100&
Const TPM_LEFTALIGN = &H0&
Const MF_DISABLED = &H2&
Const MF_GRAYED = &H1&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const TPM_RETURNCMD = &H100&
Const TPM_RIGHTBUTTON = &H2&

Private cMenuImage As clsMenuImage


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hMenu       As Long
    Dim Pt          As POINTAPI
    Dim lRet        As Long
    Dim hIcon       As Long
    Dim i           As Long
    
    hMenu = CreatePopupMenu()
    AppendMenu hMenu, MF_STRING, 1, "Hello !"
    AppendMenu hMenu, MF_GRAYED Or MF_DISABLED, 2, "Testing ..."
    AppendMenu hMenu, MF_SEPARATOR, 3, ByVal 0&
    AppendMenu hMenu, MF_CHECKED, 4, "TrackPopupMenu"
    
    Set cMenuImage = New clsMenuImage
    
    With cMenuImage
        .Init Me.HWnd, 16, 16
    
        For i = 20 To 24
            ExtractIconEx "shell32.dll", i, ByVal 0&, hIcon, 1
            .AddIconFromHandle hIcon        ' Agregamos íconos desde la dll shell 32.
            DestroyIcon hIcon
        Next
        
        ' Cuando es un menú creado por Apis o tengamos el Handle del menú podemos utilizar la función PutImageToApiMenu
        ' El primer parámetro es el ID de la imágen.
        ' El segundo parámetro es el Handle del menú.
        ' El tercer parámetro es la posición del item en el menú.
        ' A diferencia de la función PutImageToVBMenu éste no contiene un array de parámetros ya que pasamos directamente el menú o submenú con el que queremos trabajar.
        
        .PutImageToApiMenu 0, hMenu, 0
        .PutImageToApiMenu 1, hMenu, 1
        .PutImageToApiMenu 2, hMenu, 3
        
        ' Remueve el style check del hmenu que pasamos.
        If Not .IsWindowVistaOrLater Then
            .RemoveMenuCheckApi hMenu
        End If
    
    End With

    GetCursorPos Pt
    lRet = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
    
    DestroyMenu hMenu
    
    Set cMenuImage = Nothing        ' Descargamos la clase (no es necesario llamar a cMenuImage.Clear)
    
    Debug.Print lRet
    
End Sub
