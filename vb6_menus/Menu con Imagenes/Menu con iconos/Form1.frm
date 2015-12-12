VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menú con Imágenes"
   ClientHeight    =   3345
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ejemplo 3"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejemplo 2"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "&Abrir"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "G&uardar"
         Index           =   3
         Shortcut        =   ^G
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Guardar &Como"
         Index           =   4
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "&Imprimir"
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Vista Previa"
         Index           =   7
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu SubMnuArchivo 
         Caption         =   "Salir"
         Index           =   9
      End
   End
   Begin VB.Menu MnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "Deshacer"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "Cortar"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "Copiar"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "Pegar"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubMnuEdicion 
         Caption         =   "Seleccionar Todo"
         Index           =   6
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Begin VB.Menu SubMnuVer 
         Caption         =   "Horizontal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu SubMnuVer 
         Caption         =   "Vertical"
         Index           =   1
      End
      Begin VB.Menu SubMnuVer 
         Caption         =   "Cascada"
         Index           =   2
      End
      Begin VB.Menu SubMnuVer 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubMnuVer 
         Caption         =   "Opciones"
         Index           =   4
      End
   End
   Begin VB.Menu MnuFormato 
      Caption         =   "Formato"
      Begin VB.Menu SubMnuFormato 
         Caption         =   "Fuente"
         Index           =   0
         Begin VB.Menu SubMnuFuente 
            Caption         =   "Negrita"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu SubMnuFuente 
            Caption         =   "Cursiva"
            Index           =   1
         End
         Begin VB.Menu SubMnuFuente 
            Caption         =   "Subrayada"
            Index           =   2
         End
      End
      Begin VB.Menu SubMnuFormato 
         Caption         =   "Bloquear"
         Index           =   1
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu SubMnuAyuda 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "Form1"
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

Dim cMenuImage As clsMenuImage


Private Sub Form_Load()

    Dim i As Long
    Set cMenuImage = New clsMenuImage

    With cMenuImage
        
        ' Inicializamos la Clase
        ' El primer parámetro es para indicar el HWnd de la ventana que contiene o llama al menú.
        ' El segundo y tercer parámetro son las dimenciones con las que se mostrarán las imágenes en los menúes.
        ' Hay un cuarto parámetro opcional si es que queremos lanzar un evento con la subclasificación del formulario (reveer).
        .Init Me.HWnd, 16, 16
    
        ' La función AddImageFromStream carga imágenes a una lista interna de la clase desde un array de bits.
        ' Si queremos cargar desde archivo utilizamos la función AddImageFromFile.
        ' Si queremos cargar un ícono que se encuentra en la memoria utilizamos la función AddIconFromHandle.
        For i = 0 To 20
           .AddImageFromStream LoadResData("PNG_" & i, "CUSTOM")
        Next
        
        ' El primer parámetro corresponde a la lista de imágenes internas de la ClsMenuImage.
        ' El segundo parámetro corresponde a la posición del menú (comienza desde cero).
        ' El tercer parámetro corresponde a la posición en la barra de menú (comienza desde cero).
        ' Si existen varios submenú deben pasarse más parámetros según su posición.
        
        '---------'
        ' Archivo '
        '---------'
        .PutImageToVBMenu 0, 0, 0       ' Nuevo
        .PutImageToVBMenu 1, 1, 0       ' Abrir
        .PutImageToVBMenu 2, 3, 0       ' Guardar
        .PutImageToVBMenu 3, 4, 0       ' Guardar como
        .PutImageToVBMenu 4, 6, 0       ' Imprimir
        .PutImageToVBMenu 5, 7, 0       ' Vista previa

        '---------'
        ' Edición '
        '---------'
        .PutImageToVBMenu 6, 0, 1       ' Deshacer
        .PutImageToVBMenu 7, 2, 1       ' Cortar
        .PutImageToVBMenu 8, 3, 1       ' Copiar
        .PutImageToVBMenu 9, 4, 1       ' Pegar
        .PutImageToVBMenu 10, 6, 1      ' Seleccionar todo
                
        '---------'
        '   Ver   '
        '---------'
        .PutImageToVBMenu 11, 0, 2      ' Horizontal
        .PutImageToVBMenu 12, 1, 2      ' Vertical
        .PutImageToVBMenu 13, 2, 2      ' Cascada
        .PutImageToVBMenu 14, 4, 2      ' Opciones
        
        '---------'
        ' Formato '
        '---------'
        .PutImageToVBMenu 15, 0, 3      ' Fuente
        .PutImageToVBMenu 16, 0, 3, 0   ' Negrita       ' Observese que el cuarto parámetro es para indicar que es un submenú de un submenú.
        .PutImageToVBMenu 17, 1, 3, 0   ' Cursiva       ' De existir más submenús dentro de éste deberemos indicárselo con más parámetros.
        .PutImageToVBMenu 18, 2, 3, 0   ' Subrayado     ' Ejemplo: .PutImageToVBMenu 18, 2, 3, 0, 1, 0 Etc.
        .PutImageToVBMenu 19, 1, 3      ' Bloquear
        
        '---------'
        '  Ayuda  '
        '---------'
        .PutImageToVBMenu 20, 0, 4      ' Acerca de

        ' En Windows XP queda mejor si remobemos el style check, ya que éste agrega un margen adicional para las marcas de verificación.
        ' En Windows Vista o Windows 7 esto no es necesario, ya que lo remarca debajo de la imágen.
        If Not .IsWindowVistaOrLater Then
            .RemoveMenuCheckVB 0
            .RemoveMenuCheckVB 1
            .RemoveMenuCheckVB 2
            .RemoveMenuCheckVB 3
            .RemoveMenuCheckVB 3, 0     'Submenú de Submenú.
            .RemoveMenuCheckVB 4
        End If
        
    End With
    
End Sub


Private Sub SubMnuFuente_Click(Index As Integer)
    SubMnuFuente(Index).Checked = Not SubMnuFuente(Index).Checked
End Sub


Private Sub SubMnuVer_Click(Index As Integer)
    Dim i As Long
    If Index < 3 Then
        For i = 0 To 2
            SubMnuVer(i).Checked = False
        Next
        SubMnuVer(Index).Checked = True
    End If
End Sub


Private Sub Command1_Click()
    Form2.Show
End Sub


Private Sub Command2_Click()
    Form3.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cMenuImage = Nothing            ' Descargamos la clase
End Sub
