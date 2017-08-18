VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Show(pInitialColour As Long) As Long
   fColourPicker.ucColourPicker.InitialColour = pInitialColour
   fColourPicker.WindowState = vbNormal
   fColourPicker.Show 1
   Show = fColourPicker.ucColourPicker.ChosenColour
End Function
Public Property Get Cancelled() As Boolean
   Cancelled = fColourPicker.Cancelled
End Property
Private Sub Class_Terminate()
   Unload fColourPicker
End Sub


