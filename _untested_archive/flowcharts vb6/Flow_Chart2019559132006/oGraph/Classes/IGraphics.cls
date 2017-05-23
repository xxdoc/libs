VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IXGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type LINE_POINT
    X As Double
    Y As Double
End Type

Public Enum eObjectType
    oTLine = 1
    oTPicture = 2
    oTText = 3
    oTConvas = 4
End Enum

Public Enum eLineType
    OnTSuccess = 1
    OnTCompletion = 2
    OnTFail = 3
    OnCustomCol = 4
End Enum


Public Function GetoGraphicsHandle() As Long
'
End Function

Public Function GetoConvasPicBox() As Object
'
End Function

Public Sub RePaintoConvas()
'
End Sub

Public Function IsSelected(ByVal NodeKey As String) As Boolean
'
End Function

Public Function IsoConvasLocked() As Boolean

End Function
