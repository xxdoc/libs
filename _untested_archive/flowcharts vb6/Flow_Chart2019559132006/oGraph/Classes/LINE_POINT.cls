VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LINE_POINT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Type LINE_POINT
'    x As Double
'    y As Double
'End Type
Private m_X As Double
Private m_Y As Double

Public Property Get X() As Double
    X = m_X
End Property

Public Property Get Y() As Double
    Y = m_Y
End Property


Public Property Let X(ByVal pX As Double)
    m_X = Round(pX, 2)
End Property

Public Property Let Y(ByVal pY As Double)
    m_Y = Round(pY, 2)
End Property

