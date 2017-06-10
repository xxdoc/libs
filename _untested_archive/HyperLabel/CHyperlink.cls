VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : CHyperlink
' DateTime  : 04 jan 2006 03:15
' Author    : Joacim Andersson, Brixoft Software, http://www.brixoft.net
' Purpose   : A simple private class that contains location information
'             about a hyperlink in the HyperLabel control
'---------------------------------------------------------------------------------------
Option Explicit

Private m_nX1 As Long
Private m_nX2 As Long
Private m_nY1 As Long
Private m_nY2 As Long
Private m_sHRef As String
Private m_nTextHeight As Long

Public Function PosInside(ByVal X As Long, ByVal Y As Long) As Boolean
    If m_nX2 < 0 Then
        Exit Function
    End If
    If m_nY1 <> m_nY2 Then
        If (X >= m_nX1 And (Y >= m_nY1 And Y <= m_nY1 + m_nTextHeight)) Or _
         (X <= m_nX2 And (Y >= m_nY2 And Y <= m_nY2 + m_nTextHeight)) Then
            PosInside = True
        End If
    ElseIf ((X >= m_nX1 And X <= m_nX2) And _
     (Y >= m_nY1 And Y <= m_nY1 + m_nTextHeight)) Then
        PosInside = True
    End If
End Function

Public Property Get X1() As Long
    X1 = m_nX1
End Property

Public Property Let X1(ByVal nX1 As Long)
    m_nX1 = nX1
End Property

Public Property Get X2() As Long
    X2 = m_nX2
End Property

Public Property Let X2(ByVal nX2 As Long)
    m_nX2 = nX2
End Property

Public Property Get Y1() As Long
    Y1 = m_nY1
End Property

Public Property Let Y1(ByVal nY1 As Long)
    m_nY1 = nY1
End Property

Public Property Get Y2() As Long
    Y2 = m_nY2
End Property

Public Property Let Y2(ByVal nY2 As Long)
    m_nY2 = nY2
End Property

Public Property Get HRef() As String
    HRef = m_sHRef
End Property

Public Property Let HRef(ByVal sHRef As String)
    m_sHRef = sHRef
End Property

Public Property Get TextHeight() As Long
    TextHeight = m_nTextHeight
End Property

Public Property Let TextHeight(ByVal nTextHeight As Long)
    m_nTextHeight = nTextHeight
End Property

Private Sub Class_Initialize()
    X2 = -1
    Y2 = -1
End Sub
