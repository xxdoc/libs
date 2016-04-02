VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Elements
'
' Collection of VBSTATSTGs for Storage class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 08/03/1999
'
'*********************************************************************************************
Option Explicit

Dim m_Collection As Collection
Friend Sub Add(SSTG As olelib.StatStg)
Dim SS As StatStg

    Set SS = New StatStg
    
    SS.frInitialize SSTG
    
    m_Collection.Add SS
    
End Sub

Public Function NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = m_Collection.[_NewEnum]
    
End Function

Public Property Get Count() As Long

    Count = m_Collection.Count
    
End Property

Public Property Get Item(ByVal Index As Variant) As StatStg
Attribute Item.VB_UserMemId = 0
    Set Item = m_Collection(Index)
End Property

Private Sub Class_Initialize()

    Set m_Collection = New Collection
    
End Sub



