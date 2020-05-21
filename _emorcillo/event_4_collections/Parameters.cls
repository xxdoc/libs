VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 1024
'********************************************************************************
'
' Parameters Class - EventCollection Library
'
'********************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on:
'            * Windows XP Pro SP1
'            * VB6 SP5
'
' History:
'           01/02/2003 * This code replaces the old EventCollection
'                        class.
'
'********************************************************************************
Option Explicit

Dim m_avParams() As Variant
Attribute m_avParams.VB_VarHelpID = 1030

Friend Sub frInitialize(Params() As Variant)
Attribute frInitialize.VB_HelpID = 1029
   m_avParams = Params
End Sub

Friend Function frGetParams() As Variant()
Attribute frGetParams.VB_HelpID = 1028
   frGetParams = m_avParams
End Function

Public Property Get Item(Index As Variant) As Variant
Attribute Item.VB_Description = "Returns a parameter."
Attribute Item.VB_HelpID = 1027
Attribute Item.VB_UserMemId = 0
   Item = m_avParams(CLng(Index))
End Property

Public Property Let Item(Index As Variant, Value As Variant)
   m_avParams(CLng(Index)) = Value
End Property

Public Property Set Item(Index As Variant, Value As Variant)
   Set m_avParams(CLng(Index)) = Value
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection."
Attribute Count.VB_HelpID = 1026
   
   On Error GoTo IsEmpty
   
   Count = UBound(m_avParams)
   
   Exit Property
   
IsEmpty:
   
   Count = 0
   
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 1025
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   ' Not yet implemented
End Function
