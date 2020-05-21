VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 512
'********************************************************************************
'
' ObjectInfo Class - EventCollection Library
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

Private m_oObject As olelib.IUnknown
Attribute m_oObject.VB_VarHelpID = 523
Private m_sKey As String
Attribute m_sKey.VB_VarHelpID = 522
Private m_lCookie As Long
Attribute m_lCookie.VB_VarHelpID = 521
Private m_lIndex As Long
Attribute m_lIndex.VB_VarHelpID = 520
Private m_tIID As UUID
Attribute m_tIID.VB_VarHelpID = 519

Friend Property Get IID() As UUID
Attribute IID.VB_HelpID = 518
   IID = m_tIID
End Property

Friend Property Get Cookie() As Long
Attribute Cookie.VB_HelpID = 517
   Cookie = m_lCookie
End Property

Public Property Get Key() As String
Attribute Key.VB_HelpID = 516
   Key = m_sKey
End Property

Public Property Get Object() As Object
Attribute Object.VB_HelpID = 515
   Set Object = m_oObject
End Property

Public Property Get Index() As Long
Attribute Index.VB_HelpID = 514
   Index = m_lIndex
End Property

Friend Sub frInitialize( _
   ByVal Key As String, _
   ByVal Index As Long, _
   ByVal Cookie As Long, _
   ByVal Object As Object, _
   IID As UUID)
Attribute frInitialize.VB_HelpID = 513

   m_sKey = Key
   m_lIndex = Index
   m_lCookie = Cookie
   Set m_oObject = Object
   m_tIID = IID
   
End Sub
