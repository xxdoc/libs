VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 768
'********************************************************************************
'
' EventInfo Class - EventCollection Library
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

Private m_sName As String
Attribute m_sName.VB_VarHelpID = 773
Private m_oParams As Parameters
Attribute m_oParams.VB_VarHelpID = 772

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the event name or DISPID."
Attribute Name.VB_HelpID = 771
   Name = m_sName
End Property

Public Property Get Parameters() As Parameters
Attribute Parameters.VB_Description = "Returns the parameters collection."
Attribute Parameters.VB_HelpID = 770
Attribute Parameters.VB_UserMemId = 0
   Set Parameters = m_oParams
End Property

Friend Sub frInitialize( _
   ByVal Name As String, _
   Params() As Variant)
Attribute frInitialize.VB_HelpID = 769
   
   m_sName = Name
   
   Set m_oParams = New Parameters
   m_oParams.frInitialize Params
   
End Sub
