VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' $Id: Buffer.cls,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' We use this class to wrap a string so that it can be passed via
' IUnknown as an extra parameter to callback functions.

Public stringData As String