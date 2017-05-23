' $Id: Buffer.cls,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' We use this class to wrap a string so that it can be passed via
' IUnknown as an extra parameter to callback functions.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public stringData As String

