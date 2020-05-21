SendToMailRecipient

'-----------------------------------------------------------------
' Procedure : SendToMailRecipient
' Purpose   : Simulates a drop operation to
'             "Sent To/Mail Recipient" shell extension
'-----------------------------------------------------------------
'
Public Sub SendToMailRecipient( _
   ByVal Filename As String)
Dim tIID_IDropTarget As UUID
Dim tCLSID_SendMail As UUID
Dim oSendMail As IDropTarget
Dim oDO As IDataObject
Dim lRes As Long
   
   ' Initialize interface of IDropTarget
   CLSIDFromString "{00000122-0000-0000-C000-000000000046}", _
                   tIID_IDropTarget
   
   ' Initialize CLSID of ".MAPIMail"
   CLSIDFromString "{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}", _
                   tCLSID_SendMail
   
   ' Create the "SendTo/Mail Recipient" object
   lRes = CoCreateInstance(tCLSID_SendMail, _
                  Nothing, CLSCTX_INPROC_SERVER, _
                  tIID_IDropTarget, _
                  oSendMail)
   
   If lRes = S_OK Then
   
      ' Get the file IDataObject interface
      Set oDO = GetFileDataObject(Filename)
      
      ' Simulate the drop operation
      oSendMail.DragEnter oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
      oSendMail.Drop oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
   
   Else
      Err.Raise lRes
   End If

End Sub

'--------------------------------------------------------------
' Procedure : GetFileDataObject
' Purpose   : Returns the IDataObject interface for a file
'--------------------------------------------------------------
'
Private Function GetFileDataObject( _
   ByVal Filename As String) As IDataObject
Dim tIID_IDataObject As UUID
Dim tIID_IShellFolder As UUID
Dim oDesktop As IShellFolder
Dim oParent As IShellFolder
Dim oUnk As IUnknown
Dim sFolder As String
Dim lPidl As Long
Dim lPtr As Long

   ' Intialize IDs
   CLSIDFromString "{0000010e-0000-0000-C000-000000000046}", _
                   tIID_IDataObject
   CLSIDFromString IIDSTR_IShellFolder, tIID_IShellFolder
   
   sFolder = Left$(Filename, InStrRev(Filename, "\") - 1)
   Filename = Mid$(Filename, Len(sFolder) + 2)
   If Right$(sFolder, 1) = ":" Then sFolder = sFolder + "\"
   
   ' Get the parent folder object
   Set oDesktop = SHGetDesktopFolder
   
   ' Get the parent folder IDL
   oDesktop.ParseDisplayName 0, 0, StrPtr(sFolder), lPtr, lPidl, 0
   
   ' Get the parent folder object
   oDesktop.BindToObject lPidl, 0, tIID_IShellFolder, lPtr
   MoveMemory oParent, lPtr, 4&
   
   ' Release the PIDL
   CoTaskMemFree lPidl
   
   ' Get the file PIDL
   oParent.ParseDisplayName 0, 0, StrPtr(Filename), 0, lPidl, 0
   
   ' Get the file IDataObject
   lPtr = oParent.GetUIObjectOf(0, 1, lPidl, tIID_IDataObject, 0)
   MoveMemory oUnk, lPtr, 4&
   
   ' Release the file PIDL
   CoTaskMemFree lPidl

   ' Return the file IDataObject
   Set GetFileDataObject = oUnk
   
End Function
