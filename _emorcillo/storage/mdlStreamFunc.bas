Attribute VB_Name = "mdlStreamFuncs"
Option Explicit

Type SAFEARRAYBOUND
   cElements As Long
   lLBound As Long
End Type

Declare Function SafeArrayCreate Lib "oleaut32" ( _
      ByVal vt As Integer, _
      ByVal cDims As Long, _
      rgsabound As SAFEARRAYBOUND) As Long
'
' WriteArray
'
' Writes an array to the stream object
'
Sub WriteArray(ByVal oStrm As olelib.IStream, ByVal Value As Variant)
Dim lLen As Long
Dim lSAPtr As Long
Dim iDims As Integer
Dim lElemLen As Long
Dim lElems As Long
Dim lLBound As Long
Dim lDataSize As Long
Dim lDataPtr As Long
Dim lIdx As Long

Dim lPtr As Long
Dim oObj As IUnknown
Dim vVar As Variant

   lSAPtr = VarPtr(Value) + 8

   ' Get the pointer to the SAFEARRAY structure
   MoveMemory lSAPtr, ByVal lSAPtr, 4

   ' Check if the array was initialized
   If lSAPtr = 0 Then
      
      ' Write 0 dimensions
      oStrm.Write 0, 2
      
      Exit Sub
      
   End If
   
   ' Get the SAFEARRAY data
   MoveMemory iDims, ByVal lSAPtr, 2
   MoveMemory lElemLen, ByVal lSAPtr + 4, 4
   MoveMemory lDataPtr, ByVal lSAPtr + 12, 4

   ' Calculate SAFEARRAY size
   lLen = 16 + (8 * iDims)

   ' Write the SAFEARRAY structure
   oStrm.Write ByVal lSAPtr, lLen

   ' Calculate the memory used
   ' by the array items
   For lIdx = 0 To iDims - 1
      ' Get element count
      MoveMemory lElems, ByVal lSAPtr + 16 + 8 * lIdx, 4
      lDataSize = lDataSize + lElems
   Next
   lDataSize = lDataSize * lElemLen

   Select Case VariantType(Value) And Not vbArray

      Case vbString     ' Array of strings

         For lIdx = 0 To lDataSize - 1 Step 4

            ' Get the pointer to the string
            MoveMemory lPtr, ByVal lDataPtr + lIdx, 4

            ' Get the string length
            lLen = lstrlenW(lPtr) * 2

            ' Write the string length
            oStrm.Write VarPtr(lLen), 4

            ' Write the string
            oStrm.Write lPtr, lLen

         Next

      Case vbVariant

         For lIdx = 0 To lDataSize - 1 Step 16

            ' Copy the variant from the pointer
            MoveMemory vVar, ByVal lDataPtr + lIdx, 16

            ' Write the variant type
            oStrm.Write VariantType(vVar), 2
            
            ' Write the variant
            WriteValue oStrm, vVar, VariantType(vVar)

         Next

         ' Clear the temporary variant
         MoveMemory vVar, vbEmpty, 2

      Case vbObject, VT_UNKNOWN

         For lIdx = 0 To lDataSize - 1 Step 4

            ' Get the object pointer
            MoveMemory lPtr, ByVal lDataPtr + lIdx, 4

            ' Get a reference to the object
            MoveMemory oObj, lPtr, 4

            ' Write the object
            WriteObject oStrm, oObj

         Next

         ' Release the object reference
         MoveMemory oObj, 0&, 4

      Case Else

         ' Numeric data
         ' write the entire array
         oStrm.Write lDataPtr, lDataSize

   End Select

End Sub


'
' VariantType
'
' Returns the variant type. This function
' is used instead of VarType because VarType
' evaluates default properties on objects.
'
Function VariantType(Var As Variant) As Integer
   MoveMemory VariantType, Var, 2
End Function
'
' WriteNumber
'
' Writes a numeric value to the stream
'
' Parameters:
'
' Value   - The numeric value to write
' VarType - The variant type of the value
'
Sub WriteNumber(ByVal oStrm As olelib.IStream, Value As Variant, VarType As Integer)
Dim lPropLen As Long

   ' Set the data length
   Select Case VarType
      Case vbByte
         lPropLen = 1
      Case vbInteger, vbBoolean
         lPropLen = 2
      Case vbLong, vbSingle
         lPropLen = 4
      Case vbDate, vbDouble, vbCurrency
         lPropLen = 8
   End Select

   ' Write the data
   oStrm.Write ByVal VarPtr(Value) + 8, lPropLen

End Sub

'
' WriteObject
'
' Writes an object to the stream
'
Sub WriteObject(ByVal oStrm As olelib.IStream, ByVal Obj As IUnknown)
Dim CLSID As UUID
Dim aBytes() As Byte
Dim oPPB As IPersistPropertyBag
Dim oPSI As IPersistStreamInit
Dim oPS As IPersistStream
Dim oRecordset As Object

   If Obj Is Nothing Then

      ' Write the empty CLSID
      WriteClassStm oStrm, CLSID

   Else

      On Error Resume Next
      
      ' Query for IPersistStream
      Set oPS = Obj

      If Err.Number = 0 Then

         ' Get CLSID
         oPS.GetClassID CLSID
            
         ' Write the CLSID
         WriteClassStm oStrm, CLSID
         
'         If IsEqualGUID(CLSID_ADODBRecordset, CLSID) Then
'
'
'            ' Get the IDispatch interface
'            Set oRecordset = Obj
'
'            ' Save the recordset
'            oRecordset.Save oStrm
'
'         Else
            
            ' Save the object
            oPS.Save oStrm, 0
            
'         End If

      Else

         Err.Clear

         ' Query for IPersistStreamInit
         Set oPSI = Obj

         If Err.Number = 0 Then
         
            ' Get CLSID
            oPSI.GetClassID CLSID

            ' Write the CLSID
            WriteClassStm oStrm, CLSID

            ' Save the object
            oPSI.Save oStrm, 0

         Else

            Err.Clear

            ' Query for IPersistPropertyBag
            Set oPPB = Obj

            If Err.Number = 0 Then

'               ' Get the object's CLSID
'               oPPB.GetClassID CLSID
'
'               ' Write the CLSID
'               WriteClassStm oStrm, CLSID
'
'               ' Create a temporary proprety bag
'               Set m_TempPB = New PropBag.PropertyBag
'
'               ' Save the object
'               oPPB.Save Me, 0, 1
'
'               ' Get the properties as a byte array
'               aBytes = m_TempPB.Contents
'
'               ' Destroy the property bag
'               Set m_TempPB = Nothing
'
'               ' Write the array lenght
'               oStrm.Write VarPtr(UBound(aBytes) + 1), 4
'
'               ' Write the array
'               oStrm.Write VarPtr(aBytes(0)), UBound(aBytes) + 1

            Else
   
               On Error GoTo 0

               ' This error will be raised only
               ' if the object is contained in an
               ' array and is not persistable
               Err.Raise 5, , "The object class isn't persistable: " & TypeName(Obj)

            End If

         End If

      End If

   End If

   If Err.Number <> 0 Then
      
      On Error GoTo 0
      Err.Raise vbObjectError Or 1, , "The object can't be saved."
      
   End If
   
End Sub

'
' WriteString
'
' Writes a string to the stream
'
Sub WriteString(ByVal oStrm As olelib.IStream, ByVal Value As String)

   ' Write the string
   oStrm.Write ByVal StrPtr(Value) - 4, LenB(Value) + 4

End Sub

'
' WriteValue
'
' Writes a value to the stream
'
Sub WriteValue(ByVal oStrm As olelib.IStream, Value As Variant, VarType As Integer)
Dim vValue As Variant

   If VarType And VT_BYREF Then
      
      ' Remove the VT_BYREF flag
      VarType = VarType And Not VT_BYREF
      
      ' Get the pointed value
      VariantCopyInd vValue, Value
      
   Else
      vValue = Value
   End If
   
   Select Case VarType

      Case vbObject, VT_UNKNOWN
         WriteObject oStrm, vValue

      Case vbString
         WriteString oStrm, vValue

      Case vbDecimal
         oStrm.Write vValue, 16

      Case vbVariant
         WriteValue oStrm, vValue, VarType

      Case Else

         If (VarType And vbArray) = vbArray Then
            WriteArray oStrm, vValue
         Else
            WriteNumber oStrm, vValue, VarType
         End If

   End Select

End Sub
'
' ReadArray
'
' Reads an array from the stream object
'
Function ReadArray(ByVal oStrm As olelib.IStream, ByVal iPropType As Integer) As Variant
Dim iVarType As Integer
Dim lSAPtr As Long
Dim iDims As Integer
Dim lElemLen As Long
Dim Bounds() As SAFEARRAYBOUND
Dim lDataSize As Long
Dim lDataPtr As Long
Dim lIdx As Long

Dim lPtr As Long
Dim sStr As String
Dim oObj As Object
Dim vVar As Variant

   ' Read SAFEARRAY data
   oStrm.Read iDims, 2
   
   ' Check if the array contains data
   If iDims <> 0 Then
   
      oStrm.Seek 0.0002@, STREAM_SEEK_CUR
      oStrm.Read lElemLen, 4
      oStrm.Seek 0.0008@, STREAM_SEEK_CUR
   
      ReDim Bounds(0 To iDims - 1)
   
      ' Get bounds
      For lIdx = 0 To iDims - 1
   
         oStrm.Read Bounds(lIdx), 8
   
         ' Calculate total number of elements
         lDataSize = lDataSize + Bounds(lIdx).cElements
   
      Next
   
      ' Calculate data length
      lDataSize = lDataSize * lElemLen
   
      ' Create the SAFEARRAY
      lSAPtr = SafeArrayCreate(iPropType, iDims, Bounds(0))
   
      ' Get the pointer to the data
      MoveMemory lDataPtr, ByVal lSAPtr + 12, 4
   
      Select Case iPropType
   
         Case vbString
   
            For lIdx = 0 To lDataSize - 1 Step 4
   
               ' Read the string
               sStr = ReadString(oStrm)
   
               ' Create a new BSTR
               lPtr = SysAllocString(StrPtr(sStr))
   
               ' Copy the string pointer to the array
               MoveMemory ByVal lDataPtr + lIdx, lPtr, 4
   
            Next
   
         Case vbVariant
   
            For lIdx = 0 To lDataSize - 1 Step 16
   
               ' Read the variant type
               oStrm.Read iVarType, 2
   
               ' Read the value
               vVar = ReadValue(oStrm, iVarType)
   
               ' Copy the variant to the array
               VariantCopyIndPtrVar lDataPtr + lIdx, vVar
   
            Next
   
         Case vbObject
   
            For lIdx = 0 To lDataSize - 1 Step 4
   
               ' Read the object
               Set oObj = ReadObject(oStrm)
   
               ' Copy the variant to the array
               MoveMemory ByVal lDataPtr + lIdx, oObj, 4
   
               ' Release the object
               MoveMemory oObj, 0&, 4
   
            Next
   
         Case Else
   
            ' Numeric data
            ' write the entire array
            oStrm.Read lDataPtr, lDataSize
   
      End Select
   
   End If ' iDims = 0
   
   ' Copy the SAFEARRAY pointer to
   ' the return value
   MoveMemory ByVal VarPtr(ReadArray) + 8, lSAPtr, 4

   ' Set the variant type
   MoveMemory ReadArray, iPropType Or vbArray, 2

End Function

'
' ReadObject
'
' Reads and returns an object from the stream
'
Function ReadObject(ByVal oStrm As olelib.IStream) As olelib.IUnknown
Dim CLSID As UUID
Dim oPPB As IPersistPropertyBag
Dim oPSI As IPersistStreamInit
Dim oPS As IPersistStream
Dim oRecordset As Object
Dim aBytes() As Byte
Dim lLen As Long
Dim lRes As Long

   ' Read the CLSID
   ReadClassStm oStrm, CLSID

   ' Check if the CLSID
   ' is empty
   If IsEqualGUID(CLSID, IID_Null) Then Exit Function

   ' Create the object
   lRes = CoCreateInstance(CLSID, Nothing, CLSCTX_INPROC_HANDLER Or CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER, IID_IUnknown, ReadObject)
   
   If lRes = S_OK Then

      On Error Resume Next

      ' Query for IPersistStream
      Set oPS = ReadObject

      If Err.Number = 0 Then
      
'         If IsEqualGUID(CLSID_ADODBRecordset, CLSID) Then
'
'            ' Get the IDispatch interface
'            Set oRecordset = ReadObject
'
'            ' Open the recordset
'            oRecordset.Open oStrm
'
'         Else
         
            ' Load the object data
            oPS.Load oStrm
            
'         End If

      Else

         Err.Clear

         ' Query for IPersistStreamInit
         Set oPSI = ReadObject

         If Err.Number = 0 Then

            ' Load the object data
            oPSI.Load oStrm

         Else

'            ' Query for IPersistPropertyBag
'            Set oPPB = ReadObject
'
'            ' Read array length
'            oStrm.Read VarPtr(lLen), 4
'
'            ' Read the array
'            ReDim aBytes(0 To lLen - 1)
'            oStrm.Read VarPtr(aBytes(0)), lLen
'
'            ' Create a temporary property bag
'            Set m_TempPB = New PropBag.PropertyBag
'
'            ' Set the data
'            m_TempPB.Contents = aBytes
'
'            ' Load the object data
'            oPPB.Load Me, Me
'
'            ' Destroy the PropertyBag
'            Set m_TempPB = Nothing

         End If

      End If

   Else

      Err.Raise 429, , ErrorMessage(lRes)

   End If

End Function


'
' ReadValue
'
' Reads a value from the stream
'
Function ReadValue(ByVal oStrm As olelib.IStream, ByVal VarType As Integer) As Variant
Dim oUnknown As stdole.IUnknown
Dim oDispatch As Object

   Select Case VarType

      Case vbObject, VT_UNKNOWN
         
         Set oUnknown = ReadObject(oStrm)
                  
         On Error Resume Next
            
         ' Clear the Err object
         Err.Clear
         
         ' Query for IDispatch interface
         Set oDispatch = oUnknown
            
         If Err.Number = 0 Then
               
            ' Return IDispatch interface
            Set ReadValue = oDispatch
               
         Else
         
            ' Return the IUnknown interface
            Set ReadValue = oUnknown
            
         End If

      Case vbString
         ReadValue = ReadString(oStrm)

      Case vbDecimal
         oStrm.Read VarPtr(ReadValue), 16

      Case Else
         If (VarType And vbArray) = vbArray Then
            ReadValue = ReadArray(oStrm, VarType And Not vbArray)
         Else
            ReadValue = ReadNumber(oStrm, VarType)
         End If

   End Select
   
End Function


'
' ReadString
'
' Reads a string from the stream
'
Function ReadString(ByVal oStrm As olelib.IStream) As String
Dim lLen As Long

   ' Read the string len
   oStrm.Read lLen, 4

   If lLen > 0 Then

      ' Initialize the buffer
      ReadString = Space$(lLen / 2)

      ' Read the string
      oStrm.Read ByVal StrPtr(ReadString), lLen

   End If

End Function


'
' ReadNumber
'
' Reads a numeric value from the stream
'
Function ReadNumber(ByVal oStrm As olelib.IStream, ByVal VarType As Integer) As Variant
Dim lPropLen As Long

   ' Write the variant type
   ' to the return value
   MoveMemory ReadNumber, VarType, 2

   ' Set the data length
   Select Case VarType
      Case vbByte
         lPropLen = 1
      Case vbInteger, vbBoolean
         lPropLen = 2
      Case vbLong, vbSingle
         lPropLen = 4
      Case vbDate, vbDouble, vbCurrency
         lPropLen = 8
   End Select

   ' Read the data into the variant
   oStrm.Read ByVal VarPtr(ReadNumber) + 8, lPropLen

End Function


