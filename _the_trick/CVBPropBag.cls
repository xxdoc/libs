' //
' // CVBPropBag.cls
' // VBPropertyBag content parser
' // By The trick 2020
' // The class is made by reverse engineering so it may contain bugs
' //

Option Explicit

Private Const E_INVALIDARG = &H80070057
Private Const E_UNEXPECTED = &H8000FFFF

Private Type SAFEARRAYBOUND
    cElements   As Long
    lLBound     As Long
End Type

Private Type tPropBagHeader
    iSignature      As Integer
    iCatalogOffset  As Integer
    lContentSize    As Long
End Type

Private Type tCatalogItem
    iVarType        As Integer
    iStrSize        As Integer
    lHash           As Long
    lNextItemOffset As Long
    sName           As String
    ' // bData()         As Byte
End Type

Private Declare Function SafeArrayCreate Lib "oleaut32" ( _
                         ByVal vt As Long, _
                         ByVal cDims As Long, _
                         ByRef rgsaBound As SAFEARRAYBOUND) As Long
Private Declare Function SafeArrayGetElemsize Lib "oleaut32" ( _
                         ByVal pArr As Long) As Long
Private Declare Function SafeArrayAccessData Lib "oleaut32" ( _
                         ByVal pArr As Long, _
                         ByRef pData As Any) As Long
Private Declare Function SafeArrayUnaccessData Lib "oleaut32" ( _
                         ByVal pArr As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function SHCreateMemStream Lib "Shlwapi" _
                         Alias "#12" ( _
                         ByRef pInit As Any, _
                         ByVal cbInit As Long) As IStream
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
                    
Private m_cCatalog  As Collection
Private m_cItems    As Collection

Public Property Get Catalog() As Collection
    Set Catalog = m_cCatalog
End Property

Public Property Get Items() As Collection
    Set Items = m_cItems
End Property

Public Property Let Content( _
                    ByRef bData() As Byte)
    Deserialize bData
End Property

Private Sub Deserialize( _
            ByRef bData() As Byte)
    Dim lContentSize    As Long:            Dim tPropBagHdr     As tPropBagHeader
    Dim lCatIndex       As Long:            Dim tCatItem        As tCatalogItem
    Dim lCatSize        As Long:            Dim lDataOffset     As Long
    Dim vValue          As Variant:         Dim iValue          As Integer
    Dim bValue          As Byte:            Dim lValue          As Long
    Dim fValue          As Single:          Dim dValue          As Single
    Dim curValue        As Currency:        Dim datValue        As Date
    Dim sValue          As String:          Dim tArrBounds()    As SAFEARRAYBOUND
    Dim lArrElemCount   As Long:            Dim lArrTotalSize   As Long
    Dim lDimIndex       As Long:            Dim pArray          As Long
    Dim iDimensions     As Integer:         Dim pArrayData      As Long
    Dim cStm            As IStream:         Dim cCatalog        As Collection
    Dim cItemsData      As Collection
    
    If Not Not bData Then
        lContentSize = UBound(bData) + 1
    Else
        lContentSize = 0
    End If

    If lContentSize < Len(tPropBagHdr) Then
        Err.Raise E_INVALIDARG
    End If
    
    memcpy tPropBagHdr, bData(0), Len(tPropBagHdr)
    
    If tPropBagHdr.iSignature <> &HB293 Or tPropBagHdr.lContentSize > lContentSize Then
        Err.Raise E_INVALIDARG
    End If
    
    Set cItemsData = New Collection
    Set cCatalog = New Collection
    
    If tPropBagHdr.lContentSize > 0 Then
        
        lCatSize = tPropBagHdr.lContentSize - ((tPropBagHdr.iCatalogOffset And &H3FF) + Len(tPropBagHdr))
        If lCatSize <= 0 Then
            Err.Raise E_INVALIDARG
        End If
        
        lCatIndex = tPropBagHdr.lContentSize - lCatSize

        Do
            
            If lCatIndex >= lContentSize - (Len(tPropBagHdr) - 4) Then
                Err.Raise E_INVALIDARG
            End If
            
            memcpy ByVal VarPtr(tCatItem), bData(lCatIndex), Len(tCatItem) - 4
            tCatItem.sName = Space$(tCatItem.iStrSize And &H3FF)
            memcpy ByVal StrPtr(tCatItem.sName), bData(lCatIndex + Len(tCatItem) - 4), tCatItem.iStrSize * 2
            lDataOffset = lCatIndex + (Len(tCatItem) - 4) + tCatItem.iStrSize * 2
            
            If tCatItem.iVarType And vbArray Then
            
                GetMem2 bData(lDataOffset), iDimensions:    lDataOffset = lDataOffset + 2
                
                If iDimensions > 0 Then
                
                    ReDim tArrBounds(iDimensions - 1)
                    memcpy tArrBounds(0), bData(lDataOffset), iDimensions * Len(tArrBounds(0))
                    lArrElemCount = 1
                    lDataOffset = lDataOffset + iDimensions * Len(tArrBounds(0))
                    
                    For lDimIndex = 0 To iDimensions - 1
                        lArrElemCount = lArrElemCount * tArrBounds(lDimIndex).cElements
                    Next
                    
                    pArray = SafeArrayCreate(tCatItem.iVarType And (Not vbArray), iDimensions, tArrBounds(0))
                    lArrTotalSize = SafeArrayGetElemsize(pArray) * lArrElemCount
                    SafeArrayAccessData pArray, pArrayData
                    memcpy ByVal pArrayData, bData(lDataOffset), lArrTotalSize
                    SafeArrayUnaccessData pArray
                
                Else
                    pArray = 0
                End If
                
                vValue = vbNull
                GetMem2 tCatItem.iVarType, vValue
                GetMem4 pArray, ByVal VarPtr(vValue) + 8
                    
            Else
                Select Case tCatItem.iVarType
                Case vbNull
                    vValue = Null
                Case vbEmpty
                    vValue = Empty
                Case vbInteger
                    GetMem2 bData(lDataOffset), iValue
                    vValue = iValue
                Case vbLong
                    GetMem4 bData(lDataOffset), lValue
                    vValue = lValue
                Case vbSingle
                    GetMem4 bData(lDataOffset), fValue
                    vValue = fValue
                Case vbDouble
                    GetMem8 bData(lDataOffset), dValue
                    vValue = dValue
                Case vbCurrency
                    GetMem8 bData(lDataOffset), curValue
                    vValue = curValue
                Case vbDate
                    GetMem8 bData(lDataOffset), datValue
                    vValue = datValue
                Case vbString
                    GetMem4 bData(lDataOffset), lValue
                    If lValue > 0 Then
                        sValue = Space$(lValue)
                        memcpy ByVal StrPtr(sValue), bData(lDataOffset + 4), lValue * 2
                        vValue = sValue
                    Else
                        vValue = vbNullString
                    End If
                Case vbObject, vbDataObject
                    Set cStm = SHCreateMemStream(bData(lDataOffset), lContentSize - lDataOffset)
                    Set vValue = ReadObject(tCatItem.iVarType, cStm)
                Case vbError
                    GetMem4 bData(lDataOffset), lValue
                    vValue = CVErr(lValue)
                Case vbBoolean
                    GetMem2 bData(lDataOffset), iValue
                    vValue = CBool(iValue)
                Case vbVariant
                    Err.Raise E_UNEXPECTED
                Case vbDecimal
                    vValue = Null
                    GetMem8 bData(lDataOffset), vValue
                    GetMem8 bData(lDataOffset + 8), ByVal VarPtr(vValue) + 8
                Case vbByte
                    vValue = bData(lDataOffset)
                End Select
                
            End If

            cCatalog.Add tCatItem.sName
            cItemsData.Add vValue, tCatItem.sName
            
            lCatIndex = lCatIndex + tCatItem.lNextItemOffset
            
        Loop While tCatItem.lNextItemOffset > 0
        
    End If
    
    Set m_cItems = cItemsData
    Set m_cCatalog = cCatalog
    
End Sub

Private Function ReadObject( _
                 ByVal vt As Long, _
                 ByVal cStm As IStream) As Variant
    Dim tClsId      As UUID
    Dim cRet        As IUnknown
    Dim cPersist    As IPersistStream
    Dim cDisp       As Object

    ReadClassStm cStm, tClsId
    
    If IsEqualGUID(tClsId, CLSID_NULL) Then Exit Function
    
    CoCreateInstance tClsId, Nothing, CLSCTX_LOCAL_SERVER Or CLSCTX_INPROC_SERVER, IID_IUnknown, cRet
    
    Set cPersist = cRet
    
    cPersist.Load cStm
    
    If vt = vbObject Then
        Set cDisp = cRet
        Set ReadObject = cDisp
    ElseIf vt = vbDataObject Then
        Set ReadObject = cRet
    End If
    
End Function

Private Function CLSID_NULL() As UUID
    
End Function

Private Function IID_IUnknown() As UUID
    IID_IUnknown.Data4(0) = &HC0
    IID_IUnknown.Data4(7) = &H46
End Function

'Private Function CalcHash( _
'                 ByRef sValue As String) As Long
'    Dim lIndex  As Long
'
'    For lIndex = 0 To Len(sValue) - 1
'        CalcHash = AscW(Mid$(sValue, lIndex + 1, 1)) Xor (ROL(CalcHash, 13))
'    Next
'
'End Function

'Public Function ROL( _
'                 ByVal lValue As Long, _
'                 ByVal bCount As Byte) As Long
'    Static s_lMasks(1, 31)  As Long
'    Dim lResult As Long
'    Dim lIndex  As Long
'    Dim lTemp   As Long
'    Dim lL      As Long
'    Dim lR      As Long
'
'    bCount = bCount And &H1F
'
'    If bCount = 0 Then
'        ROL = lValue
'        Exit Function
'    End If
'
'    If s_lMasks(0, 1) = 0 Then
'
'        lTemp = &H80000000
'
'        For lIndex = 1 To 31
'            s_lMasks(0, lIndex) = lTemp
'            s_lMasks(1, lIndex) = (((Not lTemp) Xor &H80000000) + 1) Xor &H80000000
'            lTemp = lTemp \ 2
'        Next
'
'    End If
'
'    lL = (((lValue And s_lMasks(0, bCount))) \ s_lMasks(1, bCount)) And (Not s_lMasks(0, 32 - bCount))
'    lR = ((lValue And (Not s_lMasks(0, bCount))))
'
'    If (lR And (s_lMasks(1, bCount) \ 2)) Then
'        lR = (lR And Not (s_lMasks(1, bCount) \ 2)) * s_lMasks(1, 32 - bCount) Or &H80000000
'    Else
'        lR = lR * s_lMasks(1, 32 - bCount)
'    End If
'
'    ROL = lR Or lL
'
'End Function

Private Sub Class_Initialize()

    Set m_cItems = New Collection
    Set m_cCatalog = New Collection
    
End Sub