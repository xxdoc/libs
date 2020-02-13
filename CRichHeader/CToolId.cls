VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Compiler As Integer
Public ProductID As Integer
Public Count As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Friend Sub LoadSelf(upack0 As Long, upack1 As Long)

        'better not to use varptr with public class members, it was working but
        'they are usually actually hidden get/let ymmv
        Dim c As Integer, p As Integer
        
        CopyMemory ByVal VarPtr(c), ByVal VarPtr(upack0), 2
        CopyMemory ByVal VarPtr(p), ByVal VarPtr(upack0) + 2, 2
        
        Compiler = c
        ProductID = p
        Count = upack1
        
        'Debug.Print Join(Array(Compiler, Hex(ProductID), Hex(Count), InternalName, VSVersion), " ")

End Sub

Function dump(Optional seperator As String = ",", Optional withName As Boolean = False) As String

    Dim tmp() As String, x As String
    
    push tmp, IIf(withName, "Compiler: ", "") & Compiler
    push tmp, IIf(withName, "ProductID: ", "") & ProductID
    push tmp, IIf(withName, "Count: ", "") & Count
    push tmp, IIf(withName, "InternalName: ", "") & InternalName
    push tmp, IIf(withName, "VSVersion: ", "") & VSVersion
    dump = Join(tmp, seperator)
    
End Function

Private Function in_range(v, min As Long, max As Long) As Boolean
    If v >= min And v <= max Then in_range = True
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Property Get VSVersion() As String

    If ProductID > &H10E Or ProductID < 0 Then
        VSVersion = "<unknown> (XX.XX)"
    ElseIf in_range(ProductID, &HFD, &H10E + 1) Then
        VSVersion = "Visual Studio 2015 (14.00)"
    ElseIf in_range(ProductID, &HEB, &HFD) Then
        VSVersion = "Visual Studio 2013 (12.10)"
    ElseIf in_range(ProductID, &HD9, &HEB) Then
        VSVersion = "Visual Studio 2013 (12.00)"
    ElseIf in_range(ProductID, &HC7, &HD9) Then
        VSVersion = "Visual Studio 2012 (11.00)"
    ElseIf in_range(ProductID, &HB5, &HC7) Then
        VSVersion = "Visual Studio 2010 (10.10)"
    ElseIf in_range(ProductID, &H98, &HB5) Then     'TODO then Investigate on Phoenix
        VSVersion = "Visual Studio 2010 (10.00)"
    ElseIf in_range(ProductID, &H83, &H98) Then     'TODO then Investigate on Phoenix
        VSVersion = "Visual Studio 2008 (09.00)"
    ElseIf in_range(ProductID, &H6D, &H83) Then
        VSVersion = "Visual Studio 2005 (08.00)"
    ElseIf in_range(ProductID, &H5A, &H6D) Then
        VSVersion = "Visual Studio 2003 (07.10)"
    ElseIf ProductID = 1 Then
        VSVersion = "Visual Studio (00.00)"
    Else
        VSVersion = "<unknown> (00.00)"
    End If

End Property

Property Get InternalName() As String
    
    Select Case ProductID
        Case &H0: InternalName = "prodidUnknown"
        Case &H1: InternalName = "prodidImport0"
        Case &H2: InternalName = "prodidLinker510"
        Case &H3: InternalName = "prodidCvtomf510"
        Case &H4: InternalName = "prodidLinker600"
        Case &H5: InternalName = "prodidCvtomf600"
        Case &H6: InternalName = "prodidCvtres500"
        Case &H7: InternalName = "prodidUtc11_Basic"
        Case &H8: InternalName = "prodidUtc11_C"
        Case &H9: InternalName = "prodidUtc12_Basic"
        Case &HA: InternalName = "prodidUtc12_C"
        Case &HB: InternalName = "prodidUtc12_CPP"
        Case &HC: InternalName = "prodidAliasObj60"
        Case &HD: InternalName = "prodidVisualBasic60"
        Case &HE: InternalName = "prodidMasm613"
        Case &HF: InternalName = "prodidMasm710"
        Case &H10: InternalName = "prodidLinker511"
        Case &H11: InternalName = "prodidCvtomf511"
        Case &H12: InternalName = "prodidMasm614"
        Case &H13: InternalName = "prodidLinker512"
        Case &H14: InternalName = "prodidCvtomf512"
        Case &H15: InternalName = "prodidUtc12_C_Std"
        Case &H16: InternalName = "prodidUtc12_CPP_Std"
        Case &H17: InternalName = "prodidUtc12_C_Book"
        Case &H18: InternalName = "prodidUtc12_CPP_Book"
        Case &H19: InternalName = "prodidImplib700"
        Case &H1A: InternalName = "prodidCvtomf700"
        Case &H1B: InternalName = "prodidUtc13_Basic"
        Case &H1C: InternalName = "prodidUtc13_C"
        Case &H1D: InternalName = "prodidUtc13_CPP"
        Case &H1E: InternalName = "prodidLinker610"
        Case &H1F: InternalName = "prodidCvtomf610"
        Case &H20: InternalName = "prodidLinker601"
        Case &H21: InternalName = "prodidCvtomf601"
        Case &H22: InternalName = "prodidUtc12_1_Basic"
        Case &H23: InternalName = "prodidUtc12_1_C"
        Case &H24: InternalName = "prodidUtc12_1_CPP"
        Case &H25: InternalName = "prodidLinker620"
        Case &H26: InternalName = "prodidCvtomf620"
        Case &H27: InternalName = "prodidAliasObj70"
        Case &H28: InternalName = "prodidLinker621"
        Case &H29: InternalName = "prodidCvtomf621"
        Case &H2A: InternalName = "prodidMasm615"
        Case &H2B: InternalName = "prodidUtc13_LTCG_C"
        Case &H2C: InternalName = "prodidUtc13_LTCG_CPP"
        Case &H2D: InternalName = "prodidMasm620"
        Case &H2E: InternalName = "prodidILAsm100"
        Case &H2F: InternalName = "prodidUtc12_2_Basic"
        Case &H30: InternalName = "prodidUtc12_2_C"
        Case &H31: InternalName = "prodidUtc12_2_CPP"
        Case &H32: InternalName = "prodidUtc12_2_C_Std"
        Case &H33: InternalName = "prodidUtc12_2_CPP_Std"
        Case &H34: InternalName = "prodidUtc12_2_C_Book"
        Case &H35: InternalName = "prodidUtc12_2_CPP_Book"
        Case &H36: InternalName = "prodidImplib622"
        Case &H37: InternalName = "prodidCvtomf622"
        Case &H38: InternalName = "prodidCvtres501"
        Case &H39: InternalName = "prodidUtc13_C_Std"
        Case &H3A: InternalName = "prodidUtc13_CPP_Std"
        Case &H3B: InternalName = "prodidCvtpgd1300"
        Case &H3C: InternalName = "prodidLinker622"
        Case &H3D: InternalName = "prodidLinker700"
        Case &H3E: InternalName = "prodidExport622"
        Case &H3F: InternalName = "prodidExport700"
        Case &H40: InternalName = "prodidMasm700"
        Case &H41: InternalName = "prodidUtc13_POGO_I_C"
        Case &H42: InternalName = "prodidUtc13_POGO_I_CPP"
        Case &H43: InternalName = "prodidUtc13_POGO_O_C"
        Case &H44: InternalName = "prodidUtc13_POGO_O_CPP"
        Case &H45: InternalName = "prodidCvtres700"
        Case &H46: InternalName = "prodidCvtres710p"
        Case &H47: InternalName = "prodidLinker710p"
        Case &H48: InternalName = "prodidCvtomf710p"
        Case &H49: InternalName = "prodidExport710p"
        Case &H4A: InternalName = "prodidImplib710p"
        Case &H4B: InternalName = "prodidMasm710p"
        Case &H4C: InternalName = "prodidUtc1310p_C"
        Case &H4D: InternalName = "prodidUtc1310p_CPP"
        Case &H4E: InternalName = "prodidUtc1310p_C_Std"
        Case &H4F: InternalName = "prodidUtc1310p_CPP_Std"
        Case &H50: InternalName = "prodidUtc1310p_LTCG_C"
        Case &H51: InternalName = "prodidUtc1310p_LTCG_CPP"
        Case &H52: InternalName = "prodidUtc1310p_POGO_I_C"
        Case &H53: InternalName = "prodidUtc1310p_POGO_I_CPP"
        Case &H54: InternalName = "prodidUtc1310p_POGO_O_C"
        Case &H55: InternalName = "prodidUtc1310p_POGO_O_CPP"
        Case &H56: InternalName = "prodidLinker624"
        Case &H57: InternalName = "prodidCvtomf624"
        Case &H58: InternalName = "prodidExport624"
        Case &H59: InternalName = "prodidImplib624"
        Case &H5A: InternalName = "prodidLinker710"
        Case &H5B: InternalName = "prodidCvtomf710"
        Case &H5C: InternalName = "prodidExport710"
        Case &H5D: InternalName = "prodidImplib710"
        Case &H5E: InternalName = "prodidCvtres710"
        Case &H5F: InternalName = "prodidUtc1310_C"
        Case &H60: InternalName = "prodidUtc1310_CPP"
        Case &H61: InternalName = "prodidUtc1310_C_Std"
        Case &H62: InternalName = "prodidUtc1310_CPP_Std"
        Case &H63: InternalName = "prodidUtc1310_LTCG_C"
        Case &H64: InternalName = "prodidUtc1310_LTCG_CPP"
        Case &H65: InternalName = "prodidUtc1310_POGO_I_C"
        Case &H66: InternalName = "prodidUtc1310_POGO_I_CPP"
        Case &H67: InternalName = "prodidUtc1310_POGO_O_C"
        Case &H68: InternalName = "prodidUtc1310_POGO_O_CPP"
        Case &H69: InternalName = "prodidAliasObj710"
        Case &H6A: InternalName = "prodidAliasObj710p"
        Case &H6B: InternalName = "prodidCvtpgd1310"
        Case &H6C: InternalName = "prodidCvtpgd1310p"
        Case &H6D: InternalName = "prodidUtc1400_C"
        Case &H6E: InternalName = "prodidUtc1400_CPP"
        Case &H6F: InternalName = "prodidUtc1400_C_Std"
        Case &H70: InternalName = "prodidUtc1400_CPP_Std"
        Case &H71: InternalName = "prodidUtc1400_LTCG_C"
        Case &H72: InternalName = "prodidUtc1400_LTCG_CPP"
        Case &H73: InternalName = "prodidUtc1400_POGO_I_C"
        Case &H74: InternalName = "prodidUtc1400_POGO_I_CPP"
        Case &H75: InternalName = "prodidUtc1400_POGO_O_C"
        Case &H76: InternalName = "prodidUtc1400_POGO_O_CPP"
        Case &H77: InternalName = "prodidCvtpgd1400"
        Case &H78: InternalName = "prodidLinker800"
        Case &H79: InternalName = "prodidCvtomf800"
        Case &H7A: InternalName = "prodidExport800"
        Case &H7B: InternalName = "prodidImplib800"
        Case &H7C: InternalName = "prodidCvtres800"
        Case &H7D: InternalName = "prodidMasm800"
        Case &H7E: InternalName = "prodidAliasObj800"
        Case &H7F: InternalName = "prodidPhoenixPrerelease"
        Case &H80: InternalName = "prodidUtc1400_CVTCIL_C"
        Case &H81: InternalName = "prodidUtc1400_CVTCIL_CPP"
        Case &H82: InternalName = "prodidUtc1400_LTCG_MSIL"
        Case &H83: InternalName = "prodidUtc1500_C"
        Case &H84: InternalName = "prodidUtc1500_CPP"
        Case &H85: InternalName = "prodidUtc1500_C_Std"
        Case &H86: InternalName = "prodidUtc1500_CPP_Std"
        Case &H87: InternalName = "prodidUtc1500_CVTCIL_C"
        Case &H88: InternalName = "prodidUtc1500_CVTCIL_CPP"
        Case &H89: InternalName = "prodidUtc1500_LTCG_C"
        Case &H8A: InternalName = "prodidUtc1500_LTCG_CPP"
        Case &H8B: InternalName = "prodidUtc1500_LTCG_MSIL"
        Case &H8C: InternalName = "prodidUtc1500_POGO_I_C"
        Case &H8D: InternalName = "prodidUtc1500_POGO_I_CPP"
        Case &H8E: InternalName = "prodidUtc1500_POGO_O_C"
        Case &H8F: InternalName = "prodidUtc1500_POGO_O_CPP"
        Case &H90: InternalName = "prodidCvtpgd1500"
        Case &H91: InternalName = "prodidLinker900"
        Case &H92: InternalName = "prodidExport900"
        Case &H93: InternalName = "prodidImplib900"
        Case &H94: InternalName = "prodidCvtres900"
        Case &H95: InternalName = "prodidMasm900"
        Case &H96: InternalName = "prodidAliasObj900"
        Case &H97: InternalName = "prodidResource"
        Case &H98: InternalName = "prodidAliasObj1000"
        Case &H99: InternalName = "prodidCvtpgd1600"
        Case &H9A: InternalName = "prodidCvtres1000"
        Case &H9B: InternalName = "prodidExport1000"
        Case &H9C: InternalName = "prodidImplib1000"
        Case &H9D: InternalName = "prodidLinker1000"
        Case &H9E: InternalName = "prodidMasm1000"
        Case &H9F: InternalName = "prodidPhx1600_C"
        Case &HA0: InternalName = "prodidPhx1600_CPP"
        Case &HA1: InternalName = "prodidPhx1600_CVTCIL_C"
        Case &HA2: InternalName = "prodidPhx1600_CVTCIL_CPP"
        Case &HA3: InternalName = "prodidPhx1600_LTCG_C"
        Case &HA4: InternalName = "prodidPhx1600_LTCG_CPP"
        Case &HA5: InternalName = "prodidPhx1600_LTCG_MSIL"
        Case &HA6: InternalName = "prodidPhx1600_POGO_I_C"
        Case &HA7: InternalName = "prodidPhx1600_POGO_I_CPP"
        Case &HA8: InternalName = "prodidPhx1600_POGO_O_C"
        Case &HA9: InternalName = "prodidPhx1600_POGO_O_CPP"
        Case &HAA: InternalName = "prodidUtc1600_C"
        Case &HAB: InternalName = "prodidUtc1600_CPP"
        Case &HAC: InternalName = "prodidUtc1600_CVTCIL_C"
        Case &HAD: InternalName = "prodidUtc1600_CVTCIL_CPP"
        Case &HAE: InternalName = "prodidUtc1600_LTCG_C"
        Case &HAF: InternalName = "prodidUtc1600_LTCG_CPP"
        Case &HB0: InternalName = "prodidUtc1600_LTCG_MSIL"
        Case &HB1: InternalName = "prodidUtc1600_POGO_I_C"
        Case &HB2: InternalName = "prodidUtc1600_POGO_I_CPP"
        Case &HB3: InternalName = "prodidUtc1600_POGO_O_C"
        Case &HB4: InternalName = "prodidUtc1600_POGO_O_CPP"
        Case &HB5: InternalName = "prodidAliasObj1010"
        Case &HB6: InternalName = "prodidCvtpgd1610"
        Case &HB7: InternalName = "prodidCvtres1010"
        Case &HB8: InternalName = "prodidExport1010"
        Case &HB9: InternalName = "prodidImplib1010"
        Case &HBA: InternalName = "prodidLinker1010"
        Case &HBB: InternalName = "prodidMasm1010"
        Case &HBC: InternalName = "prodidUtc1610_C"
        Case &HBD: InternalName = "prodidUtc1610_CPP"
        Case &HBE: InternalName = "prodidUtc1610_CVTCIL_C"
        Case &HBF: InternalName = "prodidUtc1610_CVTCIL_CPP"
        Case &HC0: InternalName = "prodidUtc1610_LTCG_C"
        Case &HC1: InternalName = "prodidUtc1610_LTCG_CPP"
        Case &HC2: InternalName = "prodidUtc1610_LTCG_MSIL"
        Case &HC3: InternalName = "prodidUtc1610_POGO_I_C"
        Case &HC4: InternalName = "prodidUtc1610_POGO_I_CPP"
        Case &HC5: InternalName = "prodidUtc1610_POGO_O_C"
        Case &HC6: InternalName = "prodidUtc1610_POGO_O_CPP"
        Case &HC7: InternalName = "prodidAliasObj1100"
        Case &HC8: InternalName = "prodidCvtpgd1700"
        Case &HC9: InternalName = "prodidCvtres1100"
        Case &HCA: InternalName = "prodidExport1100"
        Case &HCB: InternalName = "prodidImplib1100"
        Case &HCC: InternalName = "prodidLinker1100"
        Case &HCD: InternalName = "prodidMasm1100"
        Case &HCE: InternalName = "prodidUtc1700_C"
        Case &HCF: InternalName = "prodidUtc1700_CPP"
        Case &HD0: InternalName = "prodidUtc1700_CVTCIL_C"
        Case &HD1: InternalName = "prodidUtc1700_CVTCIL_CPP"
        Case &HD2: InternalName = "prodidUtc1700_LTCG_C"
        Case &HD3: InternalName = "prodidUtc1700_LTCG_CPP"
        Case &HD4: InternalName = "prodidUtc1700_LTCG_MSIL"
        Case &HD5: InternalName = "prodidUtc1700_POGO_I_C"
        Case &HD6: InternalName = "prodidUtc1700_POGO_I_CPP"
        Case &HD7: InternalName = "prodidUtc1700_POGO_O_C"
        Case &HD8: InternalName = "prodidUtc1700_POGO_O_CPP"
        Case &HD9: InternalName = "prodidAliasObj1200"
        Case &HDA: InternalName = "prodidCvtpgd1800"
        Case &HDB: InternalName = "prodidCvtres1200"
        Case &HDC: InternalName = "prodidExport1200"
        Case &HDD: InternalName = "prodidImplib1200"
        Case &HDE: InternalName = "prodidLinker1200"
        Case &HDF: InternalName = "prodidMasm1200"
        Case &HE0: InternalName = "prodidUtc1800_C"
        Case &HE1: InternalName = "prodidUtc1800_CPP"
        Case &HE2: InternalName = "prodidUtc1800_CVTCIL_C"
        Case &HD3: InternalName = "prodidUtc1800_CVTCIL_CPP"
        Case &HE4: InternalName = "prodidUtc1800_LTCG_C"
        Case &HE5: InternalName = "prodidUtc1800_LTCG_CPP"
        Case &HE6: InternalName = "prodidUtc1800_LTCG_MSIL"
        Case &HE7: InternalName = "prodidUtc1800_POGO_I_C"
        Case &HE8: InternalName = "prodidUtc1800_POGO_I_CPP"
        Case &HE9: InternalName = "prodidUtc1800_POGO_O_C"
        Case &HEA: InternalName = "prodidUtc1800_POGO_O_CPP"
        Case &HEB: InternalName = "prodidAliasObj1210"
        Case &HEC: InternalName = "prodidCvtpgd1810"
        Case &HED: InternalName = "prodidCvtres1210"
        Case &HEE: InternalName = "prodidExport1210"
        Case &HEF: InternalName = "prodidImplib1210"
        Case &HF0: InternalName = "prodidLinker1210"
        Case &HF1: InternalName = "prodidMasm1210"
        Case &HF2: InternalName = "prodidUtc1810_C"
        Case &HF3: InternalName = "prodidUtc1810_CPP"
        Case &HF4: InternalName = "prodidUtc1810_CVTCIL_C"
        Case &HF5: InternalName = "prodidUtc1810_CVTCIL_CPP"
        Case &HF6: InternalName = "prodidUtc1810_LTCG_C"
        Case &HF7: InternalName = "prodidUtc1810_LTCG_CPP"
        Case &HF8: InternalName = "prodidUtc1810_LTCG_MSIL"
        Case &HF9: InternalName = "prodidUtc1810_POGO_I_C"
        Case &HFA: InternalName = "prodidUtc1810_POGO_I_CPP"
        Case &HFB: InternalName = "prodidUtc1810_POGO_O_C"
        Case &HFC: InternalName = "prodidUtc1810_POGO_O_CPP"
        Case &HFD: InternalName = "prodidAliasObj1400"
        Case &HFE: InternalName = "prodidCvtpgd1900"
        Case &HFF: InternalName = "prodidCvtres1400"
        Case &H100: InternalName = "prodidExport1400"
        Case &H101: InternalName = "prodidImplib1400"
        Case &H102: InternalName = "prodidLinker1400"
        Case &H103: InternalName = "prodidMasm1400"
        Case &H104: InternalName = "prodidUtc1900_C"
        Case &H105: InternalName = "prodidUtc1900_CPP"
        Case &H106: InternalName = "prodidUtc1900_CVTCIL_C"
        Case &H107: InternalName = "prodidUtc1900_CVTCIL_CPP"
        Case &H108: InternalName = "prodidUtc1900_LTCG_C"
        Case &H109: InternalName = "prodidUtc1900_LTCG_CPP"
        Case &H10A: InternalName = "prodidUtc1900_LTCG_MSIL"
        Case &H10B: InternalName = "prodidUtc1900_POGO_I_C"
        Case &H10C: InternalName = "prodidUtc1900_POGO_I_CPP"
        Case &H10D: InternalName = "prodidUtc1900_POGO_O_C"
        Case &H10E: InternalName = "prodidUtc1900_POGO_O_CPP"
    End Select

End Property
