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

Public compiler As Integer
Public toolID As rhToolID
Public Count As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Enum rhToolID
    rhUnknown = &H0
    rhImport0 = &H1
    rhLinker510 = &H2
    rhCvtomf510 = &H3
    rhLinker600 = &H4
    rhCvtomf600 = &H5
    rhCvtres500 = &H6
    rhUtc11_Basic = &H7
    rhUtc11_C = &H8
    rhUtc12_Basic = &H9
    rhUtc12_C = &HA
    rhUtc12_CPP = &HB
    rhAliasObj60 = &HC
    rhVisualBasic60 = &HD
    rhMasm613 = &HE
    rhMasm710 = &HF
    rhLinker511 = &H10
    rhCvtomf511 = &H11
    rhMasm614 = &H12
    rhLinker512 = &H13
    rhCvtomf512 = &H14
    rhUtc12_C_Std = &H15
    rhUtc12_CPP_Std = &H16
    rhUtc12_C_Book = &H17
    rhUtc12_CPP_Book = &H18
    rhImplib700 = &H19
    rhCvtomf700 = &H1A
    rhUtc13_Basic = &H1B
    rhUtc13_C = &H1C
    rhUtc13_CPP = &H1D
    rhLinker610 = &H1E
    rhCvtomf610 = &H1F
    rhLinker601 = &H20
    rhCvtomf601 = &H21
    rhUtc12_1_Basic = &H22
    rhUtc12_1_C = &H23
    rhUtc12_1_CPP = &H24
    rhLinker620 = &H25
    rhCvtomf620 = &H26
    rhAliasObj70 = &H27
    rhLinker621 = &H28
    rhCvtomf621 = &H29
    rhMasm615 = &H2A
    rhUtc13_LTCG_C = &H2B
    rhUtc13_LTCG_CPP = &H2C
    rhMasm620 = &H2D
    rhILAsm100 = &H2E
    rhUtc12_2_Basic = &H2F
    rhUtc12_2_C = &H30
    rhUtc12_2_CPP = &H31
    rhUtc12_2_C_Std = &H32
    rhUtc12_2_CPP_Std = &H33
    rhUtc12_2_C_Book = &H34
    rhUtc12_2_CPP_Book = &H35
    rhImplib622 = &H36
    rhCvtomf622 = &H37
    rhCvtres501 = &H38
    rhUtc13_C_Std = &H39
    rhUtc13_CPP_Std = &H3A
    rhCvtpgd1300 = &H3B
    rhLinker622 = &H3C
    rhLinker700 = &H3D
    rhExport622 = &H3E
    rhExport700 = &H3F
    rhMasm700 = &H40
    rhUtc13_POGO_I_C = &H41
    rhUtc13_POGO_I_CPP = &H42
    rhUtc13_POGO_O_C = &H43
    rhUtc13_POGO_O_CPP = &H44
    rhCvtres700 = &H45
    rhCvtres710p = &H46
    rhLinker710p = &H47
    rhCvtomf710p = &H48
    rhExport710p = &H49
    rhImplib710p = &H4A
    rhMasm710p = &H4B
    rhUtc1310p_C = &H4C
    rhUtc1310p_CPP = &H4D
    rhUtc1310p_C_Std = &H4E
    rhUtc1310p_CPP_Std = &H4F
    rhUtc1310p_LTCG_C = &H50
    rhUtc1310p_LTCG_CPP = &H51
    rhUtc1310p_POGO_I_C = &H52
    rhUtc1310p_POGO_I_CPP = &H53
    rhUtc1310p_POGO_O_C = &H54
    rhUtc1310p_POGO_O_CPP = &H55
    rhLinker624 = &H56
    rhCvtomf624 = &H57
    rhExport624 = &H58
    rhImplib624 = &H59
    rhLinker710 = &H5A
    rhCvtomf710 = &H5B
    rhExport710 = &H5C
    rhImplib710 = &H5D
    rhCvtres710 = &H5E
    rhUtc1310_C = &H5F
    rhUtc1310_CPP = &H60
    rhUtc1310_C_Std = &H61
    rhUtc1310_CPP_Std = &H62
    rhUtc1310_LTCG_C = &H63
    rhUtc1310_LTCG_CPP = &H64
    rhUtc1310_POGO_I_C = &H65
    rhUtc1310_POGO_I_CPP = &H66
    rhUtc1310_POGO_O_C = &H67
    rhUtc1310_POGO_O_CPP = &H68
    rhAliasObj710 = &H69
    rhAliasObj710p = &H6A
    rhCvtpgd1310 = &H6B
    rhCvtpgd1310p = &H6C
    rhUtc1400_C = &H6D
    rhUtc1400_CPP = &H6E
    rhUtc1400_C_Std = &H6F
    rhUtc1400_CPP_Std = &H70
    rhUtc1400_LTCG_C = &H71
    rhUtc1400_LTCG_CPP = &H72
    rhUtc1400_POGO_I_C = &H73
    rhUtc1400_POGO_I_CPP = &H74
    rhUtc1400_POGO_O_C = &H75
    rhUtc1400_POGO_O_CPP = &H76
    rhCvtpgd1400 = &H77
    rhLinker800 = &H78
    rhCvtomf800 = &H79
    rhExport800 = &H7A
    rhImplib800 = &H7B
    rhCvtres800 = &H7C
    rhMasm800 = &H7D
    rhAliasObj800 = &H7E
    rhPhoenixPrerelease = &H7F
    rhUtc1400_CVTCIL_C = &H80
    rhUtc1400_CVTCIL_CPP = &H81
    rhUtc1400_LTCG_MSIL = &H82
    rhUtc1500_C = &H83
    rhUtc1500_CPP = &H84
    rhUtc1500_C_Std = &H85
    rhUtc1500_CPP_Std = &H86
    rhUtc1500_CVTCIL_C = &H87
    rhUtc1500_CVTCIL_CPP = &H88
    rhUtc1500_LTCG_C = &H89
    rhUtc1500_LTCG_CPP = &H8A
    rhUtc1500_LTCG_MSIL = &H8B
    rhUtc1500_POGO_I_C = &H8C
    rhUtc1500_POGO_I_CPP = &H8D
    rhUtc1500_POGO_O_C = &H8E
    rhUtc1500_POGO_O_CPP = &H8F
    rhCvtpgd1500 = &H90
    rhLinker900 = &H91
    rhExport900 = &H92
    rhImplib900 = &H93
    rhCvtres900 = &H94
    rhMasm900 = &H95
    rhAliasObj900 = &H96
    rhResource = &H97
    rhAliasObj1000 = &H98
    rhCvtpgd1600 = &H99
    rhCvtres1000 = &H9A
    rhExport1000 = &H9B
    rhImplib1000 = &H9C
    rhLinker1000 = &H9D
    rhMasm1000 = &H9E
    rhPhx1600_C = &H9F
    rhPhx1600_CPP = &HA0
    rhPhx1600_CVTCIL_C = &HA1
    rhPhx1600_CVTCIL_CPP = &HA2
    rhPhx1600_LTCG_C = &HA3
    rhPhx1600_LTCG_CPP = &HA4
    rhPhx1600_LTCG_MSIL = &HA5
    rhPhx1600_POGO_I_C = &HA6
    rhPhx1600_POGO_I_CPP = &HA7
    rhPhx1600_POGO_O_C = &HA8
    rhPhx1600_POGO_O_CPP = &HA9
    rhUtc1600_C = &HAA
    rhUtc1600_CPP = &HAB
    rhUtc1600_CVTCIL_C = &HAC
    rhUtc1600_CVTCIL_CPP = &HAD
    rhUtc1600_LTCG_C = &HAE
    rhUtc1600_LTCG_CPP = &HAF
    rhUtc1600_LTCG_MSIL = &HB0
    rhUtc1600_POGO_I_C = &HB1
    rhUtc1600_POGO_I_CPP = &HB2
    rhUtc1600_POGO_O_C = &HB3
    rhUtc1600_POGO_O_CPP = &HB4
    rhAliasObj1010 = &HB5
    rhCvtpgd1610 = &HB6
    rhCvtres1010 = &HB7
    rhExport1010 = &HB8
    rhImplib1010 = &HB9
    rhLinker1010 = &HBA
    rhMasm1010 = &HBB
    rhUtc1610_C = &HBC
    rhUtc1610_CPP = &HBD
    rhUtc1610_CVTCIL_C = &HBE
    rhUtc1610_CVTCIL_CPP = &HBF
    rhUtc1610_LTCG_C = &HC0
    rhUtc1610_LTCG_CPP = &HC1
    rhUtc1610_LTCG_MSIL = &HC2
    rhUtc1610_POGO_I_C = &HC3
    rhUtc1610_POGO_I_CPP = &HC4
    rhUtc1610_POGO_O_C = &HC5
    rhUtc1610_POGO_O_CPP = &HC6
    rhAliasObj1100 = &HC7
    rhCvtpgd1700 = &HC8
    rhCvtres1100 = &HC9
    rhExport1100 = &HCA
    rhImplib1100 = &HCB
    rhLinker1100 = &HCC
    rhMasm1100 = &HCD
    rhUtc1700_C = &HCE
    rhUtc1700_CPP = &HCF
    rhUtc1700_CVTCIL_C = &HD0
    rhUtc1700_CVTCIL_CPP = &HD1
    rhUtc1700_LTCG_C = &HD2
    rhUtc1700_LTCG_CPP = &HD3
    rhUtc1700_LTCG_MSIL = &HD4
    rhUtc1700_POGO_I_C = &HD5
    rhUtc1700_POGO_I_CPP = &HD6
    rhUtc1700_POGO_O_C = &HD7
    rhUtc1700_POGO_O_CPP = &HD8
    rhAliasObj1200 = &HD9
    rhCvtpgd1800 = &HDA
    rhCvtres1200 = &HDB
    rhExport1200 = &HDC
    rhImplib1200 = &HDD
    rhLinker1200 = &HDE
    rhMasm1200 = &HDF
    rhUtc1800_C = &HE0
    rhUtc1800_CPP = &HE1
    rhUtc1800_CVTCIL_C = &HE2
    rhUtc1800_CVTCIL_CPP = &HD3
    rhUtc1800_LTCG_C = &HE4
    rhUtc1800_LTCG_CPP = &HE5
    rhUtc1800_LTCG_MSIL = &HE6
    rhUtc1800_POGO_I_C = &HE7
    rhUtc1800_POGO_I_CPP = &HE8
    rhUtc1800_POGO_O_C = &HE9
    rhUtc1800_POGO_O_CPP = &HEA
    rhAliasObj1210 = &HEB
    rhCvtpgd1810 = &HEC
    rhCvtres1210 = &HED
    rhExport1210 = &HEE
    rhImplib1210 = &HEF
    rhLinker1210 = &HF0
    rhMasm1210 = &HF1
    rhUtc1810_C = &HF2
    rhUtc1810_CPP = &HF3
    rhUtc1810_CVTCIL_C = &HF4
    rhUtc1810_CVTCIL_CPP = &HF5
    rhUtc1810_LTCG_C = &HF6
    rhUtc1810_LTCG_CPP = &HF7
    rhUtc1810_LTCG_MSIL = &HF8
    rhUtc1810_POGO_I_C = &HF9
    rhUtc1810_POGO_I_CPP = &HFA
    rhUtc1810_POGO_O_C = &HFB
    rhUtc1810_POGO_O_CPP = &HFC
    rhAliasObj1400 = &HFD
    rhCvtpgd1900 = &HFE
    rhCvtres1400 = &HFF
    rhExport1400 = &H100
    rhImplib1400 = &H101
    rhLinker1400 = &H102
    rhMasm1400 = &H103
    rhUtc1900_C = &H104
    rhUtc1900_CPP = &H105
    rhUtc1900_CVTCIL_C = &H106
    rhUtc1900_CVTCIL_CPP = &H107
    rhUtc1900_LTCG_C = &H108
    rhUtc1900_LTCG_CPP = &H109
    rhUtc1900_LTCG_MSIL = &H10A
    rhUtc1900_POGO_I_C = &H10B
    rhUtc1900_POGO_I_CPP = &H10C
    rhUtc1900_POGO_O_C = &H10D
    rhUtc1900_POGO_O_CPP = &H10E
End Enum

Friend Sub LoadSelf(upack0 As Long, upack1 As Long)

        'better not to use varptr with public class members, it was working but
        'they are usually actually hidden get/let ymmv
        Dim c As Integer, p As Integer
        
        CopyMemory ByVal VarPtr(c), ByVal VarPtr(upack0), 2
        CopyMemory ByVal VarPtr(p), ByVal VarPtr(upack0) + 2, 2
        
        compiler = c
        toolID = p
        Count = upack1
        
        'Debug.Print Join(Array(Compiler, Hex(ProductID), Hex(Count), InternalName, VSVersion), " ")

End Sub

Function dump(Optional seperator As String = ",", Optional withName As Boolean = False) As String

    Dim tmp() As String, x As String
    
    push tmp, IIf(withName, "Compiler: ", "") & compiler
    push tmp, IIf(withName, "toolID: ", "") & toolID
    push tmp, IIf(withName, "Count: ", "") & Count
    push tmp, IIf(withName, "toolName: ", "") & toolName
    push tmp, IIf(withName, "vsVer: ", "") & vsVer
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

Property Get vsVer() As String
    
    Dim x As String
    x = "<unknown> (00.00)"
    
    If in_range(toolID, &HFD, &H10E + 1) Then
        x = "Visual Studio 2015 (14.00)"
    ElseIf in_range(toolID, &HEB, &HFD) Then
        x = "Visual Studio 2013 (12.10)"
    ElseIf in_range(toolID, &HD9, &HEB) Then
        x = "Visual Studio 2013 (12.00)"
    ElseIf in_range(toolID, &HC7, &HD9) Then
        x = "Visual Studio 2012 (11.00)"
    ElseIf in_range(toolID, &HB5, &HC7) Then
        x = "Visual Studio 2010 (10.10)"
    ElseIf in_range(toolID, &H98, &HB5) Then     'TODO then Investigate on Phoenix
        x = "Visual Studio 2010 (10.00)"
    ElseIf in_range(toolID, &H83, &H98) Then     'TODO then Investigate on Phoenix
        x = "Visual Studio 2008 (09.00)"
    ElseIf in_range(toolID, &H6D, &H83) Then
        x = "Visual Studio 2005 (08.00)"
    ElseIf in_range(toolID, &H5A, &H6D) Then
        x = "Visual Studio 2003 (07.10)"
    ElseIf in_range(toolID, &H4, 14) Then
        x = "Visual Studio 6 (06.00)"           'compile some stuff with vc6 and get more, find vc4
    ElseIf toolID = 1 Then
        x = "Visual Studio (00.00)"
    End If

    vsVer = x
    
End Property

Property Get toolName() As String
    
    Dim x As String
    'x = "unk_0x" & Hex(ProductID)
    
    Select Case toolID
        'Case &H0: x = "Unknown"
        Case &H1: x = "Import0"
        Case &H2: x = "Linker510"
        Case &H3: x = "Cvtomf510"
        Case &H4: x = "Linker600"
        Case &H5: x = "Cvtomf600"
        Case &H6: x = "Cvtres500"
        Case &H7: x = "Utc11_Basic"
        Case &H8: x = "Utc11_C"
        Case &H9: x = "Utc12_Basic"
        Case &HA: x = "Utc12_C"
        Case &HB: x = "Utc12_CPP"
        Case &HC: x = "AliasObj60"
        Case &HD: x = "VisualBasic60"
        Case &HE: x = "Masm613"
        Case &HF: x = "Masm710"
        Case &H10: x = "Linker511"
        Case &H11: x = "Cvtomf511"
        Case &H12: x = "Masm614"
        Case &H13: x = "Linker512"
        Case &H14: x = "Cvtomf512"
        Case &H15: x = "Utc12_C_Std"
        Case &H16: x = "Utc12_CPP_Std"
        Case &H17: x = "Utc12_C_Book"
        Case &H18: x = "Utc12_CPP_Book"
        Case &H19: x = "Implib700"
        Case &H1A: x = "Cvtomf700"
        Case &H1B: x = "Utc13_Basic"
        Case &H1C: x = "Utc13_C"
        Case &H1D: x = "Utc13_CPP"
        Case &H1E: x = "Linker610"
        Case &H1F: x = "Cvtomf610"
        Case &H20: x = "Linker601"
        Case &H21: x = "Cvtomf601"
        Case &H22: x = "Utc12_1_Basic"
        Case &H23: x = "Utc12_1_C"
        Case &H24: x = "Utc12_1_CPP"
        Case &H25: x = "Linker620"
        Case &H26: x = "Cvtomf620"
        Case &H27: x = "AliasObj70"
        Case &H28: x = "Linker621"
        Case &H29: x = "Cvtomf621"
        Case &H2A: x = "Masm615"
        Case &H2B: x = "Utc13_LTCG_C"
        Case &H2C: x = "Utc13_LTCG_CPP"
        Case &H2D: x = "Masm620"
        Case &H2E: x = "ILAsm100"
        Case &H2F: x = "Utc12_2_Basic"
        Case &H30: x = "Utc12_2_C"
        Case &H31: x = "Utc12_2_CPP"
        Case &H32: x = "Utc12_2_C_Std"
        Case &H33: x = "Utc12_2_CPP_Std"
        Case &H34: x = "Utc12_2_C_Book"
        Case &H35: x = "Utc12_2_CPP_Book"
        Case &H36: x = "Implib622"
        Case &H37: x = "Cvtomf622"
        Case &H38: x = "Cvtres501"
        Case &H39: x = "Utc13_C_Std"
        Case &H3A: x = "Utc13_CPP_Std"
        Case &H3B: x = "Cvtpgd1300"
        Case &H3C: x = "Linker622"
        Case &H3D: x = "Linker700"
        Case &H3E: x = "Export622"
        Case &H3F: x = "Export700"
        Case &H40: x = "Masm700"
        Case &H41: x = "Utc13_POGO_I_C"
        Case &H42: x = "Utc13_POGO_I_CPP"
        Case &H43: x = "Utc13_POGO_O_C"
        Case &H44: x = "Utc13_POGO_O_CPP"
        Case &H45: x = "Cvtres700"
        Case &H46: x = "Cvtres710p"
        Case &H47: x = "Linker710p"
        Case &H48: x = "Cvtomf710p"
        Case &H49: x = "Export710p"
        Case &H4A: x = "Implib710p"
        Case &H4B: x = "Masm710p"
        Case &H4C: x = "Utc1310p_C"
        Case &H4D: x = "Utc1310p_CPP"
        Case &H4E: x = "Utc1310p_C_Std"
        Case &H4F: x = "Utc1310p_CPP_Std"
        Case &H50: x = "Utc1310p_LTCG_C"
        Case &H51: x = "Utc1310p_LTCG_CPP"
        Case &H52: x = "Utc1310p_POGO_I_C"
        Case &H53: x = "Utc1310p_POGO_I_CPP"
        Case &H54: x = "Utc1310p_POGO_O_C"
        Case &H55: x = "Utc1310p_POGO_O_CPP"
        Case &H56: x = "Linker624"
        Case &H57: x = "Cvtomf624"
        Case &H58: x = "Export624"
        Case &H59: x = "Implib624"
        Case &H5A: x = "Linker710"
        Case &H5B: x = "Cvtomf710"
        Case &H5C: x = "Export710"
        Case &H5D: x = "Implib710"
        Case &H5E: x = "Cvtres710"
        Case &H5F: x = "Utc1310_C"
        Case &H60: x = "Utc1310_CPP"
        Case &H61: x = "Utc1310_C_Std"
        Case &H62: x = "Utc1310_CPP_Std"
        Case &H63: x = "Utc1310_LTCG_C"
        Case &H64: x = "Utc1310_LTCG_CPP"
        Case &H65: x = "Utc1310_POGO_I_C"
        Case &H66: x = "Utc1310_POGO_I_CPP"
        Case &H67: x = "Utc1310_POGO_O_C"
        Case &H68: x = "Utc1310_POGO_O_CPP"
        Case &H69: x = "AliasObj710"
        Case &H6A: x = "AliasObj710p"
        Case &H6B: x = "Cvtpgd1310"
        Case &H6C: x = "Cvtpgd1310p"
        Case &H6D: x = "Utc1400_C"
        Case &H6E: x = "Utc1400_CPP"
        Case &H6F: x = "Utc1400_C_Std"
        Case &H70: x = "Utc1400_CPP_Std"
        Case &H71: x = "Utc1400_LTCG_C"
        Case &H72: x = "Utc1400_LTCG_CPP"
        Case &H73: x = "Utc1400_POGO_I_C"
        Case &H74: x = "Utc1400_POGO_I_CPP"
        Case &H75: x = "Utc1400_POGO_O_C"
        Case &H76: x = "Utc1400_POGO_O_CPP"
        Case &H77: x = "Cvtpgd1400"
        Case &H78: x = "Linker800"
        Case &H79: x = "Cvtomf800"
        Case &H7A: x = "Export800"
        Case &H7B: x = "Implib800"
        Case &H7C: x = "Cvtres800"
        Case &H7D: x = "Masm800"
        Case &H7E: x = "AliasObj800"
        Case &H7F: x = "PhoenixPrerelease"
        Case &H80: x = "Utc1400_CVTCIL_C"
        Case &H81: x = "Utc1400_CVTCIL_CPP"
        Case &H82: x = "Utc1400_LTCG_MSIL"
        Case &H83: x = "Utc1500_C"
        Case &H84: x = "Utc1500_CPP"
        Case &H85: x = "Utc1500_C_Std"
        Case &H86: x = "Utc1500_CPP_Std"
        Case &H87: x = "Utc1500_CVTCIL_C"
        Case &H88: x = "Utc1500_CVTCIL_CPP"
        Case &H89: x = "Utc1500_LTCG_C"
        Case &H8A: x = "Utc1500_LTCG_CPP"
        Case &H8B: x = "Utc1500_LTCG_MSIL"
        Case &H8C: x = "Utc1500_POGO_I_C"
        Case &H8D: x = "Utc1500_POGO_I_CPP"
        Case &H8E: x = "Utc1500_POGO_O_C"
        Case &H8F: x = "Utc1500_POGO_O_CPP"
        Case &H90: x = "Cvtpgd1500"
        Case &H91: x = "Linker900"
        Case &H92: x = "Export900"
        Case &H93: x = "Implib900"
        Case &H94: x = "Cvtres900"
        Case &H95: x = "Masm900"
        Case &H96: x = "AliasObj900"
        Case &H97: x = "Resource"
        Case &H98: x = "AliasObj1000"
        Case &H99: x = "Cvtpgd1600"
        Case &H9A: x = "Cvtres1000"
        Case &H9B: x = "Export1000"
        Case &H9C: x = "Implib1000"
        Case &H9D: x = "Linker1000"
        Case &H9E: x = "Masm1000"
        Case &H9F: x = "Phx1600_C"
        Case &HA0: x = "Phx1600_CPP"
        Case &HA1: x = "Phx1600_CVTCIL_C"
        Case &HA2: x = "Phx1600_CVTCIL_CPP"
        Case &HA3: x = "Phx1600_LTCG_C"
        Case &HA4: x = "Phx1600_LTCG_CPP"
        Case &HA5: x = "Phx1600_LTCG_MSIL"
        Case &HA6: x = "Phx1600_POGO_I_C"
        Case &HA7: x = "Phx1600_POGO_I_CPP"
        Case &HA8: x = "Phx1600_POGO_O_C"
        Case &HA9: x = "Phx1600_POGO_O_CPP"
        Case &HAA: x = "Utc1600_C"
        Case &HAB: x = "Utc1600_CPP"
        Case &HAC: x = "Utc1600_CVTCIL_C"
        Case &HAD: x = "Utc1600_CVTCIL_CPP"
        Case &HAE: x = "Utc1600_LTCG_C"
        Case &HAF: x = "Utc1600_LTCG_CPP"
        Case &HB0: x = "Utc1600_LTCG_MSIL"
        Case &HB1: x = "Utc1600_POGO_I_C"
        Case &HB2: x = "Utc1600_POGO_I_CPP"
        Case &HB3: x = "Utc1600_POGO_O_C"
        Case &HB4: x = "Utc1600_POGO_O_CPP"
        Case &HB5: x = "AliasObj1010"
        Case &HB6: x = "Cvtpgd1610"
        Case &HB7: x = "Cvtres1010"
        Case &HB8: x = "Export1010"
        Case &HB9: x = "Implib1010"
        Case &HBA: x = "Linker1010"
        Case &HBB: x = "Masm1010"
        Case &HBC: x = "Utc1610_C"
        Case &HBD: x = "Utc1610_CPP"
        Case &HBE: x = "Utc1610_CVTCIL_C"
        Case &HBF: x = "Utc1610_CVTCIL_CPP"
        Case &HC0: x = "Utc1610_LTCG_C"
        Case &HC1: x = "Utc1610_LTCG_CPP"
        Case &HC2: x = "Utc1610_LTCG_MSIL"
        Case &HC3: x = "Utc1610_POGO_I_C"
        Case &HC4: x = "Utc1610_POGO_I_CPP"
        Case &HC5: x = "Utc1610_POGO_O_C"
        Case &HC6: x = "Utc1610_POGO_O_CPP"
        Case &HC7: x = "AliasObj1100"
        Case &HC8: x = "Cvtpgd1700"
        Case &HC9: x = "Cvtres1100"
        Case &HCA: x = "Export1100"
        Case &HCB: x = "Implib1100"
        Case &HCC: x = "Linker1100"
        Case &HCD: x = "Masm1100"
        Case &HCE: x = "Utc1700_C"
        Case &HCF: x = "Utc1700_CPP"
        Case &HD0: x = "Utc1700_CVTCIL_C"
        Case &HD1: x = "Utc1700_CVTCIL_CPP"
        Case &HD2: x = "Utc1700_LTCG_C"
        Case &HD3: x = "Utc1700_LTCG_CPP"
        Case &HD4: x = "Utc1700_LTCG_MSIL"
        Case &HD5: x = "Utc1700_POGO_I_C"
        Case &HD6: x = "Utc1700_POGO_I_CPP"
        Case &HD7: x = "Utc1700_POGO_O_C"
        Case &HD8: x = "Utc1700_POGO_O_CPP"
        Case &HD9: x = "AliasObj1200"
        Case &HDA: x = "Cvtpgd1800"
        Case &HDB: x = "Cvtres1200"
        Case &HDC: x = "Export1200"
        Case &HDD: x = "Implib1200"
        Case &HDE: x = "Linker1200"
        Case &HDF: x = "Masm1200"
        Case &HE0: x = "Utc1800_C"
        Case &HE1: x = "Utc1800_CPP"
        Case &HE2: x = "Utc1800_CVTCIL_C"
        Case &HD3: x = "Utc1800_CVTCIL_CPP"
        Case &HE4: x = "Utc1800_LTCG_C"
        Case &HE5: x = "Utc1800_LTCG_CPP"
        Case &HE6: x = "Utc1800_LTCG_MSIL"
        Case &HE7: x = "Utc1800_POGO_I_C"
        Case &HE8: x = "Utc1800_POGO_I_CPP"
        Case &HE9: x = "Utc1800_POGO_O_C"
        Case &HEA: x = "Utc1800_POGO_O_CPP"
        Case &HEB: x = "AliasObj1210"
        Case &HEC: x = "Cvtpgd1810"
        Case &HED: x = "Cvtres1210"
        Case &HEE: x = "Export1210"
        Case &HEF: x = "Implib1210"
        Case &HF0: x = "Linker1210"
        Case &HF1: x = "Masm1210"
        Case &HF2: x = "Utc1810_C"
        Case &HF3: x = "Utc1810_CPP"
        Case &HF4: x = "Utc1810_CVTCIL_C"
        Case &HF5: x = "Utc1810_CVTCIL_CPP"
        Case &HF6: x = "Utc1810_LTCG_C"
        Case &HF7: x = "Utc1810_LTCG_CPP"
        Case &HF8: x = "Utc1810_LTCG_MSIL"
        Case &HF9: x = "Utc1810_POGO_I_C"
        Case &HFA: x = "Utc1810_POGO_I_CPP"
        Case &HFB: x = "Utc1810_POGO_O_C"
        Case &HFC: x = "Utc1810_POGO_O_CPP"
        Case &HFD: x = "AliasObj1400"
        Case &HFE: x = "Cvtpgd1900"
        Case &HFF: x = "Cvtres1400"
        Case &H100: x = "Export1400"
        Case &H101: x = "Implib1400"
        Case &H102: x = "Linker1400"
        Case &H103: x = "Masm1400"
        Case &H104: x = "Utc1900_C"
        Case &H105: x = "Utc1900_CPP"
        Case &H106: x = "Utc1900_CVTCIL_C"
        Case &H107: x = "Utc1900_CVTCIL_CPP"
        Case &H108: x = "Utc1900_LTCG_C"
        Case &H109: x = "Utc1900_LTCG_CPP"
        Case &H10A: x = "Utc1900_LTCG_MSIL"
        Case &H10B: x = "Utc1900_POGO_I_C"
        Case &H10C: x = "Utc1900_POGO_I_CPP"
        Case &H10D: x = "Utc1900_POGO_O_C"
        Case &H10E: x = "Utc1900_POGO_O_CPP"
    End Select

    toolName = x
    
End Property


'Function toolID_2_enum()
'
'    ex input: Case &H2F: x = "Utc12_2_Basic"
'
'    Dim s As New StringEx, v As New StringEx, ret() As String
'    Dim tmp() As String, x, v, n
'
'    tmp = Split(Clipboard.GetText, vbCrLf)
'    For Each x In tmp
'        s = x
'        v = "&h" & s.Extract("&h", ":", , , vbTextCompare).Trim()
'        n = "rh_" & s.Extract("""", """")
'        Debug.Print n & " = " & v
'        push ret, n & " = " & v
'    Next
'
'    Clipboard.Clear
'    Clipboard.SetText Join(ret, vbCrLf)
'    End
'
'End Function


