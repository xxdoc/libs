VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInflation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Tom's Inflation Calculator
' Copyright 2019, Tom R. Halfhill
' JavaScript version 3.14
' https://www.halfhill.com/inflation_js.html
'
' converted from JS -dzzie 2.5.19
' other inflation calculations not yet ported

'Usage:
'    Dim xx As Double
'    If cpi.calculate(100, 1980, 2018, xx) Then
'        MsgBox xx
'    Else
'        MsgBox cpi.errMsg
'    End If
'  - or -
'    MsgBox cpi.calculate_(100, 1980, 2018) '0 on error

    
Const oldest_year = 1665
Const OLDEST_CPIU_YEAR = 1913
Const latest_year = 2100
Const LATEST_DATA_YEAR = 2018 ' // UPDATE YEARLY!!!!!!
Const NAME = "Consumer Price Index (Urban), 1665-2100"
Const DATA_SOURCE = "U.S. Bureau of Labor Statistics 1913-2018"

Private sum  'used later to compute CPI-U average
Private average
Private rates(450) As Double 'Consumer Price Index - Urban

Public errMsg As String
Public dollarsSanitized As Long

Private Sub Class_Initialize()

    rates(0) = 0       ' 1666 = Economic History Services data
    rates(1) = 7.55    ' 1667 http:'www.eh.net/hmit/
    rates(2) = 3.51    ' 1668
    rates(3) = -6.78   ' 1669
    rates(4) = -1.82   ' 1670
    rates(5) = 1.85    ' 1671
    rates(6) = -0.91   ' 1672
    rates(7) = -4.59   ' 1673
    rates(8) = 7.69    ' 1674
    rates(9) = -13.39  ' 1675
    rates(10) = 1.03   ' 1676
    rates(11) = 1.02   ' 1677
    rates(12) = -2.02  ' 1678
    rates(13) = -1.03  ' 1679
    rates(14) = 14.58  ' 1680
    rates(15) = 2.73   ' 1681
    rates(16) = -18.58 ' 1682
    rates(17) = 0      ' 1683
    rates(18) = 0      ' 1684
    rates(19) = 7.61   ' 1685
    rates(20) = -6.06  ' 1686
    rates(21) = 0      ' 1687
    rates(22) = -8.6   ' 1688
    rates(23) = 2.35   ' 1689
    rates(24) = 2.3    ' 1690
    rates(25) = 4.49   ' 1691
    rates(26) = -5.38  ' 1692
    rates(27) = -5.68  ' 1693
    rates(28) = 6.02   ' 1694
    rates(29) = -7.95  ' 1695
    rates(30) = 17.28  ' 1696
    rates(31) = -2.11  ' 1697
    rates(32) = -5.38  ' 1698
    rates(33) = 9.09   ' 1699
    rates(34) = -3.13  ' 1700
    rates(35) = 5.38   ' 1701
    rates(36) = -8.16  ' 1702
    rates(37) = -1.11  ' 1703
    rates(38) = 13.48  ' 1704
    rates(39) = -9.9   ' 1705
    rates(40) = 7.69   ' 1706
    rates(41) = -4.08  ' 1707
    rates(42) = -21.28 ' 1708
    rates(43) = 20.27  ' 1709
    rates(44) = -10.11 ' 1710
    rates(45) = -1.25  ' 1711
    rates(46) = -5.06  ' 1712
    rates(47) = 6.67   ' 1713
    rates(48) = -8.75  ' 1714
    rates(49) = 12.33  ' 1715
    rates(50) = -6.1   ' 1716
    rates(51) = 6.49   ' 1717
    rates(52) = -6.1   ' 1718
    rates(53) = 6.49   ' 1719
    rates(54) = -7.32  ' 1720
    rates(55) = -6.58  ' 1721
    rates(56) = 5.63   ' 1722
    rates(57) = 1.33   ' 1723
    rates(58) = 3.95   ' 1724
    rates(59) = 13.92  ' 1725
    rates(60) = -7.78  ' 1726
    rates(61) = 3.61   ' 1727
    rates(62) = -5.81  ' 1728
    rates(63) = -1.23  ' 1729
    rates(64) = 0      ' 1730
    rates(65) = -11.25 ' 1731
    rates(66) = -7.04  ' 1732
    rates(67) = 3.03   ' 1733
    rates(68) = -2.94  ' 1734
    rates(69) = 3.03   ' 1735
    rates(70) = -4.41  ' 1736
    rates(71) = 3.08   ' 1737
    rates(72) = 4.48   ' 1738
    rates(73) = -11.43 ' 1739
    rates(74) = 6.45   ' 1740
    rates(75) = 37.88  ' 1741
    rates(76) = -10.99 ' 1742
    rates(77) = -12.35 ' 1743
    rates(78) = -7.04  ' 1744
    rates(79) = -3.03  ' 1745
    rates(80) = 1.56   ' 1746
    rates(81) = 10.77  ' 1747
    rates(82) = 13.89  ' 1748
    rates(83) = 2.44   ' 1749
    rates(84) = 0      ' 1750
    rates(85) = 1.19   ' 1751
    rates(86) = 1.18   ' 1752
    rates(87) = -2.33  ' 1753
    rates(88) = -3.57  ' 1754
    rates(89) = -2.47  ' 1755
    rates(90) = -2.53  ' 1756
    rates(91) = 5.19   ' 1757
    rates(92) = 7.41   ' 1758
    rates(93) = 13.79  ' 1759
    rates(94) = -3.03  ' 1760
    rates(95) = -8.33  ' 1761
    rates(96) = 7.95   ' 1762
    rates(97) = 0      ' 1763
    rates(98) = -7.37  ' 1764
    rates(99) = 1.14   ' 1765
    rates(100) = 8.99  ' 1766
    rates(101) = -2.06 ' 1767
    rates(102) = -5.26 ' 1768
    rates(103) = 3.33  ' 1769
    rates(104) = 7.53  ' 1770
    rates(105) = -4#   ' 1771
    rates(106) = 14.58 ' 1772
    rates(107) = -7.27 ' 1773
    rates(108) = -4.9  ' 1774
    rates(109) = -5.15 ' 1775
    rates(110) = 14.13 ' 1776
    rates(111) = 21.9  ' 1777
    rates(112) = 29.69 ' 1778
    rates(113) = -11.45 ' 1779
    rates(114) = 12.24 ' 1780
    rates(115) = -19.39 ' 1781
    rates(116) = 9.77  ' 1782
    rates(117) = -12.33 ' 1783
    rates(118) = -3.91 ' 1784
    rates(119) = -4.88 ' 1785
    rates(120) = -2.56 ' 1786
    rates(121) = -1.75 ' 1787
    rates(122) = -4.46 ' 1788
    rates(123) = -0.93 ' 1789
    rates(124) = 3.77  ' 1790
    rates(125) = 2.73  ' 1791
    rates(126) = 1.77  ' 1792
    rates(127) = 3.48  ' 1793
    rates(128) = 10.92 ' 1794
    rates(129) = 14.39 ' 1795
    rates(130) = 5.3   ' 1796
    rates(131) = -3.77 ' 1797
    rates(132) = -3.27 ' 1798
    rates(133) = 0     ' 1799
    rates(134) = 2.03  ' 1800
    rates(135) = 1.32  ' 1801
    rates(136) = -15.69 ' 1802
    rates(137) = 5.43  ' 1803
    rates(138) = 4.41  ' 1804
    rates(139) = -0.7  ' 1805
    rates(140) = 4.26  ' 1806
    rates(141) = -5.44 ' 1807
    rates(142) = 8.63  ' 1808
    rates(143) = -1.99 ' 1809
    rates(144) = 0     ' 1810
    rates(145) = 6.76  ' 1811
    rates(146) = 1.27  ' 1812
    rates(147) = 20#   ' 1813
    rates(148) = 9.9   ' 1814
    rates(149) = -12.32 ' 1815
    rates(150) = -8.65 ' 1816
    rates(151) = -5.33 ' 1817
    rates(152) = -4.38 ' 1818
    rates(153) = 0     ' 1819
    rates(154) = -7.84 ' 1820
    rates(155) = -4.61 ' 1821 = Federal Reserve Bank data...
    rates(156) = 3.22  ' 1822 (Global Financial Data)
    rates(157) = -4.68 ' 1823
    rates(158) = -6.56 ' 1824
    rates(159) = 1.76  ' 1825
    rates(160) = -5.17 ' 1826
    rates(161) = 3.63  ' 1827
    rates(162) = 0     ' 1828
    rates(163) = 1.76  ' 1829
    rates(164) = -6.9  ' 1830
    rates(165) = 3.7   ' 1831
    rates(166) = 1.79  ' 1832
    rates(167) = 1.76  ' 1833
    rates(168) = -12.07 ' 1834
    rates(169) = 17.64 ' 1835
    rates(170) = 13.34 ' 1836
    rates(171) = 5.88  ' 1837
    rates(172) = -1.39 ' 1838
    rates(173) = 0     ' 1839
    rates(174) = -15.49 ' 1840
    rates(175) = 0     ' 1841
    rates(176) = -8.32 ' 1842
    rates(177) = -7.27 ' 1843
    rates(178) = 1.95  ' 1844
    rates(179) = 3.86  ' 1845
    rates(180) = 7.41  ' 1846
    rates(181) = 0     ' 1847
    rates(182) = -12.07 ' 1848
    rates(183) = 5.88  ' 1849
    rates(184) = -5.55 ' 1850
    rates(185) = 17.64 ' 1851
    rates(186) = 0     ' 1852
    rates(187) = 6.67  ' 1853
    rates(188) = 0     ' 1854
    rates(189) = 4.68  ' 1855
    rates(190) = 1.5   ' 1856
    rates(191) = 2.93  ' 1857
    rates(192) = -1.42 ' 1858
    rates(193) = -8.71 ' 1859
    rates(194) = -3.17 ' 1860
    rates(195) = 3.27  ' 1861
    rates(196) = 9.54  ' 1862
    rates(197) = 13.04 ' 1863
    rates(198) = 21.79 ' 1864
    rates(199) = 7.38  ' 1865
    rates(200) = 0.97  ' 1866
    rates(201) = -0.96 ' 1867
    rates(202) = -3.92 ' 1868
    rates(203) = -3.07 ' 1869
    rates(204) = -4.2  ' 1870
    rates(205) = -2.2  ' 1871
    rates(206) = 1.13  ' 1872
    rates(207) = -2.23 ' 1873
    rates(208) = 0     ' 1874
    rates(209) = -4.59 ' 1875
    rates(210) = -3.66 ' 1876
    rates(211) = -6.33 ' 1877
    rates(212) = -6.76 ' 1878
    rates(213) = 7.25  ' 1879
    rates(214) = 4.05  ' 1880
    rates(215) = 5.19  ' 1881
    rates(216) = -2.47 ' 1882
    rates(217) = -2.53 ' 1883
    rates(218) = -3.9  ' 1884
    rates(219) = -1.35 ' 1885
    rates(220) = 0     ' 1886
    rates(221) = 2.74  ' 1887
    rates(222) = 0     ' 1888
    rates(223) = -1.33 ' 1889
    rates(224) = 1.35  ' 1890
    rates(225) = 0     ' 1891
    rates(226) = -1.33 ' 1892
    rates(227) = 0     ' 1893
    rates(228) = -2.7  ' 1894
    rates(229) = -1.39 ' 1895
    rates(230) = 0     ' 1896
    rates(231) = 0     ' 1897
    rates(232) = 1.41  ' 1898
    rates(233) = 4.17  ' 1899
    rates(234) = 1.33  ' 1900
    rates(235) = 2.63  ' 1901
    rates(236) = 2.56  ' 1902
    rates(237) = 0     ' 1903
    rates(238) = 1.25  ' 1904
    rates(239) = 2.47  ' 1905
    rates(240) = 4.82  ' 1906
    rates(241) = 3.45  ' 1907
    rates(242) = 0     ' 1908
    rates(243) = 5.56  ' 1909
    rates(244) = 0     ' 1910
    rates(245) = 2.11  ' 1911
    rates(246) = 3.09  ' 1912
    rates(247) = 2.04  ' 1913
    rates(248) = 1#    ' 1914 = U.S. Consumer Price Index (Urban)...
    rates(249) = 1#    ' 1915
    rates(250) = 7.9   ' 1916
    rates(251) = 17.4  ' 1917
    rates(252) = 18#   ' 1918
    rates(253) = 14.6  ' 1919
    rates(254) = 15.6  ' 1920
    rates(255) = -10.5 ' 1921
    rates(256) = -6.1  ' 1922
    rates(257) = 1.8   ' 1923
    rates(258) = 0     ' 1924
    rates(259) = 2.3   ' 1925
    rates(260) = 1.1   ' 1926
    rates(261) = -1.7  ' 1927
    rates(262) = -1.7  ' 1928
    rates(263) = 0     ' 1929
    rates(264) = -2.3  ' 1930
    rates(265) = -9#   ' 1931
    rates(266) = -9.9  ' 1932
    rates(267) = -5.1  ' 1933
    rates(268) = 3.1   ' 1934
    rates(269) = 2.2   ' 1935
    rates(270) = 1.5   ' 1936
    rates(271) = 3.6   ' 1937
    rates(272) = -2.1  ' 1938
    rates(273) = -1.4  ' 1939
    rates(274) = 0.7   ' 1940
    rates(275) = 5#    ' 1941
    rates(276) = 10.9  ' 1942
    rates(277) = 6.1   ' 1943
    rates(278) = 1.7   ' 1944
    rates(279) = 2.3   ' 1945
    rates(280) = 8.3   ' 1946
    rates(281) = 14.4  ' 1947
    rates(282) = 8.1   ' 1948
    rates(283) = -1.2  ' 1949
    rates(284) = 1.3   ' 1950
    rates(285) = 7.9   ' 1951
    rates(286) = 1.9   ' 1952
    rates(287) = 0.8   ' 1953
    rates(288) = 0.7   ' 1954
    rates(289) = -0.4  ' 1955
    rates(290) = 1.5   ' 1956
    rates(291) = 3.3   ' 1957
    rates(292) = 2.8   ' 1958
    rates(293) = 0.7   ' 1959
    rates(294) = 1.7   ' 1960
    rates(295) = 1#    ' 1961
    rates(296) = 1#    ' 1962
    rates(297) = 1.3   ' 1963
    rates(298) = 1.3   ' 1964
    rates(299) = 1.6   ' 1965
    rates(300) = 2.9   ' 1966
    rates(301) = 3.1   ' 1967
    rates(302) = 4.2   ' 1968
    rates(303) = 5.5   ' 1969
    rates(304) = 5.7   ' 1970
    rates(305) = 4.4   ' 1971
    rates(306) = 3.2   ' 1972
    rates(307) = 6.2   ' 1973
    rates(308) = 11#   ' 1974
    rates(309) = 9.1   ' 1975
    rates(310) = 5.8   ' 1976
    rates(311) = 6.5   ' 1977
    rates(312) = 7.6   ' 1978
    rates(313) = 11.3  ' 1979
    rates(314) = 13.5  ' 1980
    rates(315) = 10.3  ' 1981
    rates(316) = 6.2   ' 1982
    rates(317) = 3.2   ' 1983
    rates(318) = 4.3   ' 1984
    rates(319) = 3.6   ' 1985
    rates(320) = 1.9   ' 1986
    rates(321) = 3.6   ' 1987
    rates(322) = 4.1   ' 1988
    rates(323) = 4.8   ' 1989
    rates(324) = 5.4   ' 1990
    rates(325) = 4.2   ' 1991
    rates(326) = 3#    ' 1992
    rates(327) = 3#    ' 1993
    rates(328) = 2.6   ' 1994
    rates(329) = 2.8   ' 1995
    rates(330) = 3#    ' 1996
    rates(331) = 2.3   ' 1997
    rates(332) = 1.6   ' 1998
    rates(333) = 2.2   ' 1999
    rates(334) = 3.4   ' 2000
    rates(335) = 2.8   ' 2001
    rates(336) = 1.6   ' 2002
    rates(337) = 2.3   ' 2003
    rates(338) = 2.7   ' 2004
    rates(339) = 3.4   ' 2005
    rates(340) = 3.2   ' 2006
    rates(341) = 2.8   ' 2007
    rates(342) = 3.8   ' 2008
    rates(343) = -0.4  ' 2009
    rates(344) = 1.6   ' 2010
    rates(345) = 3.2   ' 2011
    rates(346) = 2.1   ' 2012
    rates(347) = 1.5   ' 2013
    rates(348) = 1.6   ' 2014
    rates(349) = 0.1   ' 2015
    rates(350) = 1.3   ' 2016
    rates(351) = 2.1   ' 2017
    rates(352) = 2.4   ' 2018 = Latest data year
    rates(353) = 2.2   ' 2019 = CBO forecasts (August 2018 Economic Outlook)...
    rates(354) = 2.5   ' 2020
    rates(355) = 2.5   ' 2021
    rates(356) = 2.5   ' 2022
    rates(357) = 2.4   ' 2023
    rates(358) = 2.4   ' 2024
    rates(359) = 2.4   ' 2025
    rates(360) = 2.4   ' 2026
    rates(361) = 2.4   ' 2027
    rates(362) = 2.4   ' 2028

    Dim x As Long
    'UPDATE THIS CODE YEARLY -----------------------------------------
    'Compute CPI-U average, 1913 to present:
    x = 248
    While x <= 352
        sum = sum + rates(x)
        x = x + 1
    Wend
    average = sum / (LATEST_DATA_YEAR - OLDEST_CPIU_YEAR) '// Example: 2018 minus 1913
    
    'Fill array to year 2100 with inflation estimate, UPDATE x TO FIRST EMPTY ARRAY ELEMENT!!!:
    x = 363
    While x <= (latest_year - oldest_year)
        rates(x) = average
        x = x + 1
    Wend
    'END OF CODE TO UPDATE YEARLY ------------------------------------
    
    'Convert percentages...
    x = 0
    While x <= (latest_year - oldest_year)
        rates(x) = rates(x) / 100
        x = x + 1
    Wend
    'end of Consumer Price Index initialization

End Sub

Private Sub checkStartYearInput(startYear, oldest_year, latest_year)
    
    If startYear < oldest_year Then Err.Raise 1, , "Error: starting year is too long ago. Try a later year."
    If startYear > latest_year Then Err.Raise 1, , "Error: inflation data for the starting year is not yet available. Try an earlier year."

End Sub

Private Sub checkTargetYearInput(targetYear, oldest_year, latest_year)

     If targetYear < oldest_year Then Err.Raise 1, , "Error: target year is too long ago. Try a later year."
     If targetYear > latest_year Then Err.Raise 1, , "Error: inflation data for the target year is not yet available. Try an earlier year."
     
End Sub

Private Function CPIUinflation(dollars, startYear, targetYear)


    Dim yearGap As Long
    Dim x As Long
    
    If startYear < targetYear Then '   // calculate forward in time...
        yearGap = targetYear - startYear
        While x < yearGap
            dollars = dollars + (dollars * rates((startYear - oldest_year) + x))
            x = x + 1
        Wend
    End If
    
    If startYear > targetYear Then  'calculate backward in time...
        yearGap = startYear - targetYear
        x = 1
        While x <= yearGap
            dollars = dollars / (1 + rates((startYear - oldest_year) - x))
            x = x + 1
        Wend
    End If
    
    CPIUinflation = Round(dollars) ', 2)

End Function

'we allow startYear to possibly be a date string or even a sql wildcard type date strig
'no wildcards in year though..expects 4 digit year..
Function calculate(ByVal dollars, ByVal startYear, ByVal targetYear As Long, ByRef outVal As Double) As Boolean
    
    On Error GoTo hell
    
    outVal = 0
    errMsg = Empty
    startYear = Trim(CStr(startYear))
    dollarsSanitized = 0
    
    dollars = Replace(dollars, ",", Empty)
    dollars = Replace(dollars, "$", Empty)
    If InStr(dollars, ".") > 0 Then dollars = Split(dollars, ".")(0)
    dollarsSanitized = CLng(dollars)
    
    If InStr(startYear, "*") > 0 Then startYear = Replace(startYear, "*", Empty)
    If InStr(startYear, "%") > 0 Then startYear = Replace(startYear, "%", Empty)
    If InStr(startYear, "/") > 0 Then startYear = Year(startYear)
    startYear = CLng(startYear)
    
    checkStartYearInput startYear, oldest_year, latest_year
    checkTargetYearInput targetYear, oldest_year, latest_year
    outVal = CPIUinflation(dollars, startYear, targetYear)
    calculate = True
    
    Exit Function
hell:
    errMsg = Err.Description
End Function

Function calculate_(ByVal dollars, ByVal startYear, ByVal targetYear As Long) As Double
    Dim outVal As Double
    calculate dollars, startYear, targetYear, outVal
    calculate_ = outVal
End Function











