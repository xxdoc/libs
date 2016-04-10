' $Id: VersionData.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' VersionData.bas - Get internal libcurl version information.

Attribute VB_Name = "VersionData"

Private Sub VersionData()
    Dim vd As Long, i As Long, lenProts As Long
    Dim prots() As String
    vd = vbcurl_version_info(CURLVERSION_NOW)
    Debug.Print "           Age: " & vbcurl_version_age(vd)
    Debug.Print "Version String: " & vbcurl_version_string(vd)
    Debug.Print "Version Number: " & vbcurl_version_num(vd)
    Debug.Print "   Host System: " & vbcurl_version_host(vd)
    Debug.Print "Feature Bitmap: " & vbcurl_version_features(vd)
    Debug.Print "   SSL Version: " & vbcurl_version_ssl(vd)
    Debug.Print "SSL VersionNum: " & vbcurl_version_ssl_num(vd)
    Debug.Print "  LibZ Version: " & vbcurl_version_libz(vd)
    Debug.Print "  ARES Version: " & vbcurl_version_ares(vd)
    Debug.Print "  ARES Ver Num: " & vbcurl_version_ares_num(vd)
    Debug.Print "libidn Version: " & vbcurl_version_libidn(vd)
    Debug.Print "Protocols:"
    vbcurl_version_protocols vd, prots
    lenProts = UBound(prots)
    For i = 0 To lenProts
        Debug.Print "  " & prots(i)
    Next
End Sub

