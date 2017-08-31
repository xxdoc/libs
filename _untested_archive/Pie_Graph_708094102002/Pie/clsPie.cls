VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'Name:          Pie Class                       Written By:    Trevor Newsome
'Module Name:   clsPie
'Date:          10/4/2002                       Version:       V1.0
'Language:      Visial Basic V6                 OS:            Windows 9x
'Email:         trevor.newsome@btopenworld.com
'Description:   This will provide a method to draw 2D pie charts by building segments
'               with given values
'---------------------------------------------------------------------------------------

Option Explicit

'Declare constants
Private Const OUT_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0
Private Const DEFAULT_CHARSET = 1
Private Const LF_FACESIZE = 32

'Declare Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type Segment
    Value As Double
    Colour As Long
    Name As String
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

'Declare DLL calls
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Misc declarations
Private lf As LOGFONT

Private intSegments As Long

Private PieChart() As Segment

'This will draw the legend with the colours specified and the segment labels
Private Sub DrawLegend(ByVal lngColour, ByVal intYPosition, ByRef lngPichDC As Long, ByRef lngPichwnd As Long)
    Dim FontToUse As Long
    Dim rc As RECT
    Dim oldhdc As Long
    Dim dl As Long
    Dim lnghBrush As Long
    Dim tmpString As String
    
    On Error GoTo errHandle
    
    'Create fill colour
    lnghBrush = CreateSolidBrush(lngColour)
    SelectObject lngPichDC, lnghBrush
    
    'Draw the legend item in the above colour
    Rectangle lngPichDC, 220, 50 + (15 * intYPosition), 240, 60 + (15 * intYPosition)
    
    'Sort out the font details
    lf.lfHeight = 7: lf.lfWidth = 5: lf.lfEscapement = 0: lf.lfWeight = 800
    lf.lfItalic = 0: lf.lfUnderline = 0: lf.lfStrikeOut = 0
    lf.lfOutPrecision = OUT_DEFAULT_PRECIS: lf.lfClipPrecision = OUT_DEFAULT_PRECIS
    lf.lfQuality = DEFAULT_QUALITY: lf.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    lf.lfCharSet = DEFAULT_CHARSET
    
    'Create Font
    FontToUse = CreateFontIndirect(lf)
    
    'Print the font to the picture box
    oldhdc = SelectObject(lngPichDC, FontToUse)
    dl = GetClientRect(lngPichwnd, rc)
    tmpString = PieChart(intYPosition).Name & " (" & PieChart(intYPosition).Value & ")"
    dl = TextOut(lngPichDC, 245, 52 + (15 * intYPosition), tmpString, Len(tmpString))
    dl = TextOut(lngPichDC, 245, 40, "KEY", Len("KEY"))
    dl = SelectObject(lngPichDC, oldhdc)

Exit Sub

errHandle:
    Exit Sub
End Sub

'Draws a segment with a passed in parameter of the percentage of the total pie
Private Sub DrawSegment(ByVal intPerc As Integer, ByVal lngColour As Long, ByRef lngPichDC As Long)
    Const pi = 3.14159265358979
    Dim X1, Y1, X2, Y2, x3, y3, x4, y4, rad, theta, beta As Double
    Dim lnghBrush As Long

    On Error GoTo errHandle
    
    X1 = 0: Y1 = 20: X2 = 200: Y2 = 220  'Initial Circle points
    x3 = (X2 - X1) / 2: y4 = Y1 'Initial the first point being at North point on the circle
    rad = (X2 - X1) / 2 'Get radius of circle
    theta = (intPerc / 100) * 360  'Get theta from the percentage of the pie segment passed as a parameter
    beta = 180 - theta - 90 'This gets the missing angle from the RHT assuming the segament is <90 degrees
    x4 = rad + ((rad * (Sin(theta * (pi / 180)))) * 180 / pi)
    y4 = rad - ((rad * (Sin(beta * (pi / 180)))) * 180 / pi) 'Converts from radians and gets y4 point
    
    'Automatically fill the segment
    lnghBrush = CreateSolidBrush(lngColour)
    SelectObject lngPichDC, lnghBrush
    
    'Draw the segment
    Pie lngPichDC, CLng(X1), CLng(Y1), CLng(X2), CLng(Y2), CLng(x4), CLng(y4), CLng(x3), CLng(y3) 'Draw Pie Swapped x3y3 for x4y4 because i want the smaller segment
    
Exit Sub

errHandle:
    Exit Sub
End Sub

'Draws the whole pie chart from repeatedly calling the drawsegments for every segment
Public Function DrawPie(lngPichDC As Long, lngPichwnd As Long, blnShowLegend, strGraphTitle As String) As Long
    Dim PieTotal As Double
    Dim intCount As Integer
    Dim SegmentTotal As Double
    Dim FontToUse As Long
    Dim rc As RECT
    Dim oldhdc As Long
    Dim dl As Long
    
    On Error GoTo errHandle
    
    'Gets Pie segments total
    PieTotal = 0
    For intCount = 0 To UBound(PieChart): PieTotal = PieTotal + PieChart(intCount).Value: Next intCount
    
    'Draw every segment
    SegmentTotal = 100
    If blnShowLegend Then DrawLegend PieChart(0).Colour, 0, lngPichDC, lngPichwnd
    DrawSegment SegmentTotal, PieChart(0).Colour, lngPichDC
    Exit Function
    
    'Loop for as many segments there are, although not infinite
    For intCount = 0 To UBound(PieChart) - 1
        SegmentTotal = SegmentTotal - ((PieChart(intCount).Value / PieTotal) * 100)
        'Draw Legend
        If blnShowLegend Then DrawLegend PieChart(intCount + 1).Colour, intCount + 1, lngPichDC, lngPichDC
        'Create and draw the segment
        DrawSegment SegmentTotal, PieChart(intCount + 1).Colour, lngPichDC
    Next intCount
    
    'Place title on screen
    'Sort out the font details
    lf.lfHeight = 10: lf.lfWidth = 6: lf.lfEscapement = 0: lf.lfWeight = 800
    lf.lfItalic = 0: lf.lfUnderline = 1: lf.lfStrikeOut = 0
    lf.lfOutPrecision = OUT_DEFAULT_PRECIS: lf.lfClipPrecision = OUT_DEFAULT_PRECIS
    lf.lfQuality = DEFAULT_QUALITY: lf.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    lf.lfCharSet = DEFAULT_CHARSET
    
    'Create font
    FontToUse = CreateFontIndirect(lf)
    
    'Print text to picture box
    oldhdc = SelectObject(lngPichDC, FontToUse)
    dl = GetClientRect(lngPichwnd, rc)
    dl = TextOut(lngPichDC, 20, 0, strGraphTitle, Len(strGraphTitle))
    dl = SelectObject(lngPichDC, oldhdc)
    
    DrawPie = 1
Exit Function

errHandle: '
    DrawPie = 2
    Exit Function
End Function

'This function will add a segment with relevant details ready for the draw routine to be called
Public Function AddSegment(dblPercentage As Double, strName As String, lngColour As Long) As Long
    intSegments = intSegments + 1   'Adds another segment
    If intSegments > 0 Then
        'Don't need to redimension if the first segment
        ReDim Preserve PieChart(intSegments)
    End If
    'copy the details to the array of tyoe PieChart
    PieChart(intSegments).Value = CLng(dblPercentage)
    PieChart(intSegments).Colour = lngColour
    PieChart(intSegments).Name = strName
End Function

'This clears all of the segments ready for the user to start all over
Public Function Clear()
    'Reset the segments
    intSegments = -1
    'Redimension and clear the array
    ReDim PieChart(0)
End Function

'Just to create PieChart array and init the intSegments
Private Sub Class_Initialize()
    Clear
End Sub

'---------------------------------------------------------------------------------------
'How to use this module
'---------------------------------------------------------------------------------------

'Create a new object of class clsPie on the form
    'Private PieGraph As New clsPie
    
'Build the pie up by adding individual segments
    'PieGraph.AddSegment 45, "Response", &HFF0000    'Blue
    'PieGraph.AddSegment 20, "Actual", &HFFFF&     'Yellow
    'PieGraph.AddSegment 100, "Delays", &HFF&      'Red
    'PieGraph.AddSegment 10, "Test", &HFF00FF        'Violet

'Draw the pie graphs built of the segments to a specified picture box
    'PieGraph.DrawPie picPie.hdc, picPie.hwnd, True, "A Graph To Show The Breakdowns"

'Clear the already added segments ready to create a new pie graph
    'PieGraph.Clear
'---------------------------------------------------------------------------------------
