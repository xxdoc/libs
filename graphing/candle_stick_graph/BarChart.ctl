VERSION 5.00
Begin VB.UserControl BarChart 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   ScaleHeight     =   2910
   ScaleWidth      =   6480
   ToolboxBitmap   =   "BarChart.ctx":0000
   Begin VB.PictureBox picChart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   6405
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      Begin VB.Timer tmrLookup 
         Interval        =   100
         Left            =   120
         Top             =   7280
      End
   End
End
Attribute VB_Name = "BarChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'BARCHART ActiveX Control by Richard Gardner support@rgsoftware.com
'If you change this code for the better please let me know!

Option Explicit

Private aryDate As Variant
Private OHLCData As Variant
Private MaxData As Long 'Store the max length of data in an array
Private MaxValue As Double 'The maximum value of high data
Private StepVar As Double
Private GapVal As Double
Private SavedBarColor As Long
Private SavedBackColor As Long
Private NewLookup As Boolean
Private VarLineData As Variant
Private MaxLineData As Double
Private MaxLineValue As Double
Private SavedLineColor As Long
Private Repainting As Boolean
Private ArrayLineColor()
Private LineChartsArray() As Variant 'Holds the arrays of all the line charts in one array

Public Event MouseMove(Price As Double, ObservationDate As String)

Private Sub picChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strLocation As String
    On Error Resume Next
    If Not NewLookup Then Exit Sub
    NewLookup = False
    RaiseEvent MouseMove(CDbl(Format(MaxValue - Y, "##.000")), CStr(DateLookup(X)))
End Sub

Private Sub tmrLookup_Timer()
    NewLookup = True
End Sub

Private Sub UserControl_Initialize()
    ReDim LineChartsArray(1 To 1)
    ReDim ArrayLineColor(1 To 1)
    tmrLookup.Enabled = True
End Sub

Private Sub UserControl_Resize()
    Dim lngLoop As Long
    'Resize the chart and update the data
    picChart.Width = UserControl.Width
    picChart.Height = UserControl.Height
    ' On Error Resume Next
    picChart.Cls
    If SavedBarColor <> -1 Then DrawBarChart OHLCData, SavedBarColor
    For lngLoop = 1 To UBound(LineChartsArray) - 1
        'loop through all arrays
        Repainting = True
        DrawLineChart LineChartsArray(lngLoop), CLng(ArrayLineColor(lngLoop))
    Next lngLoop
End Sub

Public Function ClearChart()
    Dim lngLoop As Long
    picChart.Cls
    On Error Resume Next
    For lngLoop = 1 To UBound(LineChartsArray)
        LineChartsArray(lngLoop) = 0
    Next
    ReDim LineChartsArray(1 To 1)
End Function

Public Function DrawBarChart(Open_High_Low_Close_Date, Optional ChartBarColor As Long = vbBlack)

    On Error GoTo ErrHndl

    'Open_High_Low_Close_Date stores the Open, High, Low and Close prices and dates that will be displayed on the chart

    Dim lngLoop As Long
    Dim intCount As Integer

    'Treat OHLC as public data
    OHLCData = Open_High_Low_Close_Date
    SavedBarColor = ChartBarColor

    'Check the data to see if we have enough
    MaxData = Max()
    If MaxData = 0 Then Exit Function
    ReDim aryDate(1 To MaxData, 1 To 2)

    'Scale the chart down
    ScaleToData

    StepVar = GetStepVar

    If OHLCData(1, 4) = 0 Then Err.Raise 0

    For lngLoop = 1 To MaxData Step StepVar

        intCount = intCount + 1

        'See if there is any more data to be displayed
        If intCount > MaxData Then Exit For

        'save the date information
        aryDate(intCount, 1) = OHLCData(intCount, 5)
        aryDate(intCount, 2) = GetX(intCount)

        'Paint the high and low
        picChart.Line (GetX(intCount), GetY(OHLCData(intCount, 2)))- _
                (GetX(intCount), GetY(OHLCData(intCount, 3))), ChartBarColor

        'Paint the open
        If OHLCData(intCount, 1) <> 0 Then
            picChart.Line (GetX(intCount), GetY(OHLCData(intCount, 1)))-(GetX(intCount) - GapVal, _
                    GetY(OHLCData(intCount, 1))), ChartBarColor
        End If

        'Paint the close
        picChart.Line (GetX(intCount), GetY(OHLCData(intCount, 4)))-(GetX(intCount) + GapVal, _
                GetY(OHLCData(intCount, 4))), ChartBarColor

    Next lngLoop

    Exit Function
ErrHndl:
    Err.Raise "BarChartX: " & vbCrLf & Err.Description, vbExclamation
End Function

Private Function Max()

    'Finds how many bars hold a value

    Dim lngLoop As Long
    On Error Resume Next
    If OHLCData(1, 1) = 0 Then Exit Function
    For lngLoop = 1 To UBound(OHLCData)
        'Open is optional
        If OHLCData(lngLoop, 2) = 0 Then Exit For 'No High
        If OHLCData(lngLoop, 3) = 0 Then Exit For 'No Low
        If OHLCData(lngLoop, 4) = 0 Then Exit For 'No Close
    Next lngLoop

    Max = lngLoop - 1

End Function

Private Function ScaleToData()

    'Scales picChart so that the data will fit nicely

    Dim lngLoop As Long
    Dim dblMax As Double
    Dim dblMin As Double

    'Get the maximum value
    dblMax = OHLCData(1, 2)
    For lngLoop = 1 To MaxData
        If dblMax < OHLCData(lngLoop, 2) Then dblMax = OHLCData(lngLoop, 2)
    Next lngLoop

    'Get the minimum value
    dblMin = OHLCData(1, 3)
    For lngLoop = 1 To MaxData
        If OHLCData(lngLoop, 2) < dblMin Then dblMin = OHLCData(lngLoop, 3)
    Next lngLoop

    MaxValue = dblMax

    picChart.ScaleHeight = dblMax - dblMin
    picChart.ScaleWidth = MaxData + (MaxData * 0.5) - 2

    GapVal = picChart.ScaleWidth / MaxData / 3

End Function

Private Function GetStepVar()
    GetStepVar = picChart.ScaleWidth / MaxData
End Function

Private Function GetY(Value) As Double
    'Returns the position on the chart for that data
    On Error Resume Next
    GetY = MaxValue - Value
End Function

Private Function GetX(Value) As Double
    'Returns the X position
    GetX = (picChart.ScaleWidth / MaxData) * Value
    GetX = GetX - (MaxData * 0.01) - 1
End Function

Private Function DateLookup(X)
    'Lookup a date that has been stored in aryDate

    Dim lngLoop As Long
    Dim dblLast As Double
    Dim dblFirst As Double
    Dim dblMid As Double

    On Error Resume Next
    NewLookup = False

    For lngLoop = 1 To UBound(aryDate) - 1
        dblFirst = aryDate(lngLoop, 2)
        dblLast = aryDate(lngLoop + 1, 2)
        dblMid = dblLast - dblFirst
        dblMid = dblMid / 2
        If X <= dblLast - dblMid And X >= dblFirst - dblMid Then Exit For
        DoEvents
        If NewLookup = True Then Exit For
    Next lngLoop

    If lngLoop > UBound(aryDate) Then Exit Function
    DateLookup = aryDate(lngLoop, 1)

End Function

Public Function About()
    frmAbout.Show
End Function

Public Function DrawLineChart(LineValue, Optional ChartLineColor As Long = -1)

    '  On Error GoTo ErrHndl

    Dim lngLoop As Long
    Dim intCount As Integer
    Dim VarOldVal

    SavedLineColor = ChartLineColor

    VarLineData = LineValue

    'If its being repainted then don't add it to the collection of arrays
    If Not Repainting Then
        'Make space for the new line chart
        ReDim Preserve LineChartsArray(1 To UBound(LineChartsArray) + 1)
        ReDim Preserve ArrayLineColor(1 To UBound(ArrayLineColor) + 1)
        'Save this line chart
        ArrayLineColor(UBound(ArrayLineColor) - 1) = ChartLineColor
        LineChartsArray(UBound(LineChartsArray) - 1) = LineValue
    End If

    Repainting = False

    For lngLoop = 1 To MaxData Step StepVar

        intCount = intCount + 1

        'See if there is any more data to be displayed
        If intCount > MaxData Then Exit For

        'Paint the line
        If VarOldVal = 0 Then VarOldVal = GetY(LineValue(intCount))

        picChart.Line (GetX(intCount), GetY(LineValue(intCount)))-(GetX(intCount) - GapVal * 3, _
                VarOldVal), ChartLineColor
        VarOldVal = GetY(LineValue(intCount))
    Next lngLoop

    Exit Function
ErrHndl:
    Err.Raise "BarChartX: " & vbCrLf & Err.Description, vbExclamation
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picChart,picChart,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = picChart.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    picChart.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'Load property values
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picChart.BorderStyle() = PropBag.ReadProperty("BorderStyle", 0)
    picChart.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
End Sub

Private Sub UserControl_Terminate()
    tmrLookup.Enabled = False
End Sub

'Write property values
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", picChart.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", picChart.BackColor, &HFFFFFF)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picChart,picChart,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picChart.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picChart.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

