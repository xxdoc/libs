VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunningObjectTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------------------
' Last Updated: 6/15/06
'----------------------------------------------------------------------------------------------------------------
' Notes:
'----------------------------------------------------------------------------------------------------------------

'You will need a type library with the ROT interfaces (the OLELIB.TLB that's
'in my site have those interfaces). Once you have the type library, you can
'use the following code to enumerate all Word.Application objects and save
'the Documents:
Public Function EnumROT()

' Just a General testing function. It displays various
' data from Excel, PowerPoint, and Word items in the ROT.

Dim oROT As IRunningObjectTable
Dim oEnmMk As IEnumMoniker
Dim oMK As IMoniker
Dim oBC As IBindCtx
Dim lPtr As Long
Dim sName As String
Dim oApplication As Word.Application
Dim oDocument As Word.Document
Dim oPPApp As PowerPoint.Application
Dim oPPPres As PowerPoint.Presentation
Dim oXLApp As Excel.Application
Dim oXLWb As Excel.Workbook
Dim oXLWs As Excel.Worksheet
Dim oXLWn   As Excel.Window
Dim obj As Object
Dim app As access.Application

   ' Get the ROT object
   Set oROT = GetRunningObjectTable

   ' Create a bind context object
   Set oBC = CreateBindCtx

   ' Get the ROT enumerator
   Set oEnmMk = oROT.EnumRunning

    Dim ws As Workspace
    Dim db As Database


   ' Enumeate all registered objects
   Do While oEnmMk.Next(1, oMK) = 0
        
      ' Get the object name
      lPtr = oMK.GetDisplayName(oBC, Nothing)
      sName = SysAllocString(lPtr)
      CoTaskMemFree lPtr
        
        'Debug.Print sName
        
' Start Test -------------------------------------------------------------------------------
        ' Set obj = oROT.GetObject(oMK)
        ' If TypeOf obj Is Excel.Application Then Debug.Print "Excel: " & sName
        ' If TypeOf obj Is Word.Application Then Debug.Print "Word: " & sName
        ' If TypeOf obj Is powerpoint.Application Then Debug.Print "PowerPoint: " & sName
' End Test ---------------------------------------------------------------------------------
        
      If sName = "!{73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}" Then ' Access 2000/2002/2003/XP
            Set app = oROT.GetObject(oMK)
            Debug.Print app.Name
            Debug.Print app.DBEngine.Workspaces(0).Name
            
            If TypeOf obj Is access.Application Then
                Set app = obj.Application
                Debug.Print app.DBEngine.Workspaces(0).Name
            End If
            
            'Stop
            
            For Each db In app.DefaultWorkspaceClone.Databases
                Debug.Print db.Name
            Next
            'Stop
            For Each ws In app.DBEngine.Workspaces
                Debug.Print app.DBEngine.Workspaces.Count
                For Each db In ws.Databases
                    Debug.Print db.Connection
                Next
                'Debug.Print ws.Databases(0).Name
                'Debug.Print ws.Name
            Next
      End If

      ' !{000209FF-0000-0000-C000-000000000046} is
      ' the Word.Application object
      If sName = "!{000209FF-0000-0000-C000-000000000046}" Then

         ' Get the object instance
         Set oApplication = oROT.GetObject(oMK)

         ' Save each document
         For Each oDocument In oApplication.Documents
            Debug.Print oDocument.Name
            'oDocument.Save
         Next

      End If

      ' the PowerPoint.Application object
'      If sName = "!{00024500-0000-0000-C000-000000000046}" Then
        If sName = "!{91493441-5A91-11CF-8700-00AA0060263B}" Then

         ' Get the object instance
         Set oPPApp = oROT.GetObject(oMK) ' same as "oPPApp = ObjPtr(instance of PowerPoint)" ?

         ' Save each document
         Dim sld As PowerPoint.Slide
         For Each oPPPres In oPPApp.Presentations
            Debug.Print oPPPres.Name
            Debug.Print "Slide Count: " & oPPPres.Slides.Count
            For Each sld In oPPPres.Slides
                Debug.Print sld.Name
            Next sld
            Debug.Print VBA.String$(25, "-")
            'oPPPres.Save
         Next

      End If

      '{00020812-0000-0000-C000-000000000046}
      ' the Excel.Application object
      If sName = "!{00024500-0000-0000-C000-000000000046}" Then
        
         ' Get the object instance
         Set obj = oROT.GetObject(oMK)
         'Set oXLApp = oROT.GetObject(oMK)
         
         ' Test Object
         If (TypeOf obj Is Excel.Workbook) Then
            Set oXLApp = obj.Application
         End If
            
         ' Save each document
         For Each oXLWb In obj.Workbooks
            'oXLWb.Save
            Debug.Print oXLWb.Name
            'For Each oXLWs In oXLWb.Worksheets
            '    Debug.Print "   " & oXLWs.Name
            'Next
         Next

      End If
   
   Loop

   Set oROT = Nothing

End Function

Private Function GetAccessApps() As Variant
' Does not appear to work like Excel, Word, and Powerpoint with Edanmo's control.
' Will take more research and testing to fix.
    ' Error Handling
        On Error Resume Next
    
    ' Declarations
        Dim oROT As IRunningObjectTable, oEnmMk As IEnumMoniker, oMK As IMoniker
        Dim obj As Object, oAccess As access.Application
        Dim arr()
        Dim ndx As Long

    ' Initialzation
        Set oROT = GetRunningObjectTable    ' Get the ROT object
        Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator
        ReDim Preserve arr(1, 0)
        arr(0, 0) = 0
        ndx = 0
    
    ' Enumerate all registered objects
        Do While oEnmMk.Next(1, oMK) = 0
            Set obj = oROT.GetObject(oMK)   ' Get the object instance
        
            If Err Then
                Err.Clear
            Else
                If TypeOf obj Is access.Application Then
                '    Set oAccess = obj.Application
                '    ' add to collection; (ObjPtr only generates unique key)
                '    'GetAccessApps.Add oAccess, "k" & ObjPtr(oAccess)   ' <= trick to the whole thing is adding each
                '    '                                                   '    Access application object into Collection
                    ReDim Preserve arr(1, ndx)
                    arr(0, ndx) = obj.CurrentProject.Name
                    Set arr(1, ndx) = obj
                    ndx = ndx + 1
                    'Debug.Print obj.CurrentProject.Name
                    If Err Then Err.Clear
                End If
            End If
        Loop
    
    ' Return Array
        GetAccessApps = arr
    
    ' Erase Objects from Memory
        Set oROT = Nothing
        Set oEnmMk = Nothing
        Set oMK = Nothing
        Set obj = Nothing
        Set oAccess = Nothing
        Erase arr
End Function

Private Function GetWordApps() As Collection
    Dim oROT As IRunningObjectTable, oEnmMk As IEnumMoniker, oMK As IMoniker
    Dim obj As Object, oWord As Word.Application
    On Error Resume Next

    Set GetWordApps = New Collection

    Set oROT = GetRunningObjectTable    ' Get the ROT object
    Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator
    Do While oEnmMk.Next(1, oMK) = 0    ' Enumeate all registered objects
        Set obj = oROT.GetObject(oMK)   ' Get the object instance
        If Err Then
            Err.Clear
        Else
            If TypeOf obj Is Word.Document Then
                Set oWord = obj.Application
                ' add to collection; (ObjPtr only generates unique key)
                GetWordApps.Add oWord, "k" & ObjPtr(oWord)   ' <= trick to the whole thing is adding each
                                                               '    Word application object into Collection
                If Err Then Err.Clear
            End If
        End If
    Loop
End Function

Public Function GetPowerPointApp() As Variant
'-----------------------------------------------------------------------------------------------------------
' Purpose: Return 2D array containing name and application object of each running PowerPoint instance.
'-----------------------------------------------------------------------------------------------------------
    
    ' Error Handling
        On Error Resume Next
        
    ' Declarations
        Dim oROT As IRunningObjectTable
        Dim oEnmMk As IEnumMoniker
        Dim oMK As IMoniker
        Dim oBC As IBindCtx
        Dim obj As Object
        Dim oPowerPoint As PowerPoint.Application
        Dim oPres As PowerPoint.Presentation
        Dim lPtr As Long
        Dim sName As String
        Dim sPresName As String
        Dim ndx As Long
        Dim arr()

    ' Initialization
        Set oROT = GetRunningObjectTable    ' Get the ROT object
        Set oBC = CreateBindCtx             ' Create a bind context object
        Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator
        ReDim Preserve arr(1, 0)
        arr(0, 0) = 0
        ndx = 0

    ' Enumerate all registered objects
        Do While oEnmMk.Next(1, oMK) = 0
        
            ' Get the object name
                lPtr = oMK.GetDisplayName(oBC, Nothing)
                sName = SysAllocString(lPtr)
                CoTaskMemFree lPtr
        
            If sName = "!{91493441-5A91-11CF-8700-00AA0060263B}" Then
                ' Get PowerPoint Application Object
                    Set oPowerPoint = oROT.GetObject(oMK)
                ' Get Presentation Objects and Names from Application Object and Insert into Array
                    For Each oPres In oPowerPoint.Presentations
                        ReDim Preserve arr(1, ndx)
                        arr(0, ndx) = oPres.FullName
                        Set arr(1, ndx) = oPres
                        ndx = ndx + 1
                    Next
            End If
        Loop

    ' Return Array
        GetPowerPointApp = arr
End Function

Private Function GetPowerPointApps() As Collection
' This function does not work correctly. I created GetPowerPointApp (without the 's') AFTER this one!
' Problem is that this code, which is taken from the GetExcelApps() procedure, echoes all running
' PowerPoint instances twice in the Collection.  Not sure why. The only way around it is to use the
' code from EnumROT. I've since replace it with GetPointPointApp() and it works!

    Dim oROT As IRunningObjectTable, oEnmMk As IEnumMoniker, oMK As IMoniker, oBC As IBindCtx
    Dim obj As Object, oPowerPoint As PowerPoint.Application, app As PowerPoint.Application
    Dim lPtr As Long, sName As String
    Dim flag As Boolean
    On Error Resume Next

    Set GetPowerPointApps = New Collection

    Set oROT = GetRunningObjectTable    ' Get the ROT object
    Set oBC = CreateBindCtx             ' Create a bind context object
    Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator

    Do While oEnmMk.Next(1, oMK) = 0    ' Enumerate all registered objects
        
        ' Get the object name
        lPtr = oMK.GetDisplayName(oBC, Nothing)
        sName = SysAllocString(lPtr)
        CoTaskMemFree lPtr
        
        'Debug.Print sName
        
'GoTo Skip:

        If sName = "!{0002DF01-0000-0000-C000-000000000046}" Then
            Debug.Print sName
        End If
            
            Set obj = oROT.GetObject(oMK)   ' Get the object instance
            If Err Then
                Err.Clear
            Else
                If TypeOf obj Is PowerPoint.Presentation Then
                    Set oPowerPoint = obj.Application
                    ' only add first instance of same item
                        flag = False
                        For Each app In GetPowerPointApps
                            If app = oPowerPoint Then
                                ' add to collection; (ObjPtr only generates unique key)
                                    GetPowerPointApps.Add oPowerPoint, "k" & ObjPtr(oPowerPoint) ' <= trick to the whole thing is adding each
                            End If
                        Next
                    If Err Then Err.Clear
                End If
            End If
        'End If
    Loop

    Debug.Print "(" & GetPowerPointApps.Count & ")"
End Function

Public Sub EnumPowerPointCollection()
    
    Dim oPowerPoint As PowerPoint.Application, oPres As PowerPoint.Presentation, oSlide As PowerPoint.Slide
    Dim col As Collection, S As String
    
    Set col = GetPowerPointApps() ' get collection of PowerPoint applications
    
    S = col.Count & " PowerPoint Applications:" & vbCr
    For Each oPowerPoint In col ' just references items in collection using object pointers that were put in collection
        S = S & vbCr & " - " & oPowerPoint.Presentations.Count & " Presentations:" & vbCr
        For Each oPres In oPowerPoint.Presentations
            S = S & "    " & oPres.Name & " has " & oPres.Slides.Count & " slides:"
            For Each oSlide In oPres.Slides
                S = S & " " & oSlide.Name
            Next
            S = S & vbCr
        Next
    Next
    MsgBox S

End Sub


Private Function GetExcelApp() As Variant
'-------------------------------------------------------------------------------------------------------------
' Notes:
'   * I created this procedure AFTER GetExcelApps() [no 's' on this one's name]
'-------------------------------------------------------------------------------------------------------------
    ' Error Handling
        On Error Resume Next
        
    ' Declarations
        Dim oROT As IRunningObjectTable
        Dim oEnmMk As IEnumMoniker
        Dim oMK As IMoniker
        Dim obj As Object
        Dim oExcel As Excel.Application
        Dim arr()
        Dim ndx As Long
        Dim up As UtilityPathInternal: Set up = New UtilityPathInternal

    ' Initialization
        Set oROT = GetRunningObjectTable    ' Get the ROT object
        Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator
        ReDim arr(1, 0)
        arr(0, 0) = 0
        ndx = 0

    ' Enumerate all registered objects
        Do While oEnmMk.Next(1, oMK) = 0
            Set obj = oROT.GetObject(oMK)   ' Get the object instance
            If Err Then
                Err.Clear
            Else
                If TypeOf obj Is Excel.Workbook Then
                    Set oExcel = obj.Application
                    ' add to collection; (ObjPtr only generates unique key)
                    GetExcelApps.Add oExcel, "k" & ObjPtr(oExcel)   ' <= trick to the whole thing is adding each
                                                                    '    Excel application object into Collection
                    arr(0, ndx) = up.BuildPath(oExcel.Path, oExcel.Name)
                    
                    If Err Then Err.Clear
                End If
            End If
        Loop
        
    ' Return Array
        GetExcelApp = arr
        
    ' Erase Objects from Memory
        Set up = Nothing
        Set oROT = Nothing
        Set oMK = Nothing
End Function

' returns collection of Excel applications
Private Function GetExcelApps() As Collection
    Dim oROT As IRunningObjectTable, oEnmMk As IEnumMoniker, oMK As IMoniker
    Dim obj As Object, oExcel As Excel.Application
    On Error Resume Next

    Set GetExcelApps = New Collection

    Set oROT = GetRunningObjectTable    ' Get the ROT object
    Set oEnmMk = oROT.EnumRunning       ' Get the ROT enumerator
    Do While oEnmMk.Next(1, oMK) = 0    ' Enumeate all registered objects
        Set obj = oROT.GetObject(oMK)   ' Get the object instance
        If Err Then
            Err.Clear
        Else
            If TypeOf obj Is Excel.Workbook Then
                Set oExcel = obj.Application
                ' add to collection; (ObjPtr only generates unique key)
                GetExcelApps.Add oExcel, "k" & ObjPtr(oExcel)   ' <= trick to the whole thing is adding each
                                                               '    Excel application object into Collection
                If Err Then Err.Clear
            End If
        End If
    Loop
End Function

Public Sub EnumExcelCollection()
    
    Dim oExcel As Excel.Application, oWb As Excel.Workbook, oSht As Excel.Worksheet
    Dim col As Collection, S As String
    
    Set col = GetExcelApps() ' get collection of Excel applications
    
    S = col.Count & " Excel Applications:" & vbCr
    For Each oExcel In col ' just references items in collection using object pointers that were put in collection
        S = S & vbCr & " - " & oExcel.Workbooks.Count & " Workbooks:" & vbCr
        For Each oWb In oExcel.Workbooks
            S = S & "    " & oWb.Name & " has " & oWb.Worksheets.Count & " worksheets:"
            For Each oSht In oWb.Worksheets
                S = S & " " & oSht.Name
            Next
            S = S & vbCr
        Next
    Next
    MsgBox S

End Sub

Public Function ExcelObjectsInROT() As Variant
'--------------------------------------------------------------------------------------------
' Returns: 2D array containing Excel Application objects with their corresponding names.
'          Dimension 1 is the Name. Dimension 2 is the Object (Word, Excel, or
'          PowerPoint Application object)
'--------------------------------------------------------------------------------------------

    ' Declarations
        Dim col As Collection
        Dim obj As Object
        Dim wb As Excel.Workbook
        Dim app As Excel.Application
        Dim arr()
        Const C_NAME As Long = 0
        Const C_OBJECT As Long = 1
        Dim cnt As Long
        Dim decr As Long
        Dim last_element As Long
        Dim start_index As Long

    ' Initialize
        'Set obj = CreateObject("Excel.Application")
        ReDim arr(1, 0)
        arr(0, 0) = 0
        cnt = 0
        decr = 0
        last_element = 0
        start_index = 0

    ' Get collection of Excel applications
        Set col = GetExcelApps()

    ' Enum Collection: Put Excel.Application object, along with Workbook name, into array
        For Each app In col
            ' Initialize Array's Indexes, Elements, and Size
                If arr(0, 0) = 0 Then  ' First Pass
                    ' Set Last Index Based on Workbook Count
                        If app.Workbooks.Count > 1 Then
                            last_element = app.Workbooks.Count - 1 ' use element 0 for position 1
                        Else
                            last_element = 0
                        End If
                    ' Grow Array (if needed)
                        ReDim Preserve arr(1, last_element)
                    ' Set Starting Index
                        start_index = 0
                Else ' All Passes After First Pass
                    ' Set Last Index Based on Last Element in Array and Workbook Count
                        last_element = app.Workbooks.Count + UBound(arr, 2)
                    ' Set Starting Index (must be done before growing array)
                        start_index = UBound(arr, 2) + 1
                    ' Grow Array
                        ReDim Preserve arr(1, last_element)
                End If
            ' Populate New Array Elements with Each Workbook Name in the Collection
            ' and with Each Workbook's Application Object (kinda wierd, but it works)
                cnt = start_index
                'Debug.Print "cnt=" & cnt
                For Each wb In app.Workbooks
                    'Debug.Print "  --->" & wb.Name
                    arr(C_NAME, cnt) = wb.FullName
                    Set arr(C_OBJECT, cnt) = wb
                    cnt = cnt + 1
                Next
        Next

    ' Return Array
        ExcelObjectsInROT = arr

    ' Erase Objects from Memory
        Set obj = Nothing
End Function

Public Sub TestExcelObjectsInROT()
'-------------------------------------
' Example Procedure - Fully functional
'-------------------------------------

    ' Declarations
        Dim rot As New RunningObjectTable
        Dim app As Excel.Application
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim arr()
        Dim x As Long

    ' Populate Array with Names and Objects
        arr = rot.ExcelObjectsInROT

    ' Display Running Excel Objects and Their Worksheets
        If arr(0, 0) <> "0" Then
            Debug.Print vbNewLine & "MaxIndex=" & UBound(arr, 2) & vbNewLine
            For x = 0 To UBound(arr, 2)
                Debug.Print "Spreadsheet(" & x & "): " & arr(0, x)
                Set app = arr(1, x)
                ' Display Workbooks
                    'For Each wb In app.Workbooks
                    '    Debug.Print VBA.String$(20, " ") & "Workbook: " & wb.Name
                    'Next wb
                ' Display Worksheets
                    For Each ws In app.Worksheets
                        Debug.Print VBA.String$(20, " ") & "Worksheet: " & ws.Name
                    Next ws
                ' Skip a Line
                    Debug.Print
            Next x
        Else
            Debug.Print "*** No Excel Spreadsheets are open"
        End If

    ' Erase Objects and Arrays from Memory
        Set rot = Nothing
        Set app = Nothing
        Set wb = Nothing
        Set ws = Nothing
        Erase arr
End Sub

Private Function PowerPointObjectsInROT() As Variant
'--------------------------------------------------------------------------------------------
' Returns: 2D array containing PowerPoint Application objects with their corresponding names.
'          Dimension 1:Element 0 is the Name. Dimension 1:Element 1 is the Object
'          (Word, Excel, or PowerPoint Application object)
' Problem: Returns multiple instances; for example, if I have open 2 PowerPoint files, then
'          4 titles get returned, rather then 2. Use GetPointPointApp() to get a 2D array of
'          *only* one set of filenames and the corresponding Presentation object for each.
'--------------------------------------------------------------------------------------------

    ' Declarations
        Dim col As Collection
        Dim obj As Object
        Dim pres As PowerPoint.Presentation
        Dim app As PowerPoint.Application
        Dim arr()
        Const C_NAME As Long = 0
        Const C_OBJECT As Long = 1
        Dim cnt As Long
        Dim decr As Long
        Dim last_element As Long
        Dim start_index As Long

    ' Initialize
        'Set obj = CreateObject("PowerPoint.Application")
        ReDim arr(1, 0)
        arr(0, 0) = 0
        cnt = 0
        decr = 0
        last_element = 0
        start_index = 0

    ' Get collection of PowerPoint applications
        Set col = GetPowerPointApps()

    ' Enum Collection: Put PowerPoint.Application object, along with
    '                  Presentation name, into array
        For Each app In col
            ' Initialize Array's Indexes, Elements, and Size
                If arr(0, 0) = 0 Then  ' First Pass
                    ' Set Last Index Based on Workbook Count
                        If app.Presentations.Count > 1 Then
                            last_element = app.Presentations.Count - 1 ' use element 0 for position 1
                        Else
                            last_element = 0
                        End If
                    ' Grow Array (if needed)
                        ReDim Preserve arr(1, last_element)
                    ' Set Starting Index
                        start_index = 0
                Else ' All Passes After First Pass
                    ' Set Last Index Based on Last Element in Array and Workbook Count
                        last_element = app.Presentations.Count + UBound(arr, 2)
                    ' Set Starting Index (must be done before growing array)
                        start_index = UBound(arr, 2) + 1
                    ' Grow Array
                        ReDim Preserve arr(1, last_element)
                End If
            ' Populate New Array Elements with Each Workbook Name in the Collection
            ' and with Each Workbook's Application Object (kinda wierd, but it works)
                cnt = start_index
                'Debug.Print "cnt=" & cnt
                For Each pres In app.Presentations
                    Debug.Print "  --->" & pres.Name
                    arr(C_NAME, cnt) = pres.Name
                    Set arr(C_OBJECT, cnt) = app ' set variant array element to object to get object in array; won't populate if just set equal to (=)
                    cnt = cnt + 1
                Next
        Next

    ' Return Array
        PowerPointObjectsInROT = arr

    ' Erase Objects from Memory
        Set obj = Nothing
End Function

Public Sub TestPowerPointObjectsInROT()
'-------------------------------------
' Example Procedure - Fully functional
'-------------------------------------

    ' Declarations
        Dim rot As New RunningObjectTable
        Dim app As PowerPoint.Application
        Dim pres As PowerPoint.Presentation
        Dim sld As PowerPoint.Slide
        Dim arr()
        Dim x As Long

    ' Populate Array with Names and Objects
        arr = PowerPointObjectsInROT

    ' Display Running PowerPoint Objects and Their Worksheets
        If arr(0, 0) <> 0 Then
            Debug.Print vbNewLine & "MaxIndex=" & UBound(arr, 2) & vbNewLine
            For x = 0 To UBound(arr, 2)
                Debug.Print "Presentation(" & x & "): " & arr(0, x)
                Set app = arr(1, x)
                ' Display Presentations
                    'For Each pres In app.Presentations
                    '    Debug.Print VBA.String$(20, " ") & "Presentation: " & pres.Name
                        ' Display Slides
                            For Each sld In app.Presentations(arr(0, x)).Slides
                                Debug.Print VBA.String$(15, " ") & "Slide: " & sld.Name
                            Next sld
                    'Next pres
                ' Skip a Line
                    Debug.Print
            Next x
        Else
            Debug.Print "*** No PowerPoint Presentations are open"
        End If

    ' Erase Objects and Arrays from Memory
        Set rot = Nothing
        Set app = Nothing
        Set pres = Nothing
        Set sld = Nothing
        Erase arr
End Sub

Public Function WordObjectsInROT() As Variant
'--------------------------------------------------------------------------------------------
' Returns: 2D array containing Word Application objects with their corresponding names.
'          Dimension 1:Element 0 is the Name. Dimension 1:Element 1 is the Object
'          (Word, Excel, or PowerPoint Application object)
'--------------------------------------------------------------------------------------------

    ' Declarations
        Dim col As Collection
        Dim obj As Object
        Dim doc As Word.Document
        Dim app As Word.Application
        Dim arr()
        Const C_NAME As Long = 0
        Const C_OBJECT As Long = 1
        Dim cnt As Long
        Dim decr As Long
        Dim last_element As Long
        Dim start_index As Long

    ' Initialize
        'Set obj = CreateObject("Word.Application")
        ReDim arr(1, 0)
        arr(0, 0) = 0
        cnt = 0
        decr = 0
        last_element = 0
        start_index = 0
        
    ' Get collection of PowerPoint applications
        Set col = GetWordApps()

    ' Enum Collection: Put PowerPoint.Application object, along with
    '                  Presentation name, into array
        For Each app In col
            ' Initialize Array's Indexes, Elements, and Size
                If arr(0, 0) = 0 Then  ' First Pass
                    ' Set Last Index Based on Workbook Count
                        If app.Documents.Count > 1 Then
                            last_element = app.Documents.Count - 1 ' use element 0 for position 1
                        Else
                            last_element = 0
                        End If
                    ' Grow Array (if needed)
                        ReDim Preserve arr(1, last_element)
                    ' Set Starting Index
                        start_index = 0
                Else ' All Passes After First Pass
                    ' Set Last Index Based on Last Element in Array and Workbook Count
                        last_element = app.Documents.Count + UBound(arr, 2)
                    ' Set Starting Index (must be done before growing array)
                        start_index = UBound(arr, 2) + 1
                    ' Grow Array
                        ReDim Preserve arr(1, last_element)
                End If
            ' Populate New Array Elements with Each Workbook Name in the Collection
            ' and with Each Workbook's Application Object (kinda wierd, but it works)
                cnt = start_index
                'Debug.Print "cnt=" & cnt
                For Each doc In app.Documents
                    'Debug.Print "  --->" & pres.Name
                    arr(C_NAME, cnt) = doc.FullName
                    Set arr(C_OBJECT, cnt) = doc ' Set variant array element to object to get object in array; won't populate if just set equal to (=) without 'Set'
                    cnt = cnt + 1
                Next
        Next

    ' Return Array
        WordObjectsInROT = arr

    ' Erase Objects from Memory
        Set obj = Nothing
End Function

