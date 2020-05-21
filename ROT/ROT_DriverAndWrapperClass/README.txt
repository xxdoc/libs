'---------------------
' DIRECTIONS AND NOTES
'---------------------

* Put all 4 files in the WindOWS\sYSTEM OR WINNT\System32 folder.

* Register the .tlb's with either Regsvr32.exe, or use the 
  included RegisterDriver.exe program (VB6 exe).

* The "RunningObjectTable.cls" class contains Public and Private procedures. 
  The Public ones are currently functional.  I have unresolved bugs in the 
  Private ones, which is why I marked them Private for now.

  This class is a wrapper class for Edanmo's OLE interfaces and functions v1.81 (olelib.tlb).
  It can be downloaded from: http://www.mvps.org/emorcillo/en/code/vb6/index.shtml if you 
  also want the code used to build the type library, plus the type libraries, too.

  This class will require References in VB6 (or Access?) being set to Word, PowerPoint, Excel,
  and Access; or, if you prefer using CreateObject("Access.Application"), CreateObject("Word.Application"),
  CreateObject("PowerPoint.Application"), and CreateObject("Excel.Application"). 

  The class will also require a Reference being set to "olelib.tlb".


* A quick example regarding 3 procedures in the wrapper class:

  GetPowerPointApp()
  WordObjectsInROT()
  ExcelObjectsInROT()

  Here's sample code using these 3 procedures:

' Declarations
	Dim app as Excel.Application 
	Dim wb as Excel.Workbook
	' or:  Dim app as Object: Set app = CreateObject("Excel.Application")
	' or:  Dim wb as Object: Set wb = CreateObject("Excel.Workbook")
	Dim arr
	Dim x as long 
	Dim rot as RunningObjectTable: Set rot = New RunningObjectTable
 
' Fill Array with Object Table
	arr = rot.ExcelObjectsInROT

	' arr(0, x) : Full path and filename to ALL open Excel workbooks (aka programs) that are open
	' arr(1, x) : Excel *Workbook* object.  Don't try setting an Excel.Application object to this, 
	'             or a Worksheet object. Will just get type-mismatch error.

' Get Excel Workbook object from array to be worked with; 
' Automation of Excel (or Word and PowerPoint) can be done
' with Workbook object from here
	Set wb = xl(1, x)

' Make Workbook Visible
	wb.Visible = True 

'-------------------------------
' Rest of your code goes here...
'-------------------------------


