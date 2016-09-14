VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Click Me repeatedly!"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'VB has no problems, to serialize/deserialize the nested tTest-UDT below...
'(even when those UDT-Defs contain members with a dynamic length, as Strings or Arrays)
'It can do that over its Put and Get functions, but these are restricted to
'Files only... well, this example shows how to use a named Pipes FileName to
'work around that, to keep VBs nice serialization/deserialization-feature InMemory...

Private Type tTestSubType
  S As String
  b() As Byte
  v As Variant
End Type

Private Type tTest
  L As Long
  S As String
  b() As Byte
  SubType As tTestSubType
End Type

Private struc1 As tTest
Private struc2 As tTest

Private pipe As New cPipedUDTs
Attribute pipe.VB_VarHelpID = -1

Private Sub Form_Load()
   pipe.Init "vbpipe" 'init with a Pipe-Suffix of your own choice
End Sub

Private Sub Form_Click()
       
    If pipe.handle = 0 Then
        Print "Pipe library is not initilized? handle = 0"
        Exit Sub
    End If
    
    Dim b() As Byte
    Static cc As Long: cc = cc + 1 'increment a static Counter-Value
 
    'fill in some Demo-Values into struc1 (it will be deserialized later into struc2)
    With struc1
      .L = cc
      .S = "String-" & cc
      .b = "ByteArray-Content " & cc
      .SubType.S = "SubType-String-" & cc
      .SubType.b = "SubType-ByteArray-Content " & cc
      .SubType.v = Array(1234, "blah!", 3.14)
    End With
  
    'first the serialization of struc1 into B()
    Put pipe.handle, 1, struc1 'you can serialize any (arbitrary) UDT here
    b() = pipe.ReadBytes
 
    'we can put B() into a resource or a DB-Blob-Field or transfer it over Sockets
    'now the deserializaion of B() back into the same Type of Struct - here into struc2
    pipe.WriteBytes b
    Get pipe.handle, 1, struc2
    
    With struc2 'print out, what we got from our serializations of struc1 -> over B() -> into struc2
        Print .L, .S, .b, .SubType.S, .SubType.b, Join(.SubType.v, ", ")
    End With
    
End Sub
 
 
