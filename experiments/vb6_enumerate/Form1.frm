VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "vb6 enumerate 'operator' test"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i, v, sample, colkey
    Dim c As New Collection
    Dim l As Long
    
    List1.AddItem "Experiment using a user function like a built in language operator"
    
    sample = Array("apple", "orange", "cat", "dog")
    
    List1.AddItem "----- [ walk simple array ] -------"
    Do While enumerate(i, v, sample)
        List1.AddItem "  Index: " & i & " Value: " & v
        c.Add v
        sample(i) = i
    Loop
    
    List1.AddItem "   Array Values now: " & Join(sample, ", ")
    
    List1.AddItem "----- [ walk array with long var enumerator ] -------"
    
    'walk changed array, use long as enum variable (variant not required like for each)
    Do While enumerate(i, l, sample)
        List1.AddItem "  Index: " & i & " Value: " & l
    Loop
    
    List1.AddItem "----- [ enum collection of simple values ] -------"
    
    Do While enumerate(i, v, c)
        List1.AddItem "  Collection Index: " & i & " Value: " & v
    Loop
    
    Dim c1 As Class1 'class source: Public id
    Set c = New Collection
    
    'build up a collection of class objects..
    For i = 0 To 3
        Set c1 = New Class1
        c1.id = i
        c.Add c1, "key:" & i
    Next
    
    'note the i we pass in here will be 4 and not initilized...
    'enumerate will detect the change in object and reset..
    'there can be a case if we were using same collection as last
    'enumerate call and i was not 0 we would be undefined behavior..
    'note we can use a Class1 variable as our enumerator and is not required to be variant
    'enumerator also supports enumerating the items collection key..not possible with for each..
    
    List1.AddItem "----- [ walk obj collection with key and class obj var ] -------"
    List1.AddItem "----- [ note i = " & i & " is unitilized  ] -------"
    Do While enumerate(i, c1, c, colkey)
        List1.AddItem "  Objs Index: " & i & " Class1 id: " & c1.id & " ColKey: " & colkey
    Loop

    sample = Array("apple", "orange", "cat", "dog")
    
    List1.AddItem "----- [ bug test 1 - incomplete loop exit + same obj enum ] -------"
    
    Do While enumerate(i, v, sample)
        List1.AddItem "  Index: " & i & " Value: " & v
        If i = 1 Then Exit Do
    Loop
    
    List1.AddItem "----- [ first enum call exited, enum again same i same obj ] -------"
    
    Do While enumerate(i, v, sample)
        List1.AddItem "  Index: " & i & " Value: " & v
    Loop
    
    List1.AddItem "----- [ ok maybe that is actually a feature..could be useful ] -------"
    
'    List1.AddItem "----- [ bug test 2 - nesting = endless loop :( ] -------"
'    'and this will not work and do weird stuff..
'    Do While enumerate(i, v, sample)
'
'        List1.AddItem "  Array Index: " & i & " Value: " & v
'
'        Do While enumerate(i, c1, c, colkey)
'            List1.AddItem "  Objs Index: " & i & " Class1 id: " & c1.id & " ColKey: " & colkey
'        Loop
'
'    Loop
    
    
    
End Sub

'limitation 1: you can not nest calls to enumerate because of static counter

'bug: we try to detect if this is a new call to enumerate or part of a loop already established..
    'this is not rock solid, possible bugs can be avoided if user always passes in 0 for start index..
    'we could avoid these bugs with more framework around this..like an init() call or reset() but
    'we want this to mirror a built in language function as close as possible with practically no setup required..
    'so what is the bug? imagine this
    '  you call enumerate(i,x),
    '  you exit the loop early with exit do
    '  you call enumerate again with same i same x, it can not detect this and will continue from last left off..
    '  whatever...maybe thats a feature :P


Function enumerate(ByRef startIndex, ByRef value As Variant, ByRef obj, Optional ByRef key) As Boolean
    Dim t As String
    
    Static counter As Long
    Static lastObj As Long
    
    t = TypeName(obj)
    clearVal value
    key = Empty
    
     
    If IsObject(obj) Then
        If lastObj <> ObjPtr(obj) Then
            startIndex = 0
            counter = 0
            lastObj = ObjPtr(obj)
        End If
    ElseIf IsArray(obj) Then
        If lastObj <> VarPtr(obj) Then
            startIndex = 0
            counter = 0
            lastObj = VarPtr(obj)
        End If
    Else
        Err.Raise "Invalid Source, expects array or collection - given type: " & t, "enumerate()"
    End If
        
    
    If InStr(t, "()") > 0 Then 'array type
        If AryIsEmpty(obj) Then GoTo loopDone
        If counter > UBound(obj) Then GoTo loopDone
        If startIndex < LBound(obj) Then counter = LBound(obj)
    ElseIf t = "Collection" Then
        If obj.Count = 0 Then GoTo loopDone
        If counter > obj.Count Then GoTo loopDone
        If startIndex < 1 Then counter = 1
        key = keyForIndex(counter, obj)
    Else
        Err.Raise "Invalid Source, expects array or collection - given type: " & t, "enumerate()"
    End If
        
        
    If IsObject(obj(counter)) Then
        Set value = obj(counter)
    Else
        value = obj(counter)
    End If
    
    startIndex = counter  'current item index
    counter = counter + 1 'advance to next one (limitation: nesting of enumerate calls not supported)
    enumerate = True
    
Exit Function

loopDone:
    counter = 0
    startIndex = -1
    lastObj = 0
    enumerate = False
    
End Function


