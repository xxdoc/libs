VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "vb6 enumerate 'operator' test"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBenchMark 
      Caption         =   "benchmark"
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCombo 
      Caption         =   "combo test"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8760
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "string test 2"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "string test 1"
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "textbox test 2"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "textbox test 1"
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listbox test"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
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
Private Sub cmdBenchMark_Click()
    
    'in ide 2.2ghz machine, under 10k elements dont really notice a difference..
'    For each 10000000 = 0.641 seconds
'    For i to 10000000 = 0.547 seconds
'    enumerator 10000000 = 22.296 seconds - 41x slower

'    For each 100000 = 0.016 seconds
'    For i to 100000 = 0.015 seconds
'    enumerator 100000 = 0.297 seconds - 20x slower

    'For each 100000 = 0.015 seconds
    'For i to 100000 = 0.016 seconds
    'enumerate_v1 100000 = 0.25 seconds - 17x slower
    
    'For each 10000000 = 0.563 seconds
    'For i to 10000000 = 0.515 seconds
    'enumerate_v1 10000000 = 16.532 seconds - 32x slower
    
    Dim b() As Byte
    Dim report() As String
    Dim x, i
    Const size As Long = 10000
    
    b() = StrConv(String(size, "a"), vbFromUnicode)
    
    StartBenchMark
    For Each x In b
        'x = x + 1
    Next
    push report, "For each " & size & " = " & EndBenchMark()
    
    StartBenchMark
    For i = 0 To UBound(b)
        'b(i) = b(i) + 1
    Next
    push report, "For i to " & size & " = " & EndBenchMark()
    
    StartBenchMark
    Do While enumerate_v1(b, x, i)
        'x = x + 1
    Loop
    push report, "enumerate_v1 " & size & " = " & EndBenchMark()
    
    Debug.Print Join(report, vbCrLf)
    MsgBox Join(report, vbCrLf)
    
End Sub

Private Sub cmdCombo_Click()
    
    Combo1.AddItem "item 0"
    Combo1.AddItem "item 1"
    Combo1.AddItem "item 2"
    Combo1.AddItem "item 3"
    
    Dim i, v, tmp
    While enumerate(Combo1, v, i)
        tmp = tmp & i & "=" & v & vbCrLf
    Wend
    
    MsgBox tmp
    
End Sub

Private Sub Command1_Click()
    'listbox test
    List2.AddItem "a"
    List2.AddItem "bb"
    List2.AddItem "ccc"
    
    Dim i, v, tmp
    While enumerate(List2, v, i)
        tmp = tmp & i & "=" & v & vbCrLf
    Wend
    
    MsgBox tmp
    
End Sub

'textbox test either enum by line(no key) or split at key
Private Sub Command2_Click(index As Integer)
    
    Dim key As String, i, v, tmp
    
    Text1 = Replace("1,2\n3,4\n5,6", "\n", vbCrLf)
    
    If index = 1 Then key = ","
    
    While enumerate(Text1, v, i, key)
        tmp = tmp & i & "=" & Replace(v, vbCrLf, "\n") & vbCrLf
    Wend
    
    MsgBox tmp
    
End Sub

'string test either as byte array(no key) or split at key
'we also test the key change but obj stay the same after a partial loop logic..confusing  but needs testing
Private Sub Command3_Click(index As Integer)

    Dim key As String, i, v, tmp
    Dim s As String
    
    s = "test,test"
    Text1 = s
    
    If index = 1 Then key = ","
    
    Do While enumerate(s, v, i, key)
        tmp = tmp & i & "=" & Replace(v, vbCrLf, "\n") & vbCrLf
        If i = 3 Then Exit Do
    Loop
    
    MsgBox tmp
    tmp = Empty
    
    If index = 0 Then key = "," Else key = Empty
    
    Do While enumerate(s, v, i, key)
        tmp = tmp & i & "=" & Replace(v, vbCrLf, "\n") & vbCrLf
    Loop
    
    MsgBox tmp
    
End Sub

Private Sub Form_Load()
    Dim i, v, sample, colkey
    Dim c As New Collection
    Dim l As Long
    
    List1.AddItem "Experiment using a user function like a built in language operator"
    
    sample = Array("apple", "orange", "cat", "dog")
    
    List1.AddItem "----- [ walk simple array ] -------"
    Do While enumerate(sample, v, i)
        List1.AddItem "  Index: " & i & " Value: " & v
        c.Add v
        sample(i) = i
    Loop
    
    List1.AddItem "   Array Values now: " & Join(sample, ", ")
    
    List1.AddItem "----- [ walk array with long var enumerator ] -------"
    
    'walk changed array, use long as enum variable (variant not required like for each)
    Do While enumerate(sample, l, i)
        List1.AddItem "  Index: " & i & " Value: " & l
    Loop
    
    List1.AddItem "----- [ enum collection of simple values ] -------"
    
    Do While enumerate(c, v, i)
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
    Do While enumerate(c, c1, i, colkey)
        List1.AddItem "  Objs Index: " & i & " Class1 id: " & c1.id & " ColKey: " & colkey
    Loop

    sample = Array("apple", "orange", "cat", "dog")
    
    List1.AddItem "----- [ bug test 1 - incomplete loop exit + same obj enum ] -------"
    
    Do While enumerate(sample, v, i)
        List1.AddItem "  Index: " & i & " Value: " & v
        If i = 1 Then Exit Do
    Loop
    
    List1.AddItem "----- [ first enum call exited, enum again same i same obj ] -------"
    
    Do While enumerate(sample, v, i)
        List1.AddItem "  Index: " & i & " Value: " & v
    Loop
    
    List1.AddItem "----- [ ok maybe that is actually a feature..could be useful ] -------"
    
'    List1.AddItem "----- [ bug test 2 - nesting = endless loop :( ] -------"
'    'and this will not work and do weird stuff..
'    Do While enumerate(i, v, sample)
'
'        List1.AddItem "  Array Index: " & i & " Value: " & v
'
'        Do While enumerate( c,c1, i, colkey)
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


Function enumerate(ByRef obj, ByRef value As Variant, ByRef startIndex, Optional ByRef key) As Boolean
    Dim t As String
    Dim b() As Byte
    
    Static counter As Long
    Static lastObj As Long
    Static internalData
    Static lastKey
    
    t = TypeName(obj)
    clearVal value
    If InStr("TextBox,String", t) < 1 Then key = Empty
    
    If IsObject(obj) Then
        If lastObj <> ObjPtr(obj) Then
            startIndex = 0
            counter = 0
            lastObj = ObjPtr(obj)
            If t = "TextBox" Then
                internalData = Split(obj.Text, IIf(Len(key) = 0, vbCrLf, key), , vbTextCompare)
                lastKey = key
            End If
        Else
            'same object but key changed..they must have aborted an enum loop early and changed key
            If t = "TextBox" And lastKey <> key Then
                startIndex = 0
                counter = 0
                lastObj = ObjPtr(obj)
                internalData = Split(obj.Text, IIf(Len(key) = 0, vbCrLf, key), , vbTextCompare)
                lastKey = key
            End If
        End If
    ElseIf IsArray(obj) Then
        If lastObj <> VarPtr(obj) Then
            startIndex = 0
            counter = 0
            lastObj = VarPtr(obj)
        End If
    ElseIf t = "String" Then
        If lastObj <> StrPtr(obj) Then
            startIndex = 0
            counter = 0
            lastObj = StrPtr(obj)
            If Len(key) > 0 Then
                internalData = Split(obj, key, , vbTextCompare)
                lastKey = key
            Else
                b() = StrConv(obj, vbFromUnicode, &H409)
                internalData = b()
                lastKey = Empty
            End If
        Else
            'same object but key changed..they must have aborted an enum loop early and changed key
            If t = "String" And lastKey <> key Then
                startIndex = 0
                counter = 0
                lastObj = StrPtr(obj)
                If Len(key) > 0 Then
                    internalData = Split(obj, key, , vbTextCompare)
                    lastKey = key
                Else
                    b() = StrConv(obj, vbFromUnicode, &H409)
                    internalData = b()
                    lastKey = Empty
                End If
            End If
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
    ElseIf t = "ListBox" Or t = "ComboBox" Then
        If obj.ListCount = 0 Then GoTo loopDone
        If counter = obj.ListCount Then GoTo loopDone
        If startIndex < 0 Then counter = 0
    ElseIf t = "TextBox" Then
        If AryIsEmpty(internalData) Then GoTo loopDone
        If counter > UBound(internalData) Then GoTo loopDone
        If startIndex < 0 Then counter = 0
    ElseIf t = "String" Then
        If AryIsEmpty(internalData) Then GoTo loopDone
        If counter > UBound(internalData) Then GoTo loopDone
        If startIndex < 0 Then counter = 0
    Else
        Err.Raise "Invalid Source, expects array or collection - given type: " & t, "enumerate()"
    End If
        
    If t = "ListBox" Or t = "ComboBox" Then
        value = obj.List(counter)
    ElseIf t = "TextBox" Then
        value = internalData(counter)
    ElseIf t = "String" Then
        If TypeName(internalData(counter)) = "Byte" Then
            value = Chr(internalData(counter))
        Else
            value = internalData(counter)
        End If
    Else
        If IsObject(obj(counter)) Then
            Set value = obj(counter)
        Else
            value = obj(counter)
        End If
    End If
    
    startIndex = counter  'current item index
    counter = counter + 1 'advance to next one (limitation: nesting of enumerate calls not supported)
    enumerate = True
    
Exit Function

loopDone:
    counter = 0
    startIndex = -1
    lastObj = 0
    clearVal internalData
    lastKey = Empty
    enumerate = False
    
End Function


Function enumerate_v1(ByRef obj, ByRef value As Variant, ByRef startIndex, Optional ByRef key) As Boolean
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
    enumerate_v1 = True
    
Exit Function

loopDone:
    counter = 0
    startIndex = -1
    lastObj = 0
    enumerate_v1 = False
    
End Function


