Attribute VB_Name = "modSubclass"
Option Explicit

'Design: you can have an unlimited number of classes, and should be
'        able to subclass an unlimited number of windows. Each class
'        can in turn subclass as many windows as it wants, and it should
'        be able to attach to as many messages per window as it wants.
'
'        You can edit message parameters and cancel them directly in
'        the eventhandler you implement for the clsSubClass in your code.
'
'       One thing you cannot do, is to have multiple classes subclass the
'       same window looking for the same message. (although multiple classes
'       should be able to share the subclass of any hwnd, they just cant
'       receive the same msg. I dont subclass a whole bunch so this is ok for me.
'
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_WNDPROC = (-4)

'collection of references to all of the initalized classes
'key= "Objptr:" & objptr(class) value = that class object
Private ActiveClasses As New Collection

'maps hwnds to oldProcs key= "Hwnd:" & hwnd  value=oldProc
Private cHwnds As New Collection

'key="Hwnd:" & hwnd & " Msg:" & msg,  value=ObjPtr(class to notify)
Private cMsgs As New Collection

'this count figure is global per hwnd basic key="Hwnd:" & hwnd value=count
Private cCount As New Collection

'needed to assure complete tear down key=hwnd value=hwnd
Private cHwndList As New Collection


Sub RegisterClassActive(clsActive As clsSubClass)
    ActiveClasses.Add clsActive, "Objptr:" & ObjPtr(clsActive)
End Sub

Sub RemoveActiveClass(clsTerminate As clsSubClass)
    On Error Resume Next
    ActiveClasses.Remove "Objptr:" & ObjPtr(clsTerminate)
    If ActiveClasses.Count = 0 Then InitTearDown
End Sub

Private Sub InitTearDown()
    On Error Resume Next
    Dim pOldProc As Long, hwnd As Long, i As Integer
    
    For i = 0 To cHwndList.Count
        hwnd = cHwndList(i)
        pOldProc = cHwnds("Hwnd:" & hwnd)
        SetWindowLong hwnd, GWL_WNDPROC, pOldProc
    Next
    
    Set cHwnds = Nothing
    Set cHwndList = Nothing
    
    'just in case so we dont run into weird bugs in some cases
    Set cHwnds = New Collection
    Set cHwndList = New Collection
    
End Sub


Sub MonitorWindowMessage(notify As clsSubClass, ByVal hwnd As Long, ByVal wMsg As Long)
    
    Dim pOldProc As Long, c As Long
    
    If IsWindow(hwnd) = False Then Err.Raise 1, , "Invalid Window Handle"

    If Not KeyExistsInCollection(cHwnds, "Hwnd:" & hwnd) Then
        'subclass each window only once
        pOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
        
        If pOldProc = 0 Then Err.Raise 2, , "Subclass failed hwnd: " & hwnd
        
        cCount.Add 1, "Hwnd:" & hwnd
        cHwnds.Add pOldProc, "Hwnd:" & hwnd
        cHwndList.Add hwnd, "Hwnd:" & hwnd
        cMsgs.Add ObjPtr(notify), "Hwnd:" & hwnd & " Msg:" & wMsg
        
   Else
        'window already subclassed
        'is this message already handled for this window?
        If Not KeyExistsInCollection(cMsgs, "Hwnd:" & hwnd & " Msg:" & wMsg) Then
            IncrementCollectionVal cCount, "Hwnd:" & hwnd
            cMsgs.Add ObjPtr(notify), "Hwnd:" & hwnd & " Msg:" & wMsg
        Else
            'we already have a class watching for this message on this hwnd now what?
            If ObjPtr(notify) = cMsgs("Hwnd:" & hwnd & " Msg:" & wMsg) Then
                'ok, they have a memory problem,same class watching same win, for samemsg?
                Err.Raise 3, , "Same class watching same window for same message?"
            Else
                'do we want to support different classes watching the same hwnd and msg?
                Err.Raise 4, , "Another class is already subclassing this window looking for this message! MsgNum:" & wMsg
            End If
        End If
    End If
                

End Sub

Sub DetachWindowMessage(ByVal hwnd As Long, ByVal wMsg As Long, Optional clsParent As clsSubClass)
    Dim pOldProc As Long
    
    If cCount("Hwnd:" & hwnd) <= 1 Then
        ' This is the last message, so remove subclass
        pOldProc = cHwnds("Hwnd:" & hwnd)
       
        SetWindowLong hwnd, GWL_WNDPROC, pOldProc
         
        cHwnds.Remove "Hwnd:" & hwnd
        cHwndList.Remove "Hwnd:" & hwnd
    Else
        DecrementCollectionVal cCount, "Hwnd:" & hwnd
    End If
    
End Sub


Function WindowProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pOldProc As Long, pClass As Long
    Dim mySubClass As clsSubClass
    Dim stopIt As Boolean
    
    pOldProc = cHwnds("Hwnd:" & hwnd)

    'find out who is handling this windows message it can be only one class
    If KeyExistsInCollection(cMsgs, "Hwnd:" & hwnd & " Msg:" & wMsg) Then
                
        pClass = cMsgs("Hwnd:" & hwnd & " Msg:" & wMsg)

        'we have the objPtr to the class to call, which is its key in
        'the ActiveClasses collection, collection val= class object if it
        'makes it to here, there should always be an this element in the col
        'if not..then we have to catch it like this.
        On Error Resume Next
        
        Set mySubClass = ActiveClasses("Objptr:" & pClass)
        
        If Err.Number <> 0 Then 'class terminated but subclass was still active :-\
            Err.Clear
            DetachWindowMessage hwnd, wMsg
        Else
            mySubClass.ForwardMessage hwnd, wMsg, wParam, lParam, stopIt
            Set mySubClass = Nothing
        End If
        
    End If
    
    'stopit can be set = true when event is raised from clsSubClass
    'any of the parameters passed in forward message could be changed
    If Not stopIt Then
        WindowProc = CallWindowProc(pOldProc, hwnd, wMsg, wParam, ByVal lParam)
    End If
    
    'If Not (mySubClass Is Nothing) Then
    '    mySubClass.ForwardMessage hwnd, wMsg, wParam, lParam, , True
    '    Set mySubClass = Nothing
    'End If
    
End Function


'and this is part of why collections can suck
Private Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Private Sub UpdateCollectionVal(c As Collection, key As String, newVal)
    c.Remove key
    c.Add newVal, key
End Sub

Private Sub IncrementCollectionVal(c As Collection, key As String)
    Dim x As Long
    x = c(key) + 1
    UpdateCollectionVal c, key, x
End Sub

Private Sub DecrementCollectionVal(c As Collection, key As String)
    Dim x As Long
    x = c(key) - 1
    UpdateCollectionVal c, key, x
End Sub
