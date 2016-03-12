VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLowest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public allowNeg As Boolean
Public allowZero As Boolean
Public selIndex As Long
Public errorReturn As Long

Private Const MAX_LONG As Long = 2147483647

'note: if you pass in an array() dont utilize the paramarray() and submit multiple
'      you actually can, but dont expect selIndex to work..
'maxlong = error basically..
 Function lowest(ParamArray vals())
       
    Dim tmp As Long, i As Long
    
    tmp = MAX_LONG
    
    selIndex = -1
    
    If allowNeg Then allowZero = True
    
    For Each x In vals
        If IsArray(x) Then
            For Each xx In x
                setIfLowest xx, tmp, i
                i = i + 1
            Next
            Exit For
        Else
            setIfLowest x, tmp, i
            i = i + 1
        End If
    Next
                    
    'in case the user has changed the errorReturn to something they desire..
    If tmp = MAX_LONG Then tmp = errorReturn
    lowest = tmp
        
 End Function

 

 Private Sub setIfLowest(curVal, ByRef curLow, index As Long)

        If curVal < 0 Then
            If allowNeg Then
                If curVal < curLow Then
                    curLow = curVal
                    selIndex = index
                End If
            End If
        ElseIf curVal = 0 Then
            If allowZero Then
                If curVal < curLow Then
                    curLow = curVal
                    selIndex = index
                End If
            End If
        Else
            If curVal < curLow Then
                curLow = curVal
                selIndex = index
            End If
        End If
        
 End Sub

Private Sub Class_Initialize()
    errorReturn = MAX_LONG
End Sub
