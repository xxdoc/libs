VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsExecute.cls - list of execution command class
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private mCollectionItems    As Collection   ' // List of items
Private mProject            As clsProject   ' // Parent project

' // Parent project
Public Property Get Project() As clsProject
    Set Project = mProject
End Property
Public Property Set Project( _
                    ByRef Value As clsProject)
    Set mProject = Value
End Property

' // List item
Public Property Get Item( _
                    ByVal index As Long)
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = mCollectionItems(index + 1)
End Property

' // Number of items
Public Property Get Count() As Long
    Count = mCollectionItems.Count
End Property

' // IEnumVariant
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCollectionItems.[_NewEnum]
End Property

' // Remove from list
Public Sub Remove(ByVal index As Long)
    mCollectionItems.Remove (index + 1)
    mProject.Modify = True
End Sub

' // Clear list
Public Sub Clear()

    Do While mCollectionItems.Count
        mCollectionItems.Remove 1
    Loop
    
    mProject.Modify = True
    
End Sub

' // Add to list
Public Function Add( _
                ByVal Command As String, _
                ByVal Parameters As String, _
                Optional ByVal index As Long = -1) As clsExecuteItem
                
    Set Add = New clsExecuteItem
    Set Add.Project = mProject
    
    Add.fileName = Command
    Add.Parameters = Parameters
    
    If index = -1 Then
        mCollectionItems.Add Add
    Else: mCollectionItems.Add Add, , index + 1
    End If
    
    mProject.Modify = True
    
End Function

' // Save list to property bag
Public Function Save() As PropertyBag
    Dim itm As clsExecuteItem:  Dim idx As Long
    
    Set Save = New PropertyBag
    
    Save.WriteProperty "ItemCount", mCollectionItems.Count
    
    For Each itm In mCollectionItems
         Save.WriteProperty "Item" & idx, itm.Save.Contents
         idx = idx + 1
    Next
    
End Function

' // Load list from property bag
Public Function Load( _
                ByRef Prop As PropertyBag) As Boolean
    Dim Count   As Long:            Dim idx     As Long
    Dim buf()   As Byte:            Dim mCol    As Collection
    Dim itm     As clsExecuteItem:  Dim lProp   As PropertyBag
    
    Count = Prop.ReadProperty("ItemCount")
    
    Set mCol = New Collection
    
    For idx = 0 To Count - 1
    
        buf = Prop.ReadProperty("Item" & idx)
        
        Set itm = New clsExecuteItem
        Set itm.Project = mProject
        Set lProp = New PropertyBag
        
        lProp.Contents = buf
        itm.Load lProp
        mCol.Add itm
        
        Set lProp = Nothing
        Set itm = Nothing
        
    Next
    
    Me.Clear
    Set mCollectionItems = mCol
    mProject.Modify = False
    
    Load = True
    
End Function

' // Swap two items
Public Function Swap( _
                ByVal I1 As Long, _
                ByVal I2 As Long) As Boolean
    Dim itm1    As clsExecuteItem:  Dim itm2    As clsExecuteItem
    
    If I1 < 0 Or I1 >= mCollectionItems.Count Or I2 < 0 Or I2 >= mCollectionItems.Count Then Exit Function
    
    Set itm1 = mCollectionItems(I1 + 1)
    Set itm2 = mCollectionItems(I2 + 1)
    
    mCollectionItems.Remove I1 + 1
    If I1 >= mCollectionItems.Count Then mCollectionItems.Add itm2 Else mCollectionItems.Add itm2, , I1 + 1
    
    mCollectionItems.Remove I2 + 1
    If I2 >= mCollectionItems.Count Then mCollectionItems.Add itm1 Else mCollectionItems.Add itm1, , I2 + 1
    
    mProject.Modify = True
    Swap = True
    
End Function

Private Sub Class_Initialize()
    Set mCollectionItems = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCollectionItems = Nothing
End Sub
