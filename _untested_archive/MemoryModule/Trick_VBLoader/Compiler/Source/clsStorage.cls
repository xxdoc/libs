VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsStorage.cls - file storage class for unpacking
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private mMainExecutable     As Long         ' // Index of main executable file
Private mCollectionItems    As Collection   ' // List of storage items
Private mProject            As clsProject   ' // Parent project

' // Parent project
Public Property Get Project() As clsProject
    Set Project = mProject
End Property
Public Property Set Project( _
                    ByRef Value As clsProject)
    Set mProject = Value
End Property

' // Index of main executable file
Public Property Get MainExecutable() As Long
    MainExecutable = mMainExecutable
End Property
Public Property Let MainExecutable( _
                    ByVal Value As Long)
    mMainExecutable = Value
    mProject.Modify = True
End Property

' // Get storage item by index
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

' // Remove item from list
Public Sub Remove( _
           ByVal index As Long)
           
    If index = mMainExecutable Then
        mMainExecutable = -1
    ElseIf index < mMainExecutable Then
        mMainExecutable = mMainExecutable - 1
    End If
    
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

' // Add new item to list
Public Function Add( _
                ByVal SrcFileName As String, _
                ByVal DstPath As String, _
                ByVal Flags As FileFlags, _
                Optional ByVal index As Long = -1) As clsStorageItem
                
    Set Add = New clsStorageItem
    Set Add.Project = mProject
    
    ' // Get relative path
    SrcFileName = mProject.ToRelative(SrcFileName, mProject.BaseDirectory)
    
    Add.FileName = SrcFileName
    Add.DestinationPath = DstPath
    Add.Flags = Flags
    
    If index = -1 Then
        mCollectionItems.Add Add
    Else: mCollectionItems.Add Add, , index + 1
    End If
    
    mProject.Modify = True
    
End Function

' // Save storage to property bag
Public Function Save() As PropertyBag
    Dim itm As clsStorageItem
    Dim idx As Long
    
    Set Save = New PropertyBag
    
    Save.WriteProperty "MainExecutable", mMainExecutable
    Save.WriteProperty "ItemCount", mCollectionItems.Count
    
    For Each itm In mCollectionItems
         Save.WriteProperty "Item" & idx, itm.Save.Contents
         idx = idx + 1
    Next
    
End Function

' // Load storage from property bag
Public Function Load( _
                ByRef Prop As PropertyBag) As Boolean
    Dim Count   As Long:            Dim idx     As Long
    Dim buf()   As Byte:            Dim mCol    As Collection
    Dim itm     As clsStorageItem:  Dim lProp   As PropertyBag
    
    mMainExecutable = Prop.ReadProperty("MainExecutable")
    Count = Prop.ReadProperty("ItemCount")
    
    Set mCol = New Collection
    
    For idx = 0 To Count - 1
    
        ' // Read ecah item
        buf = Prop.ReadProperty("Item" & idx)
        Set itm = New clsStorageItem
        Set itm.Project = mProject
        Set lProp = New PropertyBag
        lProp.Contents = buf
        itm.Load lProp
        mCol.Add itm
        Set itm = Nothing
        Set lProp = Nothing
        
    Next
    
    Me.Clear
    Set mCollectionItems = mCol
    mProject.Modify = True
    Load = True
    
End Function

' // Swap two items
Public Function Swap( _
                ByVal i1 As Long, _
                ByVal i2 As Long) As Boolean
    Dim itm1 As clsStorageItem
    Dim itm2 As clsStorageItem
    
    If i1 < 0 Or i1 >= mCollectionItems.Count Or i2 < 0 Or i2 >= mCollectionItems.Count Then Exit Function
    
    Set itm1 = mCollectionItems(i1 + 1)
    Set itm2 = mCollectionItems(i2 + 1)
    
    mCollectionItems.Remove i1 + 1
    If i1 >= mCollectionItems.Count Then mCollectionItems.Add itm2 Else mCollectionItems.Add itm2, , i1 + 1
    
    mCollectionItems.Remove i2 + 1
    If i2 >= mCollectionItems.Count Then mCollectionItems.Add itm1 Else mCollectionItems.Add itm1, , i2 + 1
    
    If i1 = mMainExecutable Then mMainExecutable = i2 Else If i2 = mMainExecutable Then mMainExecutable = i1
    
    mProject.Modify = True
    Swap = True
    
End Function

' -
Private Sub Class_Initialize()
    mMainExecutable = -1
    Set mCollectionItems = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCollectionItems = Nothing
End Sub
