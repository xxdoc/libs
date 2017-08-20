VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStorageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsStorageItem.cls - file storage item class
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Public Enum FileFlags
    FF_REPLACEONEXISTS = 1   ' // Replace if file already exists
    FF_IGNOREERROR = 2      ' // Ignore errors
End Enum

Private mFileName           As String       ' // File name
Private mDestinationPath    As String       ' // Destination path
Private mSize               As Currency     ' // Size of file
Private mFlags              As FileFlags    ' // Flags
Private mProject            As clsProject   ' // Parent project

' // Parent project
Public Property Get Project() As clsProject
    Set Project = mProject
End Property
Public Property Set Project( _
                    ByRef Value As clsProject)
    Set mProject = Value
End Property

' // File name
Public Property Get fileName() As String
    fileName = mFileName
End Property
Public Property Let fileName( _
                    ByVal Value As String)
    mFileName = Value
    mProject.Modify = True
End Property

' // Destination path
Public Property Get DestinationPath() As String
    DestinationPath = mDestinationPath
End Property
Public Property Let DestinationPath(ByVal Value As String)
    mDestinationPath = Value
    mProject.Modify = True
End Property

' // Flags
Public Property Get Flags() As FileFlags
    Flags = mFlags
End Property
Public Property Let Flags(ByVal Value As FileFlags)
    mFlags = Value
    mProject.Modify = True
End Property

' // Save item to property bag
Public Function Save() As PropertyBag

    Set Save = New PropertyBag
    
    Save.WriteProperty "FileName", mFileName
    Save.WriteProperty "DestinationPath", mDestinationPath
    Save.WriteProperty "Flags", mFlags
    
End Function

' // Load item from property bag
Public Function Load( _
                ByRef Prop As PropertyBag) As Boolean
                
    mFileName = Prop.ReadProperty("FileName")
    mDestinationPath = Prop.ReadProperty("DestinationPath")
    mFlags = Prop.ReadProperty("Flags")
    
    mProject.Modify = True
    Load = True
    
End Function
