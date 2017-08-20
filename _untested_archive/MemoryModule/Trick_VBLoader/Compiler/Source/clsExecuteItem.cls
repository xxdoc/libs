VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExecuteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsExecuteItem.cls - execution command class
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

' // Execution flags
Public Enum ExeFlags
    EF_IGNOREERROR = 1  ' // Ignore errors
End Enum

Private mFileName   As String       ' // Execution file name
Private mParameters As String       ' // Parameters
Private mFlags      As ExeFlags     ' // Flags
Private mProject    As clsProject   ' // Parent project

' // Parent project
Public Property Get Project() As clsProject
    Set Project = mProject
End Property
Public Property Set Project( _
                    ByRef Value As clsProject)
    Set mProject = Value
End Property

' // File name
Public Property Get FileName() As String
    FileName = mFileName
End Property
Public Property Let FileName( _
                    ByVal Value As String)
    mFileName = Value
    mProject.Modify = True
End Property

' // Command line parameters
Public Property Get Parameters() As String
    Parameters = mParameters
End Property
Public Property Let Parameters( _
                    ByVal Value As String)
    mParameters = Value
    mProject.Modify = True
End Property

' // Flags
Public Property Get Flags() As ExeFlags
    Flags = mFlags
End Property
Public Property Let Flags( _
                    ByVal Value As ExeFlags)
    mFlags = Value
    mProject.Modify = True
End Property

' // Save object to property bag
Public Function Save() As PropertyBag

    Set Save = New PropertyBag
    
    Save.WriteProperty "FileName", mFileName
    Save.WriteProperty "Parameters", mParameters
    Save.WriteProperty "Flags", mFlags
    
End Function

' // Load object from property bag
Public Function Load( _
                ByRef Prop As PropertyBag) As Boolean
                
    mFileName = Prop.ReadProperty("FileName")
    mParameters = Prop.ReadProperty("Parameters")
    mFlags = Prop.ReadProperty("Flags")
    
    Load = True
    
End Function
