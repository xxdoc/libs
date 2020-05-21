VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WitheventsArray Test"
   ClientHeight    =   3264
   ClientLeft      =   2016
   ClientTop       =   1932
   ClientWidth     =   4104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3264
   ScaleWidth      =   4104
   Begin ComctlLib.ListView lstObjs 
      Height          =   1500
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   2646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Object"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox lstEvents 
      Height          =   1200
      Left            =   15
      TabIndex        =   5
      Top             =   1980
      Width           =   4050
   End
   Begin VB.CommandButton cmdEvent 
      Caption         =   "Raise the &event"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2325
      TabIndex        =   3
      Top             =   1005
      Width           =   1725
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Object"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2325
      TabIndex        =   2
      Top             =   645
      Width           =   1725
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Object"
      Height          =   345
      Left            =   2325
      TabIndex        =   1
      Top             =   285
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Events:"
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   1770
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Objects:"
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   585
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'
' EventCollection Test Project
'
'********************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on:
'            * Windows XP Pro SP1
'            * VB6 SP5
'
' History:
'           01/02/2003 * This code replaces the old EventCollection
'                        class.
'
'********************************************************************************
Option Explicit

Dim WithEvents oEvntColl As EventCollection
Attribute oEvntColl.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
Static lIdx As Long

   oEvntColl.Add New TestClass, , "Object" & lIdx
   lstObjs.ListItems.Add , , "Object" & lIdx
   
   lIdx = lIdx + 1
   
End Sub

Private Sub cmdEvent_Click()
   
   On Error Resume Next
   
   oEvntColl.Item(lstObjs.SelectedItem.Text).object.oRaiseTheEvent
 
End Sub

Private Sub cmdRemove_Click()
   
   If Not lstObjs.SelectedItem Is Nothing Then
      
      With lstObjs
      
         oEvntColl.Remove .SelectedItem.Text
         .ListItems.Remove .SelectedItem.Index
         
      End With
   
   End If
   
End Sub


Private Sub Form_Load()
Dim BtnEvnt As CButtonEventWrapper

   ' Create the weCollection object
   Set oEvntColl = New EventCollection
   
   ' Add the listview to the array
   oEvntColl.Add lstObjs.object, , "lstObjs"
   
   ' Add the object to the listview
   lstObjs.ListItems.Add(, , "ListView").Tag = 0
   
   ' Create a button wrapper object
   Set BtnEvnt = New CButtonEventWrapper
   
   ' Set the button
   Set BtnEvnt.Button = cmdAdd
   
   ' Add the button to the array
   oEvntColl.Add BtnEvnt, , "cmdAdd"
   
   ' Add the button to the listview
   lstObjs.ListItems.Add(, , "cmdAdd").Tag = 0
   
End Sub
Private Sub lstObjs_ItemClick(ByVal Item As ComctlLib.ListItem)

   If Left$(Item.Text, 6) = "Object" Then
      cmdEvent.Enabled = True
      cmdRemove.Enabled = True
   Else
      cmdEvent.Enabled = False
      cmdRemove.Enabled = False
   End If
   
End Sub


Private Sub oEvntColl_HandleEvent(ByVal ObjectInfo As EventColl2.ObjectInfo, ByVal EventInfo As EventColl2.EventInfo)
Dim sEvent As String
Dim lMax As Long
Dim lIdx As Long

   sEvent = ObjectInfo.Key & "_" & EventInfo.Name & "("
   
   If EventInfo.Parameters.Count > 0 Then
      
      With EventInfo.Parameters
      
         For lIdx = 1 To .Count
         
            If VarType(.Item(lIdx)) = vbString Then
               sEvent = sEvent & """" & CStr(.Item(lIdx)) & """"
            Else
               sEvent = sEvent & CStr(.Item(lIdx))
            End If
         
            If lIdx < .Count Then sEvent = sEvent & ", "
         
         Next
         
         If Left$(ObjectInfo.Key, 6) = "Object" Then
            .Item(2) = "Test"
         End If
      
      End With
      
   End If
   
   sEvent = sEvent & ")"
   
   lstEvents.AddItem sEvent
   lstEvents.ListIndex = lstEvents.ListCount - 1

End Sub


