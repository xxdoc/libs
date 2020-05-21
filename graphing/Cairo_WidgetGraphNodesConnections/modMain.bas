Attribute VB_Name = "modMain"
Option Explicit
 
Public fMain As New cfMain

Sub Main()
  AddImageResources 'see helper-function further below in this module
  ChangeToolTipWidgetDefaultSettings Cairo.ToolTipWidget.Widget '...same here...

  fMain.Form.Show
  
  Cairo.WidgetForms.EnterMessageLoop 'we require as Msg-Pump, since no VB-Forms are involved
End Sub

Private Sub AddImageResources()
  Cairo.ImageList.AddImage "BGBlack", App.Path & "\Res\BGBlack.png"
  Cairo.ImageList.AddImage "FormIco", App.Path & "\Res\FormIco.png"
  Cairo.ImageList.AddImage "NodeIco", App.Path & "\Res\NodeIco.png"
End Sub

Private Sub ChangeToolTipWidgetDefaultSettings(W As cWidgetBase)
  W.Alpha = 0.5 'a bit more than the default-transparency for the tooltips
  W.BackColor = RGB(255, 255, 100) 'a light yellow for the tooltips
  W.BorderColor = vbGreen 'a green border
End Sub
