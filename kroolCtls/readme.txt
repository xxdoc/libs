
This is a trimmed down version of the CommonControls Replacement 
project released by Kroll.

This source snapshot was taken 6.19.15

Original forum link:

http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29

a full ocx version of the code is here:

http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29&p=4716949&viewfull=1#post4716949


Notes:
- When using the SetParent API, then you should pass .hWndUserControl and not .hWnd to it.
- When changing the "Project Name", then you should have all forms open, else all properties are lost. Reason is due to the fact that the library to which the controls are referring is the "Project Name" itself. Keeping all forms open will ensure that the .frx files will be updated with the new "Project Name".
- In order to trap error raises via "On Error Goto ..." or "On Error Resume Next" it is necessary to have "Break on Unhandled Errors" selected instead of "Break in Class Module" on Tools -> Options... -> General -> Error Trapping.
- If you want to embed the controls into another UserControl then you need to add the following code (Post #597) into your UserControl. As else the accelerator keys like the Left/Right key will not work. This issue is only relevant when using the Std-EXE Version. (The OCX Version will just work fine without any additional code)


A serious bug? The TextBoxW and RichTextBox can't move cursor by Left/Right Key after adding into a UserControl.
Version: 29Dec2014
This is not a bug. The VB.Form is forwarding the IOleInPlaceActiveObject::TranslateAccelerator method to the UserControl. If now a UserControl (e.g. TextBoxW) is embedded into another UserControl then VB will not forward the IOleInPlaceActiveObject::TranslateAccelerator method to the underlying UserControl.

So in each chain of this embedding the IOleInPlaceActiveObject interface must be implemented and forwarded.

In your case, probably only the case where my UserControl (e.g. TextBoxW) is embedded in "your" UserControl you need to add the following code to "your" UserControl:

Code:

Implements OLEGuids.IOleInPlaceActiveObjectVB

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
On Error Resume Next
Dim This As OLEGuids.IOleInPlaceActiveObjectVB
Set This = UserControl.ActiveControl.Object
This.TranslateAccelerator Handled, RetVal, wMsg, wParam, lParam, Shift
End Sub

This topic was already discussed in this thread. 

