Attribute VB_Name = "Init_Form_Code"
Sub init_form(formname As Form)

'sets target formname text box(es) text blank and clears cbo(s)
Dim savemask As String
Dim ctrl As control
For Each ctrl In formname.Controls
 If TypeOf ctrl Is TextBox Then ctrl.Text = ""  ' clear all text boxes
 If TypeOf ctrl Is MaskEdBox Then   ' clear mask edit boxes
   savemask = ctrl.Mask
   ctrl.Mask = ""
   ctrl.Text = ""
   ctrl.Mask = savemask
 End If
 If TypeOf ctrl Is CheckBox Then ctrl.Value = vbUnchecked   'uncheck all checkboxes
 If TypeOf ctrl Is ComboBox Then ctrl.Clear  'clear all combo boxes
 Next
End Sub
Public Sub Text2Tip(theForm)
 Dim ctrl As control
 Screen.MousePointer = vbHourglass
 For Each ctrl In theForm.Controls
  If TypeOf ctrl Is TextBox Or TypeOf ctrl Is MaskEdBox Then
    ctrl.ToolTipText = ctrl.Text
  End If
 Next
 Screen.MousePointer = vbDefault
End Sub


Public Sub Tip2Text()
 If (TypeOf Screen.ActiveControl Is TextBox Or TypeOf Screen.ActiveControl Is MaskEdBox) Then
   Screen.ActiveControl.Text = Screen.ActiveControl.ToolTipText
 End If
End Sub


