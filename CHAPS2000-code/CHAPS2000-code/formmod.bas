Attribute VB_Name = "Form_routines"
Sub centerform(theform As Form, XPOS As Long, YPOS As Long)
' THEFORM: represents the form you are centering on the screen
'    XPOS: X POSITION ON THE SCREEN. IF 0 IS SENT IN THEN
'          THE FORM IS CENTERED ON THE X AXIS ELSE IF 0 IS SENT
'          SENT IN THEN IT IS PLACE ON THE X AXIS WHEREVER
'          THE XPOS Is LOCATED
'    YPOS: Y POSITION ON THE SCREEN. IF 0 IS SENT IN THEN
'          THE FORM IS CENTERED ON THE Y AXIS ELSE IF 0 IS
'          SENT IN THEN IT IS PLACE ON THE Y AXIS WHEREVER
'          THE YPOS Is LOCATED
 Screen.MousePointer = vbHourglass
 If XPOS = 0 Then
   theform.Top = Screen.Height / 2 - theform.Height / 2
  Else
   theform.Top = XPOS
 End If
 If YPOS = 0 Then
   theform.Left = Screen.Width / 2 - theform.Width / 2
  Else
   theform.Left = YPOS
 End If
 Screen.MousePointer = vbDefault
End Sub
Public Sub disable_controls(theform As Form)
 Dim i As Integer
 For i = 0 To theform.Controls.count - 1
  theform.Controls(i).Enabled = False
 Next
End Sub


Sub centermdiform(theform As Form, mainform As Form, XPOS As Long, YPOS As Long)
'  THEFORM: represents the form you are centering on the screen
' MAINFORM: represents the PARENT FORM FOR THE MDI PROJECT
'     XPOS: X POSITION ON THE SCREEN. IF 0 IS SENT IN THEN
'           THE FORM IS CENTERED ON THE X AXIS ELSE IF 0 IS SENT
'           SENT IN THEN IT IS PLACE ON THE X AXIS WHEREVER
'           THE XPOS Is LOCATED
'     YPOS: Y POSITION ON THE SCREEN. IF 0 IS SENT IN THEN
'           THE FORM IS CENTERED ON THE Y AXIS ELSE IF 0 IS
'           SENT IN THEN IT IS PLACE ON THE Y AXIS WHEREVER
'           THE YPOS Is LOCATED
 Screen.MousePointer = vbHourglass
 If XPOS = 0 Then
   theform.Top = (mainform.ScaleHeight / 2) - (theform.Height / 2)
   If theform.MDIChild = False Then
     theform.Top = theform.Top + ((mainform.Height - mainform.ScaleHeight) / 1.4)
   End If
  Else
   theform.Top = XPOS
 End If
 If YPOS = 0 Then
   theform.Left = (mainform.ScaleWidth / 2) - (theform.Width / 2)
   If theform.MDIChild = False Then
     theform.Left = theform.Left - (mainform.Width - mainform.ScaleWidth)
   End If
  Else
   theform.Left = YPOS
 End If
 Screen.MousePointer = vbDefault
End Sub

