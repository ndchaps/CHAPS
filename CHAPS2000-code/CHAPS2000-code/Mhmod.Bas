Attribute VB_Name = "Mh_Routines"
Option Explicit

Function Find_In_Listbox_Col_String(TheControl As Control, thecol As Integer, thestring As String) As Integer
 TheControl.ColInstr = thecol
 TheControl.Col = thecol
 TheControl.FindInstr = -1
 TheControl.FoundIndex = -1
Do
  TheControl.FindInstr = thestring
  If TheControl.FoundIndex = -1 Then Exit Do
  If UCase$(thestring) = UCase$(left$(TheControl.ColList(TheControl.FoundIndex), Len(thestring))) Then
   Exit Do
  End If
'  thecontrol.FoundIndex = thecontrol.FoundIndex + 1
Loop

  Find_In_Listbox_Col_String = TheControl.FoundIndex

End Function

Function Find_In_Listbox_Col_StartString(TheControl As Control, thecol As Integer, thestring As String) As Integer
 TheControl.ColInstr = thecol
 TheControl.FindString = -1
 TheControl.FoundIndex = 0
 TheControl.FindString = thestring
 Find_In_Listbox_Col_StartString = TheControl.FoundIndex
End Function

Sub Update_mh_ListBoxes(controlname$, SearchCol%, Search$, replace$)
 ' controlname$ - name of the control we are looking for
 ' searchcol%   - column number to search
 ' search$      - the text we are searching for
 ' replace$     - the text we are replacing with
 '                make "" if it is just a delete
 '                if need more then one col then
 '                embed the chr$9) in the string
 Dim indx%, INDX2%, SAVEINDEX%
 Screen.MousePointer = vbHourglass
 For indx% = 0 To Forms.Count - 1
  For INDX2% = 0 To Forms(indx%).Controls.Count - 1
   If TypeOf Forms(indx%).Controls(INDX2%) Is Mh3dList Then
     If UCase$(Forms(indx%).Controls(INDX2%).Name) = UCase$(controlname$) Then
       If Len(Search$) Then
         Forms(indx%).Controls(INDX2%).ColInstr = SearchCol%
         Forms(indx%).Controls(INDX2%).FoundIndex = -1
         Forms(indx%).Controls(INDX2%).FindString = Search$
         If Forms(indx%).Controls(INDX2%).FoundIndex <> -1 Then
          'REMOVE THE ITEM FROM THE LIST BOX
           SAVEINDEX% = Forms(indx%).Controls(INDX2%).FoundIndex
           Forms(indx%).Controls(INDX2%).RemoveItem Forms(indx%).Controls(INDX2%).FoundIndex
         End If
       End If
       If Len(replace$) Then
         Forms(indx%).Controls(INDX2%).additem replace$, SAVEINDEX%
         Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).LastAdded
       End If
       Exit For
     End If
   End If
  Next INDX2%
 Next indx%
 Screen.MousePointer = vbDefault
End Sub

Public Sub Update_Multi_Mh_Listboxes(hmlists, list$(), Col(), ID$(), replace$())
 ' this routine updates the list boxes on forms that have multiple list boxes
 ' like grower/field form
 ' for each list box execpt the last it will call the click event for then list box

 ' hmlists    = the number of list to scan thru to put in the correct id
 ' list$()    = the array of lists to change
 ' col()      = the col of where the id is in each list box
 ' id$()      = the id we are looking for in the list box
 ' replace$() = the replacement string for the list box
 '               if "" then it is a delete else it is the string in the list box with tab's for the col changes
 Dim indx%, INDX2%, index3%, savecaption$
 'savecaption$ = MDIMAIN!pnlhint.Caption
 'MDIMAIN!pnlhint.Caption = "Updating Lists..."
 Screen.MousePointer = vbHourglass
 For index3% = 1 To hmlists
 For indx% = 0 To Forms.Count - 1
   For INDX2% = 0 To Forms(indx%).Controls.Count - 1
    If TypeOf Forms(indx%).Controls(INDX2%) Is Mh3dList Then
      If UCase$(Forms(indx%).Controls(INDX2%).Name) = UCase$(list$(index3%)) Then
        If Len(ID$(index3%)) Then
          Forms(indx%).Controls(INDX2%).ColInstr = Col(index3%)
          Forms(indx%).Controls(INDX2%).FoundIndex = -1
          Forms(indx%).Controls(INDX2%).FindString = ID$(index3%)
          If Forms(indx%).Controls(INDX2%).FoundIndex <> -1 Then
           ' REMOVE THE ITEM FROM THE LIST BOX
            If left$(replace$(index3%), 8) <> "setindex" Then
              Forms(indx%).Controls(INDX2%).RemoveItem Forms(indx%).Controls(INDX2%).FoundIndex
            End If
          End If
        End If
        If Len(replace$(index3%)) Then
          If left$(replace$(index3%), 8) = "setindex" Then
            If Forms(indx%).Controls(INDX2%).FoundIndex <> -1 Then
              Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).FoundIndex
             Else
              Forms(indx%).Controls(INDX2%).additem Mid$(replace$(index3%), 9)
              Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).LastAdded
            End If
            Screen.MousePointer = vbHourglass ' turn the hourglass back on in case it got turned off
           Else
            Forms(indx%).Controls(INDX2%).additem replace$(index3%)
            Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).LastAdded
          End If
         Else
          If Forms(indx%).Controls(INDX2%).ListIndex < 0 Then Forms(indx%).Controls(INDX2%).ListIndex = 0
        End If
        Exit For
      End If
    End If
   Next INDX2%
  Next indx%
 Next index3%
 Screen.MousePointer = vbDefault
 'MDIMAIN!pnlhint.Caption = savecaption$
End Sub
