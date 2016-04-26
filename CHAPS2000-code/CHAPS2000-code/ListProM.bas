Attribute VB_Name = "ListProModules"
Option Explicit

Public Enum SelectWindowType
 SingleSelect = 1
 MultiSelect = 2
End Enum

Public Sub SearchListBox(SearchLst As fpList, SearchTxt As TextBox, SearchCol As Integer)
'*****************************************************************************
'***
'*** Procedure Name: SearchListBox
'***     Created By: Lynn
'***   Date Created: 10/10/2007
'***
'***   Modify Dates:
'***
'***    Description:
'***
'*** Parameter List:
'***
'*****************************************************************************

 Dim Found As Long
 On Local Error GoTo LeHandle

 If Len(SearchTxt.text) = 0 Then Exit Sub
 Found = Find_In_ListPro_Listbox_Col_StringLong(SearchLst, SearchCol, SearchTxt.text)
 If Found <> -1 Then
   SearchLst.ListIndex = Found
   SearchLst.TopIndex = Found
  Else
   SearchTxt.text = Left$(SearchTxt.text, Len(SearchTxt.text) - 1)
   SearchTxt.SelStart = Len(SearchTxt.text)
 End If

 On Local Error GoTo 0
Exit Sub

LeHandle:
 err.Raise err.Number, err.Source & "ListProModules - SearchListBox"

End Sub
Public Sub SetSort(ColtoSort As Integer, ListBox As Control, SearchLabel As Control)
 Screen.MousePointer = vbHourglass
 Dim Col As Long
 Dim header As String
 With ListBox
  .SortState = SortStateSuspend
  For Col = 0 To .Columns - 1
   .Col = Col
   header = .ColHeaderText
   If Right$(header, 1) = "*" Then
     header = Left$(header, Len(header) - 1)
   End If
   If Col = ColtoSort Then
     SearchLabel.Caption = header
     .ColHeaderText = header & "*"
     .ColSorted = SortedAscending
     .ColSortSeq = 0
    Else
     .ColHeaderText = header
     .ColSorted = SortedNone
     .ColSortSeq = -1
   End If
  Next Col
  .SortState = SortStateActiveReSort
 End With
 Screen.MousePointer = vbDefault
End Sub

Public Sub SetSortNoSearch(ColtoSort As Integer, ListBox As Control)
   Screen.MousePointer = vbHourglass
   Dim Col As Long
   Dim header As String
   With ListBox
      .SortState = SortStateSuspend
      For Col = 0 To .Columns - 1
         .Col = Col
         header = .ColHeaderText
         If Right$(header, 1) = "*" Then
            header = Left$(header, Len(header) - 1)
         End If
         If Col = ColtoSort Then
            .ColHeaderText = header & "*"
            .ColSorted = SortedAscending
            .ColSortSeq = 0
         Else
            .ColHeaderText = header
            .ColSorted = SortedNone
            .ColSortSeq = -1
         End If
      Next Col
      .SortState = SortStateActiveReSort
   End With
   Screen.MousePointer = vbDefault
End Sub

Public Function Find_In_ListPro_Listbox_Col_String_BigList(TheControl As Control, thecol As Integer, thestring As String) As Long
With TheControl
 .SearchMethod = SearchMethodPartialMatch
 .ColumnSearch = thecol
 .Col = thecol
 .SearchText = thestring
 .SearchIndex = -1
 .action = ActionSearch 'do search
 
 'returns row index of match or returns -1 if no match
 Find_In_ListPro_Listbox_Col_String_BigList = .SearchIndex
End With
End Function



Public Sub Set_Combo_LP(theCombo As Control, TheData As String, thecol As Integer)
  Dim Found As Integer
  Found = Find_In_ListPro_Listbox_Col_String(theCombo, thecol, TheData)
  If Found <> -1 Then
    theCombo.ListIndex = Found
  End If
End Sub


Public Function Find_In_ListPro_Listbox_Col_String(TheControl As Control, thecol As Integer, thestring As String) As Integer
With TheControl
 .SearchMethod = SearchMethodPartialMatch
 .ColumnSearch = thecol
 .Col = thecol
 .SearchText = thestring
 .SearchIndex = -1
 .action = ActionSearch 'do search
 
 'returns row index of match or returns -1 if no match
 Find_In_ListPro_Listbox_Col_String = .SearchIndex
End With
End Function

Public Function Find_In_ListPro_Listbox_Col_StringLong(TheControl As Control, thecol As Integer, thestring As String) As Long
With TheControl
 .SearchMethod = SearchMethodPartialMatch
 .ColumnSearch = thecol
 .Col = thecol
 .SearchText = thestring
 .SearchIndex = -1
 .action = ActionSearch 'do search
 
 'returns row index of match or returns -1 if no match
 Find_In_ListPro_Listbox_Col_StringLong = .SearchIndex
End With
End Function

Public Function Find_In_ListPro_Listbox_Col_WholeString(TheControl As Control, thecol As Integer, thestring As String) As Integer
With TheControl
 .SearchMethod = SearchMethodExactMatch
 .ColumnSearch = thecol
 .Col = thecol
 .SearchText = thestring
 .SearchIndex = -1
 .action = ActionSearch 'do search
 
 'returns row index of match or returns -1 if no match
 Find_In_ListPro_Listbox_Col_WholeString = .SearchIndex
End With
End Function

Public Function Find_In_ListPro_Listbox_Col_WholeStringLong(TheControl As Control, thecol As Integer, thestring As String) As Long
With TheControl
 .SearchMethod = SearchMethodExactMatch
 .ColumnSearch = thecol
 .Col = thecol
 .SearchText = thestring
 .SearchIndex = -1
 .action = ActionSearch 'do search
 
 'returns row index of match or returns -1 if no match
 Find_In_ListPro_Listbox_Col_WholeStringLong = .SearchIndex
End With
End Function

Public Sub UpdateListProListBoxes(ControlName$, SearchCol%, Search$, Replace$)
 Dim frm As Form
 Dim ctrl As Control
 Dim Found As Long
 Screen.MousePointer = vbHourglass
 For Each frm In Forms
  For Each ctrl In frm.Controls
   If TypeOf ctrl Is fpList Then
    If LCase$(ctrl.name) = LCase$(ControlName$) Then
      Found = Find_In_ListPro_Listbox_Col_WholeStringLong(ctrl, SearchCol%, Search$)
      If Found <> -1 Then ctrl.RemoveItem Found
      If Len(Replace$) Then
        ctrl.AddItem Replace$
        ctrl.ListIndex = ctrl.newindex
       Else
        ctrl.ListIndex = 0
      End If
    End If
   End If
  Next
 Next
 Screen.MousePointer = vbDefault
End Sub

Public Sub UpdateListProListBoxesMultiColumn(ControlName$, SearchCol() As Integer, Search() As String, Replace$)
 Dim frm As Form
 Dim ctrl As Control
 Dim Found As Integer
 Dim Row As Long
 Dim ColSearch As Integer
 Dim strText As String
 Screen.MousePointer = vbHourglass
 For Each frm In Forms
  For Each ctrl In frm.Controls
  If TypeOf ctrl Is fpList Then
    If LCase$(ctrl.name) = LCase$(ControlName$) Then
      ctrl.ListIndex = 0
      Found = -1
      For Row = 0 To ctrl.ListCount - 1
       ctrl.Row = Row
       Found = Row
       For ColSearch = 1 To SearchCol(0)
        ctrl.Col = SearchCol(ColSearch)
        strText = ctrl.ColList(ctrl.Col, Row)
        If LCase$(strText) <> LCase$(Search(ColSearch)) Then
          Found = -1
          Exit For
        End If
       Next ColSearch
       If Found <> -1 Then Exit For
      Next Row
      If Found <> -1 Then ctrl.RemoveItem Found
      If Len(Replace$) Then
        ctrl.AddItem Replace$
        ctrl.ListIndex = ctrl.newindex
      End If
    End If
   End If
  Next
 Next
 Screen.MousePointer = vbDefault
End Sub


Public Sub Update_Multi_ListPro_Listboxes(hmlists, list$(), Col(), ID$(), Replace$())
 ' this routine updates the list boxes on forms that have multiple list boxes
 ' like grower/field form
 ' for each list box execpt the last it will call the click event for then list box

 ' hmlists    = the number of list to scan thru to put in the correct id
 ' list$()    = the array of lists to change
 ' col()      = the col of where the id is in each list box
 ' id$()      = the id we are looking for in the list box
 ' replace$() = the replacement string for the list box
 '               if "" then it is a delete else it is the string in the list box with tab's for the col changes
 Dim indx%, INDX2%, index3%
 Dim Found As Long
 Screen.MousePointer = vbHourglass
 For index3% = 1 To hmlists
  For indx% = 0 To Forms.count - 1
   For INDX2% = 0 To Forms(indx%).Controls.count - 1
    If TypeOf Forms(indx%).Controls(INDX2%) Is fpList Then
      If UCase$(Forms(indx%).Controls(INDX2%).name) = UCase$(list$(index3%)) Then
        If Len(ID$(index3%)) Then
          Found = Find_In_ListPro_Listbox_Col_WholeStringLong(Forms(indx%).Controls(INDX2%), CInt(Col(index3%)), ID$(index3%))
           If Found <> -1 Then
           ' REMOVE THE ITEM FROM THE LIST BOX
            If Left$(Replace$(index3%), 8) <> "setindex" Then
              Forms(indx%).Controls(INDX2%).RemoveItem Found
            End If
          End If
        End If
        If Len(Replace$(index3%)) Then
          If Left$(Replace$(index3%), 8) = "setindex" Then
            If Found <> -1 Then
              Forms(indx%).Controls(INDX2%).ListIndex = -1
              Forms(indx%).Controls(INDX2%).ListIndex = Found
             Else
              'Forms(indx%).Controls(INDX2%).AddItem Mid$(replace$(index3%), 9)
              'Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).LastAdded
            End If
            Screen.MousePointer = vbHourglass ' turn the hourglass back on in case it got turned off
           Else
            Forms(indx%).Controls(INDX2%).AddItem Replace$(index3%)
            Forms(indx%).Controls(INDX2%).ListIndex = Forms(indx%).Controls(INDX2%).newindex
          End If
         Else
          If Forms(indx%).Controls(INDX2%).ListIndex < 0 And Forms(indx%).Controls(INDX2%).ListCount > 0 Then Forms(indx%).Controls(INDX2%).ListIndex = 0
        End If
        Exit For
      End If
    End If
   Next INDX2%
  Next indx%
 Next index3%
 Screen.MousePointer = vbDefault
End Sub

