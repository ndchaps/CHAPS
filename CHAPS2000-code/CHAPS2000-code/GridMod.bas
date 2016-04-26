Attribute VB_Name = "GridRoutines"
Option Explicit
Public Sub clear_grid(gridcntrl As Control)
   ' this clears out grid data not including headings
   ' gridcntrl = formname!gridname
   ' formname!gridname = name of grid control
 Dim Col As Long, Row As Long
 Screen.MousePointer = vbHourglass
 For Col = 1 To gridcntrl.MaxCols
  For Row = 1 To gridcntrl.MaxRows
   gridcntrl.SetText Col, Row, ""
  Next Row
 Next Col
 Screen.MousePointer = vbDefault
End Sub
Public Sub clear_grid2(gridcntrl As Control, startrow As Long, endrow As Long, startcol As Long, endcol As Long)
 Dim Col As Long, Row As Long
 Screen.MousePointer = vbHourglass
 For Col = startcol To endcol
  For Row = startrow To endrow
   gridcntrl.SetText Col, Row, ""
  Next Row
 Next Col
 Screen.MousePointer = vbDefault
End Sub

Public Sub ClearGridSelected(spread1 As Control)
 If Not spread1.IsBlockSelected Then Exit Sub
 Screen.MousePointer = vbHourglass
 spread1.ReDraw = False
 spread1.Row = spread1.SelBlockRow
 spread1.Row2 = spread1.SelBlockRow2
 If spread1.Row = -1 Then spread1.Row = 1
 If spread1.Row2 = -1 Then spread1.Row2 = spread1.MaxRows
 spread1.Col = 1
 spread1.col2 = spread1.MaxCols
 spread1.BlockMode = True
 spread1.Action = SS_ACTION_DELETE_ROW
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 spread1.col2 = 0
 spread1.Row2 = 0
 spread1.BlockMode = False
 spread1.ReDraw = True
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 Screen.MousePointer = vbDefault
End Sub

Public Sub DisableGrid(spread1 As Control)
 Dim Col As Long
 spread1.Row = -1
 For Col = 1 To spread1.MaxCols
  spread1.Col = Col
  spread1.Lock = True
 Next Col
End Sub

Public Sub SelectClearGrid(spread1 As Control, col1 As Long, col2 As Long, row1 As Long, Row2 As Long)
 Screen.MousePointer = vbHourglass
 spread1.ReDraw = False
 spread1.Row = row1
 spread1.Row2 = Row2
 spread1.Col = col1
 spread1.col2 = col2
 spread1.Action = SS_ACTION_SELECT_BLOCK

 If Not spread1.IsBlockSelected Then
   Screen.MousePointer = vbDefault
   Exit Sub
 End If
 spread1.ReDraw = False
 spread1.Row = spread1.SelBlockRow
 spread1.Row2 = spread1.SelBlockRow2
 spread1.Col = spread1.SelBlockCol
 spread1.col2 = spread1.SelBlockCol2
 spread1.BlockMode = True
 spread1.Action = SS_ACTION_CLEAR
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 spread1.Row = 0
 spread1.Row2 = 0
 spread1.Col = 0
 spread1.col2 = 0
 spread1.ReDraw = True
 spread1.Refresh
 spread1.ClearSelection
 Screen.MousePointer = vbDefault
 
End Sub
Public Sub SelectDeleteGrid(spread1 As Control, col1 As Long, col2 As Long, row1 As Long, Row2 As Long)
 Screen.MousePointer = vbHourglass
 spread1.ReDraw = False
 spread1.Row = row1
 spread1.Row2 = Row2
 spread1.Col = col1
 spread1.col2 = col2
 spread1.Action = SS_ACTION_SELECT_BLOCK

 If Not spread1.IsBlockSelected Then Exit Sub
 spread1.ReDraw = False
 spread1.Row = spread1.SelBlockRow
 spread1.Row2 = spread1.SelBlockRow2
 spread1.Col = spread1.SelBlockCol
 spread1.col2 = spread1.SelBlockCol2
 spread1.BlockMode = True
 spread1.Action = SS_ACTION_DELETE_ROW
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 spread1.Row = 0
 spread1.Row2 = 0
 spread1.Col = 0
 spread1.col2 = 0
 spread1.ReDraw = True
 spread1.Refresh
 Screen.MousePointer = vbDefault
 
End Sub

Public Sub GetGridIDs(spread1 As Control, Col As Long, id$(), HMSel)
 Dim Row As Long, TempVar
 HMSel = 0
 If Not spread1.IsBlockSelected Then Exit Sub
 Screen.MousePointer = vbHourglass
 For Row = spread1.SelBlockRow To spread1.SelBlockRow2
  spread1.GetText Col, Row, TempVar
  HMSel = HMSel + 1
  id$(HMSel) = TempVar
 Next Row
 Screen.MousePointer = vbDefault
End Sub
Public Sub GetGridMultiIDs(spread1 As Control, Col() As Long, id$(), HMSel)
 Dim Row As Long, TempVar
 HMSel = 0
 If Not spread1.IsBlockSelected Then Exit Sub
 Screen.MousePointer = vbHourglass
 For Row = spread1.SelBlockRow To spread1.SelBlockRow2
  HMSel = HMSel + 1
  spread1.GetText Col(1), Row, TempVar
  id$(HMSel, 1) = TempVar
  spread1.GetText Col(2), Row, TempVar
  id$(HMSel, 2) = TempVar
 Next Row
 Screen.MousePointer = vbDefault
End Sub


Public Sub SetGridArray(theGrid As Control, column As Long, TheArray() As String, StopOnBlank As Boolean)
'**************************************************************************
' Programmer: Lynn Owens
' Date Created: 10-28-99
'
' Modified Dates:
'
' Description: this routine will return an array with values from
'              a grid in a certian column
'
' Paramerters In:
'      TheGrid: the grid that we will be adding to the array
'       Column: the column number of the grid
'   TheArray(): the array to put the valus in
'  StopOnBlank: true if we want to stop on a blank cell false if we do not
'
' Parameters Out:
'  TheArray() filled out with the data in the column
'
' Example:
'  call SetGridArray(Frmfplan.GrdSplits, 6, Grower(), true)
'
'
'**************************************************************************
 Dim Row As Long
 Dim var As Variant
 For Row = 1 To theGrid.MaxRows
  theGrid.GetText column, Row, var
  If StopOnBlank Then
    If var = "" Then Exit For
  End If
  ReDim Preserve TheArray(Row)
  TheArray(Row) = var
 Next Row
End Sub


