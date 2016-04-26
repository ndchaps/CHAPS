Attribute VB_Name = "list_boxes"
Option Explicit

Public Sub load_calf_list2(TheControl As Control, where$)
 Dim SQL$
 'On Local Error Resume Next
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select * from calfbirth "
 If where$ <> "" Then SQL$ = SQL$ & where$ & " order by calfid desc"
 Set tbData = DB.OpenRecordset(SQL$, dbOpenDynaset)
 TheControl.Clear
 While Not tbData.EOF
    If UCase$(Trim$(herdid$)) = UCase$(Trim$(tbData!herdid)) Then
      TheControl.InsertRow = Field2Str(tbData!calfid) & Chr$(9) & Field2Str(tbData!birthdate)
    End If
    tbData.MoveNext
 Wend
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 If TheControl.ListCount > 0 Then TheControl.ListIndex = 0
 Screen.MousePointer = vbDefault
End Sub

Public Sub load_cow_list(TheControl As Control, where$)
 Screen.MousePointer = vbHourglass
 Dim SQL$
 'sortcows$ = " order by cowid"
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select * from cowprof "
 If where$ <> "" Then SQL$ = SQL$ & where$
 SQL$ = SQL$ & " " & sortcows$
 Set tbData = DB.OpenRecordset(SQL$, dbOpenDynaset)
 TheControl.Clear
 While Not tbData.EOF
    If UCase$(Trim$(herdid$)) = UCase$(Trim$(tbData!herdid)) Then TheControl.AddItem Field2Str(tbData!CowID) & Chr$(9) & Field2Str(tbData!breed) & Chr$(9) & Field2Str(tbData!birthdate)
    tbData.MoveNext
 Wend
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 If TheControl.ListCount > 0 Then TheControl.ListIndex = 0
 Screen.MousePointer = vbDefault

End Sub

Public Sub load_calf_list(TheControl As Control, where$)
 Dim SQL$
 'On Local Error Resume Next
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select * from calfbirth "
 If where$ <> "" Then SQL$ = SQL$ & where$ & " order by calfid desc"
 Set tbData = DB.OpenRecordset(SQL$, dbOpenDynaset)
 TheControl.Clear
 While Not tbData.EOF
    If UCase$(Trim$(herdid$)) = UCase$(Trim$(tbData!herdid)) Then
            TheControl.AddItem Field2Str(tbData!calfid) & Chr$(9) & Field2Str(tbData!birthdate)
    End If
    tbData.MoveNext
 Wend
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 If TheControl.ListCount > 0 Then TheControl.ListIndex = 0
 Screen.MousePointer = vbDefault


 
End Sub
Public Sub load_sire_list(TheControl As Control, where$)
 Dim SQL$
 sortsires$ = " order by sireid"
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select * from sireprof "
 If where$ <> "" Then SQL$ = SQL$ & where$
 SQL$ = SQL$ & " " & sortsires$
 Set tbData = DB.OpenRecordset(SQL$, dbOpenDynaset)
 TheControl.Clear
 While Not tbData.EOF
    If UCase$(Trim$(herdid$)) = UCase$(Trim$(tbData!herdid)) Then TheControl.AddItem Field2Str(tbData!sireid) & Chr$(9) & Field2Str(tbData!breed) & Chr$(9) & Field2Str(tbData!birthdate)
    tbData.MoveNext
 Wend
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 If TheControl.ListCount > 0 Then TheControl.ListIndex = 0
 Screen.MousePointer = vbDefault
End Sub


Public Sub loadherd(TheControl As Control)
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("herd", dbOpenTable)
 TheControl.Clear
 While Not tbData.EOF
    TheControl.AddItem Field2Str(tbData!herdid) & Chr$(9) & Field2Str(tbData!herddesc)
    tbData.MoveNext
 Wend
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 If TheControl.ListCount > 0 Then TheControl.ListIndex = 0
 Screen.MousePointer = vbDefault
End Sub

