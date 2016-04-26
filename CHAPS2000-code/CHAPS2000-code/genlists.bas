Attribute VB_Name = "Gen_Lists"
Option Explicit
 Dim tbgrower As Recordset
 Dim sqlstmt$

Public Sub Load_Customer_List(thecontrol As Control, letter$)
' This Routine Loads A List Box With Clients
'
' TheControl = The Control Name For The List Box ex. frmselect_Clients!lsitclient
' letter$    = The First Lettter Of The Last Name To Populate The List Box With Ex. A Would Populate The List Box With Last Names That Start With A
'
 Dim tbgrower As Recordset
 If Len(letter$) = 0 Then Exit Sub
 Screen.MousePointer = vbHourglass
 Set db = DBEngine(0).OpenDatabase(dbfile$, False, False)
 sqlstmt$ = "select * from GROWER where GROWname2 like '" & letter$ & "*'"
 Set tbgrower = db.OpenRecordset(sqlstmt$, dbOpenSnapshot)
 thecontrol.Clear
 While Not tbgrower.EOF
  thecontrol.AddItem Field2Str(tbgrower!growname2) & Chr$(9) & Field2Str(tbgrower!growname1) & Chr$(9) & Field2Str(tbgrower!growid)
  tbgrower.MoveNext
 Wend
 tbgrower.Close: Set tbgrower = Nothing
 db.Close: Set db = Nothing
 If thecontrol.ListCount > 0 Then thecontrol.ListIndex = 0
 Screen.MousePointer = vbDefault
End Sub

