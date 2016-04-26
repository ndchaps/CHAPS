Attribute VB_Name = "gen_stuff"
Option Explicit

Public DB As database
Public dbfile$
Public repfile$
Public tbData
Public herdid$
Public epdhead1$
Public epdhead2$
Public epdhead3$
Public epdhead4$
Public epdhead5$
Public epdhead6$
Public epdhead7$
Public epdhead8$
Public epdhead9$
Public epdhead10$
Public clfhead1$
Public clfhead2$
Public clfhead3$
Public clfhead4$
Public clfhead5$
Public clfhead6$
Public clfback1$, clfback2$, clfback3$, clfref1$, clfref2$, clfref3$, clffeed1$, clffeed2$, clffeed3$, clfcarc1$, clfcarc2$, clfcarc3$

'Declare Function SendMessage Lib "USER" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long


 'for the print class
 Public report As New clsPrintCrystal

Public gcaldate

'FOR FRMGETDISK
Public DISKCANCEL%, DISKDRIVE$, diskcaption$


Public Function BuildChapsPassword() As String
 Dim DayName$(7), dow%
 DayName$(1) = "ns"
 DayName$(2) = "nm"
 DayName$(3) = "et"
 DayName$(4) = "dw"
 DayName$(5) = "ut"
 DayName$(6) = "if"
 DayName$(7) = "ts"
 dow% = Weekday(Date)
 BuildChapsPassword = "Chaps" & Trim$(Str$(Val(Left$(Date$, 2)) * Val(Mid$(Date$, 4, 2)))) & DayName$(dow%)
End Function
Sub Build_RPT_Herd(FrmH As FrmSelect_Multi_Herds)
Dim pDB As DAO.database, indx%
Set pDB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
Do Until indx = FrmH.lstherd.ListCount
   If FrmH.lstherd.Tagged(indx) Then
      pDB.Execute "insert into RPTHerd (HerdID) Values ('" & FrmH.lstherd.ColList(indx) & "')"
   End If
   indx = indx + 1
Loop
pDB.Close: Set pDB = Nothing
End Sub

Public Function NumericOnly(KeyCode As Integer) As Integer
 'use this function in the keypress event of text boxes to force all keystrokes
 'to be numeric. it basically filters the keystroke and changes ity
 'to a zero (no keypressed at all). if it is non numeric. any nonzero
 'value returned for NumericOnly indicates that the keypress was valid and
 'will be accepted by the keypress event.
 
 If KeyCode = 8 Then NumericOnly = KeyCode: Exit Function  'backspace
 If KeyCode = 46 Then NumericOnly = KeyCode: Exit Function 'dec point
 If KeyCode = 45 Then NumericOnly = KeyCode: Exit Function 'neg sign (dash)
 
 If KeyCode < 48 Or KeyCode > 57 Then
   NumericOnly = 0
  Else
   NumericOnly = KeyCode
 End If
End Function

Function ReturnBullTurnOutDate(pHerdID$, OLDDATE() As String) As String
On Local Error GoTo ErrHandler
Dim pDB As database, pRS As Recordset, pOldDate$(5)
Set pDB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set pRS = pDB.OpenRecordset("TurnOutDate", dbOpenTable)
pRS.Index = "primarykey"
pRS.Seek "=", pHerdID
If pRS.NoMatch Then GoTo Close_DB
ReDim OLDDATE$(5)
ReturnBullTurnOutDate = Field2Date(pRS!currentdate)
OLDDATE(0) = Field2Date(pRS!date1)
OLDDATE(1) = Field2Date(pRS!date2)
OLDDATE(2) = Field2Date(pRS!date3)
OLDDATE(3) = Field2Date(pRS!date4)
OLDDATE(4) = Field2Date(pRS!date5)
OLDDATE(5) = Field2Date(pRS!date6)
Close_DB:
pRS.Close: Set pRS = Nothing
pDB.Close: Set pDB = Nothing
Exit Function
ErrHandler:
TEXT$(1) = ""
TEXT$(2) = ""
TEXT$(3) = ""
TEXT$(4) = ""
TEXT$(5) = ""
GMODNAME$ = "Gen_Stuff - ReturnBullTurnOutDate"
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())
End Function

Sub SaveBullTurnOutDate(NewDate$, pHerdID$, AddEditFlag$)
Dim pDB As database, pRS As Recordset, LastDate$
On Local Error GoTo ErrHandler
Set pDB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set pRS = pDB.OpenRecordset("TurnOutDate", dbOpenTable)
pRS.Index = "primarykey"
pRS.Seek "=", pHerdID
If pRS.NoMatch Then
    With pRS
      .AddNew
      !herdid = pHerdID
      !currentdate = CDate(NewDate)
      .Update
   End With
Else
   If AddEditFlag = "A" Then
      With pRS
         .Edit
         !date6 = !date5
         !date5 = !date4
         !date4 = !date3
         !date3 = !date2
         !date2 = !date1
         !date1 = !currentdate
         !currentdate = CDate(NewDate)
         .Update
      End With
   Else
      pRS.Edit
      pRS!currentdate = NewDate
      pRS.Update
   End If
End If
pRS.Close: Set pRS = Nothing
pDB.Close: Set pDB = Nothing
Exit Sub
ErrHandler:
TEXT$(1) = ""
TEXT$(2) = ""
TEXT$(3) = ""
TEXT$(4) = ""
TEXT$(5) = ""
GMODNAME$ = "Gen_Stuff - SaveBullTurnOutDate"
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())
End Sub

Sub set_combo(theCombo As Control, TheData As String)
 Dim lpos As Long
 Dim loopcnt As Long
 If TheData$ = "" Then Exit Sub
' lpos = SendMessage(thecombo.hWnd, CB_FINDSTRING, 0, ByVal thedata)
' lpos = SendMessage(thecombo.hWnd, CB_FINDSTRING, 0, ByVal thedata)
 lpos = -1
 For loopcnt = 0 To theCombo.ListCount - 1
   If Len(theCombo.list(loopcnt)) = Len(TheData$) Then
      If Trim$(Left$(theCombo.list(loopcnt), Len(TheData$))) = TheData Then lpos = loopcnt: Exit For
   End If
 Next loopcnt
 If lpos >= 0 Then
  theCombo.ListIndex = lpos
 End If

End Sub

Function ReadDefaultHerdID() As String
Dim ret%
ReadDefaultHerdID = String(80, vbNullChar)
ret = GetPrivateProfileString("chaps", "DefaultHerd", "", ReadDefaultHerdID, Len(ReadDefaultHerdID), "chaps.ini")
ReadDefaultHerdID = Left$(ReadDefaultHerdID, ret)
End Function

Public Sub UpdateDataBase()
 
 Dim DataBaseUpToDate As Boolean
 Dim DB As database
 Dim dusers As Recordset
 Dim Field(5) As String
 Dim Field2(5) As String
 Dim PropertyArray(5) As String
 Dim exitcode%
 Dim message As String
 
 On Local Error GoTo LeHandle

 DataBaseUpToDate = True
1 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  'StartTime = Timer
2 'Set fld = DB.TableDefs("salesmen").Fields("inactivedate")
  Set dusers = DB.OpenRecordset("SELECT Query.Name From Query", dbOpenSnapshot)
  'EndTime = Timer
  dusers.Close: Set dusers = Nothing
3 DB.Close: Set DB = Nothing
  'MsgBox "It Took: " & Format$(EndTime - StartTime, "#####.########")
4 If DataBaseUpToDate Then
   Exit Sub
  End If




 '
 ' add a new table to the database
 '
 PropertyArray(0) = "Text"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "0"
 PropertyArray(3) = "80"
 Call Update_Database(dbfile$, "Query", "Name", "", "", "", Field(), Field(), PropertyArray(), 11, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3010 Then
   message = message & "Add Table Query and Field Name: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
 
 PropertyArray(0) = "Text"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "0"
 PropertyArray(3) = "1"
 Call Update_Database(dbfile$, "Query", "QueryType", "", "", "", Field(), Field(), PropertyArray(), 1, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3191 Then
   message = message & "Add Field QueryType to Query Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
 
 PropertyArray(0) = "Memo"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "0"
 PropertyArray(3) = "10"
 Call Update_Database(dbfile$, "Query", "SQL", "", "", "", Field(), Field(), PropertyArray(), 1, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3191 Then
   message = message & "Add Field SQL to Query Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
 

 
 
' Setup the primarykey
 ' Change the unique property of the field CrossRefID in the Product to false
 '  Dups ok
 PropertyArray(0) = "-1"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "-1"
 PropertyArray(3) = "0"
 Field(0) = "1"
 Field(1) = "Name"
 Call Update_Database(dbfile$, "Query", "", "PrimaryKey", "", "", Field(), Field(), PropertyArray(), 4, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3283 Then
   message = message & "Add a Primary Key to Query Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
Exit Sub
LeHandle:
 If Erl = 2 Then
   If Err.Number = 3265 Then
     DataBaseUpToDate = False
     Resume Next
   End If
   If Err.Number = 3078 Then
     DataBaseUpToDate = False
     Resume 3
   End If
   

   If Err.Number = 3061 Then
    DataBaseUpToDate = False
    Resume 3
   End If
 End If
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "MakeAgvanceDatabaseCurrent - MakeDatabaseCurrent"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())


End Sub

Public Sub UpdateDataBase2()
 
 Dim DataBaseUpToDate As Boolean
 Dim DB As database
 Dim dusers As Recordset
 Dim Field(5) As String
 Dim Field2(5) As String
 Dim PropertyArray(5) As String
 Dim exitcode%
 Dim message As String
 
 On Local Error GoTo LeHandle

 DataBaseUpToDate = True
1 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  'StartTime = Timer
2 'Set fld = DB.TableDefs("salesmen").Fields("inactivedate")
  Set dusers = DB.OpenRecordset("SELECT eidlist.eid From eidlist", dbOpenSnapshot)
  'EndTime = Timer
  dusers.Close: Set dusers = Nothing
3 DB.Close: Set DB = Nothing
  'MsgBox "It Took: " & Format$(EndTime - StartTime, "#####.########")
4 If DataBaseUpToDate Then
   Exit Sub
  End If




 '
 ' add a new table to the database
 '
 PropertyArray(0) = "Text"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "0"
 PropertyArray(3) = "20"
 Call Update_Database(dbfile$, "eidlist", "EID", "", "", "", Field(), Field(), PropertyArray(), 11, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3010 Then
   message = message & "Added Table EIDList and Field Name: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
 
' PropertyArray(0) = "Text"
' PropertyArray(1) = "-1"
' PropertyArray(2) = "0"
' PropertyArray(3) = "1"
' Call Update_Database(dbfile$, "Query", "QueryType", "", "", "", Field(), Field(), PropertyArray(), 1, exitcode%, "")
' If exitcode% <> 0 And exitcode <> 3191 Then
'   message = message & "Add Field QueryType to Query Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
' End If
'
' PropertyArray(0) = "Memo"
' PropertyArray(1) = "-1"
' PropertyArray(2) = "0"
' PropertyArray(3) = "10"
' Call Update_Database(dbfile$, "Query", "SQL", "", "", "", Field(), Field(), PropertyArray(), 1, exitcode%, "")
' If exitcode% <> 0 And exitcode <> 3191 Then
'   message = message & "Add Field SQL to Query Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
' End If
'

 
 
' Setup the primarykey
 ' Change the unique property of the field CrossRefID in the Product to false
 '  Dups ok
 PropertyArray(0) = "-1"
 PropertyArray(1) = "-1"
 PropertyArray(2) = "-1"
 PropertyArray(3) = "0"
 Field(0) = "1"
 Field(1) = "eid"
 Call Update_Database(dbfile$, "EIDList", "", "PrimaryKey", "", "", Field(), Field(), PropertyArray(), 4, exitcode%, "")
 If exitcode% <> 0 And exitcode <> 3283 Then
   message = message & "Add a Primary Key to EIDList Table: " & exitcode% & " " & Error(exitcode%) & vbCrLf
 End If
Exit Sub
LeHandle:
 If Erl = 2 Then
   If Err.Number = 3265 Then
     DataBaseUpToDate = False
     Resume Next
   End If
   If Err.Number = 3078 Then
     DataBaseUpToDate = False
     Resume 3
   End If
   

   If Err.Number = 3061 Then
    DataBaseUpToDate = False
    Resume 3
   End If
 End If
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "MakeAgvanceDatabaseCurrent - MakeDatabaseCurrent"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())


End Sub



Sub WriteDefaultHerdID(sHerdID As String)
Dim tmp%
 tmp% = WritePrivateProfileString("chaps", "DefaultHerd", sHerdID, "chaps.ini")
 End Sub
 
Sub SetupDirectoryForm(bChangeDir As Boolean)
Load frmchangedir
If bChangeDir = True Then
   frmchangedir.LBLDir.Visible = False
   frmchangedir.txtDirName.Visible = False
   frmchangedir.cmdchangedir(0).Top = 1920
   frmchangedir.cmdchangedir(1).Top = 1920
   frmchangedir.Height = 2805
Else
   frmchangedir.LBLDir.Visible = True
   frmchangedir.txtDirName.Visible = True
   frmchangedir.cmdchangedir(0).Top = 2445
   frmchangedir.cmdchangedir(1).Top = 2445
   frmchangedir.Height = 3300
End If
End Sub
