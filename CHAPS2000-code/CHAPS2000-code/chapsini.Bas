Attribute VB_Name = "ini_read"
Option Explicit
Public DBNAMELEN%
Public datadrive$
Public dbdir$
Public DbName$
Public HelpFile$
Public calfhead(20) As String
Public sortcows$
Public sortsires$

'fix add for print crystal class
Public APP_COMPANY_name, CURRENT_USER_NAME$
Public exclusiveyn%, readonlyyn%

' fix ADDED FOR POP ERROR
Public current_USER$, APP_COMPANY_phone


Public Sub read_ini()
  Screen.MousePointer = vbHourglass
 exclusiveyn% = False
 readonlyyn% = False
 '
 ' Database Name
 '
 dbfile$ = Space$(80)
 DBNAMELEN% = GetPrivateProfileString("DataSource", "DBName", "", dbfile$, Len(dbfile$), "chaps.ini")
 DbName$ = Left$(dbfile$, DBNAMELEN%)
 '
 ' Database Dir
 '
 dbdir$ = Space$(80)
 DBNAMELEN% = GetPrivateProfileString("DataSource", "DBdir", "", dbdir$, Len(dbdir$), "chaps.ini")
 dbdir$ = Left$(dbdir$, DBNAMELEN%)
 '
 ' Database Drive
 '
 If Mid$(dbdir$, 2, 1) = ":" Then
   datadrive$ = Left$(dbdir$, 3)
  Else
   datadrive$ = Left$(CurDir$, 3)
 End If
 '
 ' Set The Database Name Used In The Program
 '
 dbfile$ = dbdir$ & DbName$
 repfile$ = dbdir$ & "chapreps.mdb"
' helpfile$ = DBDIR$ + "agvance.hlp"
 HelpFile$ = App.Path & "\chaps.hlp"
 App.HelpFile = HelpFile$
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 herdid = ReadDefaultHerdID
 If herdid <> "" Then
   mdimain.Caption = "Default Herd: " & herdid & "   C.H.A.P.S. (Herd Analysis)"
 Else
   mdimain.Caption = "Default Herd Not Set   C.H.A.P.S. (Herd Analysis)"
 End If
 Set tbData = DB.OpenRecordset("select * from prefer where herdid = '" & herdid & "'", dbOpenSnapshot)
 If Not tbData.EOF Then
    '  sought and found the data now put the data to the controls on the form
   epdhead1$ = Field2Str(tbData!epdhead1)
   epdhead2$ = Field2Str(tbData!epdhead2)
   epdhead3$ = Field2Str(tbData!epdhead3)
   epdhead4$ = Field2Str(tbData!epdhead4)
   epdhead5$ = Field2Str(tbData!epdhead5)
   epdhead6$ = Field2Str(tbData!epdhead6)
   epdhead7$ = Field2Str(tbData!epdhead7)
   epdhead8$ = Field2Str(tbData!epdhead8)
   epdhead9$ = Field2Str(tbData!epdhead9)
   epdhead10$ = Field2Str(tbData!epdhead10)
   calfhead(1) = Field2Str(tbData!prof1)
   calfhead(2) = Field2Str(tbData!prof2)
   calfhead(3) = Field2Str(tbData!prof3)
   calfhead(4) = Field2Str(tbData!wean1)
   calfhead(5) = Field2Str(tbData!wean2)
   calfhead(6) = Field2Str(tbData!wean3)
   calfhead(7) = Field2Str(tbData!background1)
   calfhead(8) = Field2Str(tbData!background2)
   calfhead(9) = Field2Str(tbData!background3)
   calfhead(10) = Field2Str(tbData!replacement1)
   calfhead(11) = Field2Str(tbData!replacement2)
   calfhead(12) = Field2Str(tbData!replacement3)
   calfhead(13) = Field2Str(tbData!feedlot1)
   calfhead(14) = Field2Str(tbData!feedlot2)
   calfhead(15) = Field2Str(tbData!feedlot3)
   calfhead(16) = Field2Str(tbData!carcass1)
   calfhead(17) = Field2Str(tbData!carcass2)
   calfhead(18) = Field2Str(tbData!carcass3)
   Select Case tbData!cowsort
    Case 0
     sortcows$ = "order by cowprof.cowid"
    Case 1
     sortcows$ = "order by cowprof.Birthdate, cowprof.cowid, cowprof.breed"
    Case 2
     sortcows$ = "order by cowprof.Birthdate,cowprof.breed, cowprof.cowid"
    Case 3
     sortcows$ = "order by cowprof.breed, cowprof.cowid, cowprof.Birthdate"
    Case 4
     sortcows$ = "order by cowprof.breed, cowprof.Birthdate, cowprof.cowid"
   End Select
   Select Case tbData!siresort
    Case 0
     sortsires$ = "order by sireprof.sireid"
    Case 1
     sortsires$ = "order by sireprof.Birthdate, sireprof.sireid, sireprof.breed"
    Case 2
     sortsires$ = "order by sireprof.Birthdate, sireprof.breed, sireprof.sireid"
    Case 3
     sortsires$ = "order by sireprof.breed, sireprof.sireid, sireprof.Birthdate"
    Case 4
     sortsires$ = "order by sireprof.breed, sireprof.Birthdate, sireprof.sireid"
   End Select
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
 
End Sub

