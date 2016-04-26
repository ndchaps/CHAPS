Attribute VB_Name = "CowRep"
Option Explicit
Public CullCodes(6) As String

Function Calc_Yield_Grade(HCW#, FT#, KPH#, REA#)
If HCW = 0 Or FT = 0 Or KPH = 0 Or REA = 0 Then Exit Function

Calc_Yield_Grade = 2.5 + (2.5 * FT) + (0.2 * KPH) + (0.0038 * HCW) - (0.32 * REA)

End Function

Sub Create_Lifetime_Progeny_RPTS(ReportType%, order$)
Dim SQL$, DB As DAO.database, TableName$, repdb As DAO.database
Screen.MousePointer = vbHourglass
GoSub Build_Cow_Header

Select Case ReportType
   Case 2
      
   Case 3
      
   Case 4
      'GoSub Build_Repl_DT
      GoSub Build_Back_DT
   Case 5
      'GoSub Build_Feed_DT
      GoSub Build_Repl_DT
   Case 6
      'GoSub Build_Carcass_DT
      GoSub Build_Feed_DT
   Case 7
      GoSub Build_Carcass_DT
End Select
GoSub Build_Order
Screen.MousePointer = vbDefault
Exit Sub

Build_Back_DT:
   SQL = " insert into lpr_back in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, calfbirth.CalfID, calfbirth.sex, calfback.recdate, calfback.findate, [calfback].[findate]-[calfback].[recdate] AS Days_On_Feed, calfback.recscore, calfback.recweight, calfback.finweight,  (calfback.finweight-calfback.recweight)/([calfback].[findate]-[calfback].[recdate])  AS ADG, calfback.misc1, calfback.misc2, calfback.misc3, calfbirth.sireID FROM calfbirth INNER JOIN calfback ON calfbirth.CalfID = calfback.CalfID AND calfbirth.HerdID = calfback.HerdID where calfbirth.herdid = '" & herdid & "'"
   Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   repdb.Execute "delete * from lpr_back"
   DB.Execute SQL
   SQL = "insert into lpr_back_avg SELECT DISTINCTROW LPR_Back.CowID, Sum(IIf([days_on_feed]>0,[days_on_feed],0))/Sum(IIf([days_on_feed]>0,1,0)) AS DOF, Sum(IIf([recscore]>0,[recscore],0))/Sum(IIf([recscore]>0,1,0)) AS RFS, Sum(IIf([recweight]>0,[recweight],0))/Sum(IIf([recweight]>0,1,0)) AS RW, Sum(IIf([finweight]>0,[finweight],0))/Sum(IIf([finweight]>0,1,0)) AS FW, Sum(IIf([adg]>0,[adg],0))/Sum(IIf([adg]>0,1,0)) AS Avg_ADG From LPR_Back GROUP BY LPR_Back.CowID"
   repdb.Execute "delete * from lpr_back_avg"
   repdb.Execute SQL
   repdb.Close: Set repdb = Nothing
   DB.Close: Set DB = Nothing
Return

Build_Repl_DT:
   SQL = " insert into lpr_repl in '" & repfile & "' SELECT DISTINCTROW calfbirth.cowid, calfbirth.CalfID, calfbirth.sex, calfrep.recdate, calfrep.daydate, calfrep.daydate-calfrep.recdate AS Days_On_Test, calfrep.recwt, calfrep.daywt, (calfrep.daywt-calfrep.recwt)/(calfrep.daydate-calfrep.recdate) AS ADG, calfrep.w365, calfrep.pelvic, calfrep.scrotumcir, calfrep.misc1, calfrep.misc2, calfrep.misc3, calfbirth.sireID FROM calfbirth INNER JOIN calfrep ON calfbirth.CalfID = calfrep.CalfID AND calfbirth.HerdID = calfrep.HerdID where calfbirth.herdid = '" & herdid & "' and daydate <>#01/01/1900#"
   Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   repdb.Execute "delete * from lpr_repl"
   DB.Execute SQL
   SQL = "insert into lpr_repl_avg SELECT DISTINCTROW LPR_Repl.CowID, Sum(IIf([days_on_test]>0,[days_on_test],0))/Sum(IIf([days_on_test]>0,1,0)) AS DOT, Sum(IIf([recwt]>0,[recwt],0))/Sum(IIf([recwt]>0,1,0)) AS RW, Sum(IIf([daywt]>0,[daywt],0))/Sum(IIf([daywt]>0,1,0)) AS DW, Sum(IIf([adg]>0,[adg],0))/Sum(IIf([adg]>0,1,0)) AS Avg_ADG, Sum(IIf([w365]>0,[w365],0))/Sum(IIf([w365]>0,1,0)) AS Avg_W365, Sum(IIf([pelvic]>0,[pelvic],0))/Sum(IIf([pelvic]>0,1,0)) AS PA, Sum(IIf([scrotumcir]>0,[scrotumcir],0))/Sum(IIf([scrotumcir]>0,1,0)) AS SC From LPR_Repl GROUP BY LPR_Repl.CowID"
   repdb.Execute "delete * from lpr_repl_avg"
   repdb.Execute SQL
   repdb.Close: Set repdb = Nothing
   DB.Close: Set DB = Nothing
Return

Build_Feed_DT:
   SQL = " insert into lpr_feed in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, calfbirth.sex, calffeed.calfid, calffeed.int1date, calffeed.findate, [calffeed]![findate]-[calffeed]![int1date] AS Days_On_Feed, calffeed.recscore, calffeed.int1wt, calffeed.finwt, (calffeed.finwt-calffeed.int1wt)/(calffeed!findate-calffeed!int1date) AS ADG, calffeed.misc1, calffeed.misc2, calffeed.misc3, calfbirth.sireID FROM calfbirth INNER JOIN calffeed ON calfbirth.CalfID = calffeed.CalfID AND calfbirth.HerdID = calffeed.HerdID where calfbirth.herdid = '" & herdid & "'"
   Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   repdb.Execute "delete * from lpr_feed"
   repdb.Execute "delete * from lpr_feed_avg"
   DB.Execute SQL
   SQL = "insert into lpr_feed_avg SELECT DISTINCTROW LPR_Feed.CowID, Sum(IIf([days_on_feed]>0,[days_on_feed],0))/Sum(IIf([days_on_feed]>0,1,0)) AS DOF, Sum(IIf([recscore]>0,[recscore],0))/Sum(IIf([recscore]>0,1,0)) AS RFS, Sum(IIf([int1wt]>0,[int1wt],0))/Sum(IIf([int1wt]>0,1,0)) AS RW, Sum(IIf([finwt]>0,[finwt],0))/Sum(IIf([finwt]>0,1,0)) AS FW, Sum(IIf([adg]>0,[adg],0))/Sum(IIf([adg]>0,1,0)) AS Avg_ADG From LPR_Feed GROUP BY LPR_Feed.CowID"
   repdb.Execute SQL
   repdb.Close: Set repdb = Nothing
   DB.Close: Set DB = Nothing
Return

Build_Cow_Header:
   'SQL = "insert into LPR_Header in '" & repfile & "' SELECT DISTINCTROW cowprof.cowID, Max(calfbirth.CowAge) AS CowAge, " & _
      "cowprof.breed, cowprof.sire, 100+(((Count([calfbirth].[calfid])*0.4)/(1+(Count([calfbirth].[calfid])-1)*0.4))*((Sum(IIf([calfwean].[rat" & _
      "io]>0 And [calfwean].[actweight]>0,[calfwean].[ratio],0)))/Count([calfbirth].[calfid]))) AS MPPA "
   
   SQL = "insert into LPR_Header in '" & repfile & "' SELECT DISTINCTROW cowprof.cowID, Max(calfbirth.CowAge) AS CowAge, " & _
      "cowprof.breed, cowprof.sire, cowprof.mpda AS MPPA "
   
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
      If Val(cowreps.lblMultiCows.Caption) > 0 Then
         Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
         SQL = SQL & " FROM RPTCows INNER JOIN (cowprof LEFT JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON cowprof.cowID = calfbirth.CowID) ON (RPTCows.CowID = cowprof.cowID) AND (RPTCows.HerdID = cowprof.HerdID) "
      Else
         SQL = SQL & " FROM cowprof LEFT JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON cowprof.cowID = calfbirth.CowID "
      End If
   'End If
   SQL = SQL & " where cowprof.herdid = '" & herdid & "' and cowprof.cowid <> 'Unknown' GROUP BY cowprof.cowID, cowprof.breed, cowprof.sire, cowprof.mpda"
   Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   DB.Execute "delete * from lpr_header"
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   DB.Execute SQL
   On Error Resume Next
   DB.Execute "drop table rptherd"
   On Error GoTo 0
   DB.Close: Set DB = Nothing
   Call DeleteTableAttachment(dbfile, "RPTCows")
Return

Build_Carcass_DT:
   SQL = "insert into LPR_Carc in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, calfcarcass.CalfID, calfbirth.sex, calfcarcass.carcassdate, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.yribeye, calfcarcass.qgrade, calfcarcass.qscore, calfcarcass.qcolor, calfcarcass.score, calfcarcass.misc1, calfcarcass.misc2, calfcarcass.misc3, calfbirth.sireID FROM calfbirth INNER JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID) where calfbirth.herdid = '" & herdid & "'"
   Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   repdb.Execute "delete * from lpr_carc"
   repdb.Execute "delete * from lpr_carc_avg"
   DB.Execute SQL
   SQL = "insert into lpr_carc_avg SELECT DISTINCTROW LPR_Carc.CowID, Sum(IIf([ygrade]>0,[ygrade],0))/Sum(IIf([ygrade]>0,1,0)) AS YG, Sum(IIf([ywt]>0,[ywt],0))/Sum(IIf([ywt]>0,1,0)) AS HCW, Sum(IIf([yfat]>0,[yfat],0))/Sum(IIf([yfat]>0,1,0)) AS Avg_YFat, Sum(IIf([ykidney]>0,[ykidney],0))/Sum(IIf([ykidney]>0,1,0)) AS KPH, Sum(IIf([yribeye]>0,[yribeye],0))/Sum(IIf([yribeye]>0,1,0)) AS REA, 0 AS MarbSC, Sum(IIf([score]>0,[score],0))/Sum(IIf([score]>0,1,0)) AS MusSC From LPR_Carc GROUP BY LPR_Carc.CowID"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 10.5 where qgrade = 'Prime+'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 9.5 where qgrade = 'Prime'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 8.5 where qgrade = 'Prime-'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 7.5 where qgrade = 'Choice+'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 6.75 where qgrade = 'CAB'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 6.75 where qgrade = 'STS'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 6.5 where qgrade = 'Choice'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 5.5 where qgrade = 'Choice-'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 5.5 where qgrade = 'AAA'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 4.75 where qgrade = 'Select+'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 4.5 where qgrade = 'Select'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 4.5 where qgrade = 'AA'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 4.25 where qgrade = 'Select-'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 3.5 where qgrade = 'Standard+'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 3.5 where qgrade = 'A'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = '3.0' where qgrade = 'Standard'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 2.5 where qgrade = 'Standard-'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = '1.0' where qgrade = 'B1'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'HRI'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'NoRoll'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'B2'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'B3'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'B4'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'D1'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'D2'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'D3'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'D4'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'C'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'Dark'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'Stag'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'Comm'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = 'Other'"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = ''"
   repdb.Execute SQL
   
   SQL = "update lpr_carc SET qgrade = 0 where qgrade = null "
   repdb.Execute SQL
   
   
   repdb.Close: Set repdb = Nothing
   DB.Close: Set DB = Nothing
Return

Build_Order:
   Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   SQL = "select * into tmp from lpr_header order by " & order
   DB.Execute SQL
   DB.Execute "delete * from lpr_header"
   SQL = "insert into lpr_header select * from tmp "
   DB.Execute SQL
   DB.Execute "drop table tmp"
   DB.Close: Set DB = Nothing
End Sub

Public Sub CowOrder()
Dim DB As database, SQL$
Screen.MousePointer = vbHourglass
If cowreps.optOrder(1).Value Then SQL$ = "insert into cwlst select * from cwlstsort order by cowage, cowid "
If cowreps.optOrder(0).Value Then SQL$ = "insert into cwlst select * from cwlstsort order by cowid"
If cowreps.optOrder(3).Value Then SQL$ = "insert into cwlst select * from cwlstsort order by cow_sire, cowid "
If cowreps.optOrder(2).Value Then SQL$ = "insert into cwlst select * from cwlstsort order by MPPA desc, cowid "
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
DB.Execute (SQL$)
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub

Public Sub Create_Cow_List(XAvg#, XCows&)
Dim DB As database, RS As Recordset, SQL$, order$
Dim repdb As database, RepRS As Recordset, FORMULA$, pRS As Recordset, pCowAge As Recordset
Dim tbCalfWean As DAO.Recordset
Dim CULTest$
Screen.MousePointer = vbHourglass
On Local Error GoTo ErrHandler
Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False)
Call CreateTableAttachment(dbfile, repfile, "CwLstSort", "CwLstSort")
Screen.MousePointer = vbHourglass
GoSub Build_Cow_Data
Screen.MousePointer = vbHourglass
GoSub Update_Number_Weaned
Screen.MousePointer = vbHourglass
GoSub Build_Calf_Detail
Screen.MousePointer = vbHourglass
GoSub Update_WeanCond_WeanWt
Screen.MousePointer = vbHourglass
Call CreateMPPA
Screen.MousePointer = vbHourglass
GoSub Update_Avgs
Screen.MousePointer = vbHourglass
GoSub Build_Cav_Interval
Screen.MousePointer = vbHourglass
TEXT(1) = ""
Call DeleteTableAttachment(dbfile, "CwLstSort")
pRS.Close: Set pRS = Nothing
repdb.Close: Set repdb = Nothing
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
Exit Sub

ErrHandler:
   If TEXT(1) = "Build_Cav_Interval" And Err.Number = 6 Then: Err.Clear: Resume Next
   GMODNAME$ = "CowRep.Bas -- Create_Cow_List"
   GERRNUM$ = Str$(Err.Number)
   GERRSOURCE$ = Err.Source
   MsgBox Err.Description
   Resume
   Call POP_ERROR(TEXT$())
   
Exit Sub

Build_Calf_Detail:
   TEXT(1) = "Build_Calf_Detail"
   repdb.Execute "delete * from siresumtmp"
   SQL = " INSERT INTO SireSumTmp ( herdid, CowID, CalfID, Sex, birthdate, birthwt, actweight, adj205wt, adj205rat, managecode, cframe, avgdailygain, wt2daygain, misc1, SireID, [group],dateweighed ) IN '" & repfile$ & "'"
   SQL = SQL & " SELECT DISTINCTROW CwLstSort.HerdID, CwLstSort.cowID, calfbirth.CalfID, calfbirth.sex, calfbirth.birthdate, calfbirth.birthwt, calfwean.actweight, Switch(calfbirth.sex = '0', calfwean.wt205 * 1, calfbirth.sex = '1', calfwean.wt205 * .95, calfbirth.sex = '2', calfwean.wt205 * 1.05, calfbirth.sex = '3', calfwean.wt205 * 1) as tmpwt205, calfwean.ratio, calfwean.managecode, calfwean.score, ([calfwean].[actweight]-[calfbirth].[birthwt])/([calfwean].[dateweighed]-[calfbirth].[birthdate]) AS ADG, [calfwean].[actweight]/([calfwean].[dateweighed]-[calfbirth].[birthdate]) AS WDA, calfbirth.misc1, calfbirth.sireID, calfwean.group, calfwean.dateweighed "
   SQL = SQL & " FROM (CwLstSort INNER JOIN calfbirth ON (CwLstSort.HerdID = calfbirth.HerdID) AND (CwLstSort.cowID = calfbirth.CowID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID) "
   DB.Execute SQL
   
   repdb.Execute "UPDATE DISTINCTROW SireSumTmp SET SireSumTmp.avgdailygain = 0 WHERE (((SireSumTmp.managecode)='B')) OR (((SireSumTmp.managecode)='C')) OR (((SireSumTmp.managecode)='D')) OR (((SireSumTmp.managecode)='X'))"

   
Return

Build_Cav_Interval:
   TEXT(1) = "Build_Cav_Interval"
   'cavling interval = max(birthdate) - min(birthdate) / number calves born - 1
   SQL = "SELECT DISTINCTROW CwLstSort.HerdID, CwLstSort.cowID, (Max([calfbirth].[birthdate])-Min([calfbirth].[birthdate]))/(Sum(IIf([calfwean].[managecode]='T',0.5,1))-1) AS CavInt, Max(calfbirth.birthdate) as max_bd, min(calfbirth.birthdate) as min_bd, Sum(IIf([managecode]='E',1,0)) AS E_Count FROM (CwLstSort LEFT JOIN calfbirth ON (CwLstSort.HerdID = calfbirth.HerdID) AND (CwLstSort.cowID = calfbirth.CowID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID) GROUP BY CwLstSort.HerdID, CwLstSort.cowID "
   Set pRS = DB.OpenRecordset(SQL, dbOpenSnapshot)
   Do Until pRS.EOF
      CULTest = ""
      If Field2Num(pRS!e_count) = 0 And IsDate(pRS!min_bd) And IsDate(pRS!max_bd) Then
         'On Error Resume Next
         Set tbCalfWean = DB.OpenRecordset("select calfwean.managecode from calfbirth, calfwean where calfbirth.herdid = calfwean.herdid and calfbirth.calfid = calfwean.calfid and calfbirth.herdid = '" & Field2Str(pRS!herdid) & "' and calfbirth.cowid = '" & Field2Str(pRS!CowID) & "' and calfbirth.birthdate = #" & pRS!min_bd & "#", dbOpenSnapshot)
         If Not tbCalfWean.EOF Then
            If Field2Str(tbCalfWean!managecode) = "A" Or Field2Str(tbCalfWean!managecode) = "B" Then CULTest = "CUL"
         End If
         tbCalfWean.Close: Set tbCalfWean = Nothing
         Set tbCalfWean = DB.OpenRecordset("select calfwean.managecode from calfbirth, calfwean where calfbirth.herdid = calfwean.herdid and calfbirth.calfid = calfwean.calfid and calfbirth.herdid = '" & Field2Str(pRS!herdid) & "' and calfbirth.cowid = '" & Field2Str(pRS!CowID) & "' and calfbirth.birthdate = #" & pRS!max_bd & "#", dbOpenSnapshot)
         If Not tbCalfWean.EOF Then
            If Field2Str(tbCalfWean!managecode) = "A" Or Field2Str(tbCalfWean!managecode) = "B" Then CULTest = "CUL"
         End If
         tbCalfWean.Close: Set tbCalfWean = Nothing
         On Local Error GoTo ErrHandler
      Else
         CULTest = "XXX"
      End If
      On Error Resume Next
      SQL = "update cwlstsort set cavint = " & Field2Num(pRS!cavint) & ", cavinttext = '" & CULTest & "' where cowid = '" & Field2Str(pRS!CowID) & "' and herdid = '" & Field2Str(pRS!herdid) & "' "
      If Err Then SQL = "update cwlstsort set cavint = 0, cavinttext = '" & CULTest & "' where cowid = '" & Field2Str(pRS!CowID) & "' and herdid = '" & Field2Str(pRS!herdid) & "' "
      On Error GoTo ErrHandler
      DB.Execute SQL
      pRS.MoveNext
   Loop
   
   On Local Error GoTo ErrHandler
   'xavg need to be the sum of the cavint where it is > 0 / the number of cows with a cavint > 0  not the total number of cows per Doni 1/29/04
   'SQL = "SELECT DISTINCTROW Sum(IIf([cavint]>0,[cavint],0))/Sum(IIf([cavint]>0,1,1)) AS AvgCI, Sum(IIf([cavint]>0,1,0)) AS CowCount From CwLstSort"
   SQL = "SELECT DISTINCTROW Sum(IIf([cavint]>0,[cavint],0))/Sum(IIf([cavint]>0,1,0)) AS AvgCI, Sum(IIf([cavint]>0,1,0)) AS CowCount From CwLstSort"
   Set pRS = DB.OpenRecordset(SQL, dbOpenSnapshot)
   If Not pRS.EOF Then
      If Field2Num(pRS!cowcount) <> 0 Then
        XAvg = Field2Num(pRS!avgci)
      Else
        XAvg = 0
      End If
      XCows = Field2Num(pRS!cowcount)
   End If
Return

Build_Cow_Data:
   TEXT(1) = "Build_Cow_Data"
   repdb.Execute ("delete * from CwLst")
   repdb.Execute ("delete * from CwLstSort")
   SQL$ = "insert into CwLstSort in '" & repfile$ & "' SELECT DISTINCTROW cowprof.cowID, cowprof.HerdID, count(iif(calfwean.actweight > 0, calfbirth.calfid, 0)) as brn, cowprof.breed as brd, max(calfbirth.CowAge) as cowage, cowprof.sire as cow_sire "
   If Val(cowreps.lblMultiCows) > 0 Then
      SQL = SQL & " FROM RPTCows INNER JOIN ((cowprof LEFT JOIN calfbirth ON (cowprof.HerdID = calfbirth.HerdID) AND (cowprof.cowID = calfbirth.CowID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (RPTCows.CowID = cowprof.cowID) AND (RPTCows.HerdID = cowprof.HerdID) "
      Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
   Else
      SQL = SQL & " FROM (cowprof LEFT JOIN calfbirth ON (cowprof.HerdID = calfbirth.HerdID) AND (cowprof.cowID = calfbirth.CowID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID) " & FORMULA$
   End If
   SQL = SQL & " where cowprof.herdid = '" & herdid & "' and cowprof.cowid <> 'Unknown'"
   SQL = SQL & " GROUP BY cowprof.cowID, cowprof.HerdID, cowprof.breed, cowprof.sire "
   DB.Execute (SQL$), dbFailOnError
   DBEngine.Idle dbRefreshCache
   Call DeleteTableAttachment(dbfile, "RPTCows")
Return

Update_Number_Weaned:
   TEXT(1) = "Update_Number_Weaned"
   SQL$ = "SELECT DISTINCTROW calfwean.HerdID, Sum(iif(calfwean.actweight>0,1, 0)) AS Wnd, calfbirth.CowID"
   SQL$ = SQL$ & " FROM (calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) INNER JOIN CwLstSort ON calfbirth.CowID = CwLstSort.cowID "
   SQL$ = SQL$ & " GROUP BY calfwean.HerdID, calfbirth.CowID "
   Set RS = DB.OpenRecordset(SQL$, dbOpenSnapshot)
   Do Until RS.EOF
      repdb.Execute "update cwlstsort set wnd = " & Field2Num(RS!wnd) & " where herdid = '" & Field2Str(RS!herdid) & "' and cowid = '" & Field2Str(RS!CowID) & "'"
      RS.MoveNext
   Loop
Return

Update_WeanCond_WeanWt:
   TEXT(1) = "Update_WeanCond_WeanWt"
   SQL$ = "select herdid, cowid from siresumtmp group by cowid, herdid order by cowid"
   Set pRS = repdb.OpenRecordset(SQL$, dbOpenSnapshot)
   Do Until pRS.EOF
      SQL$ = "select max(cowage) as maxcowage from calfbirth where cowid = '" & pRS!CowID & "' and herdid = '" & pRS!herdid & "'"
      Set pCowAge = DB.OpenRecordset(SQL$, dbOpenDynaset)
      SQL$ = "update siresumtmp set cowage = " & Field2Num(pCowAge!maxcowage) & " where cowid = '" & Field2Str(pRS!CowID) & "' and herdid = '" & Field2Str(pRS!herdid) & "' "
      repdb.Execute SQL$, dbFailOnError
      pCowAge.Close: Set pCowAge = Nothing
    '  SQL$ = "select * from siresumtmp where cowid = '" & Field2Str(pRS!CowID) & "' and herdid = '" & Field2Str(pRS!herdid) & "' and calfid = '" & pRS!calfid & "' order by cowid, birthdate "
    '  Set RepRS = repdb.OpenRecordset(SQL$, dbOpenDynaset)
      SQL$ = "select calfdate, weancond, weanwt,weandate from cowbrd where cowid = '" & Field2Str(pRS!CowID) & "' and herdid = '" & Field2Str(pRS!herdid) & "' order by calfdate"
      Set RS = DB.OpenRecordset(SQL$)
     '   Do Until RepRS.EOF
       Do Until RS.EOF
         If IsDate(Field2Date(RS!weandate)) Then
           SQL$ = "UPDATE DISTINCTROW SireSumTmp SET SireSumTmp.weanwt = " & Field2Num(RS!weanwt) & ", SireSumTmp.weancond = " & Field2Num(RS!weancond) & " WHERE (((SireSumTmp.herdid)='" & Field2Str(pRS!herdid) & "') AND ((SireSumTmp.CowID)='" & Field2Str(pRS!CowID) & "')) and year( siresumtmp.dateweighed) = " & Year(Field2Date(RS!weandate))
           repdb.Execute SQL$
         End If
         RS.MoveNext
         

       '     If RS.EOF Or RepRS.EOF Then Exit Do
      '      RepRS.Edit
     '       RepRS!weancond =
    '        RepRS!weanwt = Field2Num(RS!weanwt)
   '         RepRS.Update
  '          RS.MoveNext: RepRS.MoveNext
        Loop
    '  RepRS.Close: Set RepRS = Nothing
Next_Row:
      pRS.MoveNext
   Loop
   
Return

Update_Avgs:
   TEXT(1) = "Update_Avgs"
   'update averages table.
   SQL = "insert into cwlst_avg SELECT DISTINCTROW SireSumTmp.CowID, Sum(IIf([birthwt]>0,[birthwt],0))/Sum(IIf([birthwt]>0,1,0)) AS BW, Sum(IIf([actweight]>0,[actweight],0))/Sum(IIf([actweight]>0,1,0)) AS ActWt, Sum(IIf([adj205wt]>0,[adj205wt],0))/Sum(IIf([adj205wt]>0,1,0)) AS Adj205, Sum(IIf([adj205rat]>0,[adj205rat],0))/Sum(IIf([adj205rat]>0,1,0)) AS Rat, Sum(IIf([cframe]>0,[cframe],0))/Sum(IIf([cframe]>0,1,0)) AS Fr, Sum(IIf([avgdailygain]>0 and [birthwt]>0,[avgdailygain],0))/Sum(IIf([avgdailygain]>0 and [birthwt]>0,1,0)) AS ADG, Sum(IIf([wt2daygain]>0,[wt2daygain],0))/Sum(IIf([wt2daygain]>0,1,0)) AS WDA, Sum(IIf([weancond]>0,[weancond],0))/Sum(IIf([weancond]>0,1,0)) AS Cond, Sum(IIf([weanwt]>0,[weanwt],0))/Sum(IIf([weanwt]>0,1,0)) AS CowWt From SireSumTmp GROUP BY SireSumTmp.CowID"
   repdb.Execute "delete * from cwlst_avg"
   repdb.Execute SQL
Return
   
End Sub
Public Function Create_Herd_Str(Herds$)
Dim theand$, t As Integer, FORMULA$
'If cowreps.lblhow_many_herd.Caption <> "All" Then
'   If FrmSelect_Multi_Herds.lstherd.SelectedCount > 0 Then
'      theand$ = "": formula$ = ""
'         For t = 0 To FrmSelect_Multi_Herds.lstherd.ListCount - 1
'            If FrmSelect_Multi_Herds.lstherd.Tagged(t) = True Then
'               FrmSelect_Multi_Herds.lstherd.ListIndex = t
'               formula$ = formula$ & theand$ & Trim$(FrmSelect_Multi_Herds.lstherd.ColText)
'               theand$ = ", "
'            End If
'         Next t
'      Herds$ = "Herds: " & formula$
'   End If
'   Else
'      Herds$ = "All Herds"
' End If
End Function



Public Sub CreateCalfList()
'On Error Resume Next
Dim RS As Recordset, strSQL$, FORMULA$
Dim DB As database, DBREP As database
Dim SQL As String, agedays As Boolean, adj205 As Boolean, adj205r As Boolean, avgdgain As Boolean, wt2day As Boolean
Dim dam As Double, allcalf As Double, AvgAge As Double, x As Double, allwt As Double, IrrCalf As Double, y As Double
Dim wt205 As Double, birthwt As Double, cease As Double, frscor As Double, adg As Double, wdg As Double
Dim skipme As Double, theand$, t As Long

SQL = "INSERT INTO siresumtmp IN '" & repfile$ & "' SELECT DISTINCTROW calfbirth.herdid, calfbirth.CalfID, calfbirth.birthdate, calfbirth.sex, calfbirth.birthwt, calfwean.dateweighed, calfbirth.calvingease, calfwean.actweight, calfwean.managecode, calfwean.cframe, calfwean.group, calfbirth.misc1, calfbirth.CowID, calfbirth.sireID, calfbirth.CowAge, "
SQL = SQL & " cowprof.breed as cow_breed, sireprof.breed as sire_breed FROM CwLstSort INNER JOIN (sireprof RIGHT JOIN ((cowprof INNER JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID)) ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID)) ON CwLstSort.cowID = calfbirth.CowID "
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False)
DBREP.Execute ("delete * from siresumtmp")
Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
DB.Execute (SQL$)
DB.Close: Set DB = Nothing
Set RS = DBREP.OpenRecordset("siresumtmp", dbOpenTable)
If RS.RecordCount > 0 Then
Do Until RS.EOF
   GoSub agedays
   GoSub adj205
   GoSub avgdgain
   GoSub wt2day
   If RS!skipme = False And RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
      x = x + RS!age_in_days
   Else
      skipme = skipme + 1
   End If
   RS.MoveNext
Loop
If RS.RecordCount <> skipme Then
   AvgAge = x / (RS.RecordCount - skipme)
End If
x = 0
IrrCalf = 0
RS.MoveFirst
Do Until RS.EOF
   If RS!age_in_days > 0 And RS!age_in_days >= AvgAge - 45 And RS!age_in_days <= AvgAge + 45 Then
            If RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
               x = x + RS!age_in_days
               allcalf = allcalf + 1
               allwt = allwt + RS!adj205wt
            Else
               RS.Edit: RS!skipme = True: RS.Update
            End If
         Else
               RS.Edit: RS!skipme = True: RS.Update
         End If
      'End If
   RS.MoveNext
Loop
RS.MoveFirst
Do Until RS.EOF
   GoSub adj205r
   RS.MoveNext
Loop

RS.Close: Set RS = Nothing
DBREP.Close: Set DBREP = Nothing

End If


Exit Sub

agedays:
 With RS
   .Edit
   agedays = True
   If RS!skipme = True Or IsNull(RS!dateweighed) Or RS!dateweighed = "" Or RS!managecode = "A" Or RS!managecode = "B" Or RS!managecode = "F" Or RS!birthdate = #1/1/1900# Then
      RS!age_in_days = -9999
      RS!skipme = True
      agedays = False
   End If
   If agedays = True Then
      RS!age_in_days = Abs(RS!dateweighed - RS!birthdate)
   End If
   .Update
 End With
Return

adj205:
   adj205 = True
   dam = 0
   RS.Edit
   If RS!skipme = True Or RS!actweight = "" Or RS!actweight = 0 Or RS!managecode = "F" Or RS!managecode = "X" Then
      RS!adj205wt = -9999
      RS!skipme = True
      adj205 = False
   End If
   If Not IsDate(Format(RS!dateweighed, "mm/dd/yyyy")) Or RS!dateweighed = Null Or RS!dateweighed = "" Then
      RS!adj205wt = -9999
      RS!skipme = True
      adj205 = False
   End If
   If adj205 = True Then
      If RS!birthwt = "" Or IsNull(RS!birthwt) Then
         If RS!Sex = 1 Or RS!Sex = 3 Then
            RS!birthwt = 75
         End If
         If RS!Sex = 2 Then
            RS!birthwt = 70
         End If
      End If
      If RS!Sex = 2 Then
         If RS!cowage = 2 Then
            dam = 54
         End If
         If RS!cowage = 3 Then
            dam = 36
         End If
      If RS!cowage = 4 Then
         dam = 18
      End If
      If RS!cowage > 11 Then
         dam = 18
      End If
   End If
   If RS!Sex = 1 Or RS!Sex = 3 Then
      If RS!cowage = 2 Then
         dam = 60
      End If
      If RS!cowage = 3 Then
         dam = 40
      End If
      If RS!cowage = 4 Then
         dam = 20
      End If
      If RS!cowage > 11 Then
         dam = 20
      End If
   End If
   RS!adj205wt = (((RS!actweight - RS!birthwt) / RS!age_in_days) * 205) + RS!birthwt + dam
   End If
   RS.Update
Return

adj205r:
   adj205r = True
   RS.Edit
   If RS!skipme = True Or RS!managecode >= "A" And RS!managecode <= "F" Or RS!managecode = "K" Or RS!managecode = "N" Or RS!managecode = "P" Or RS!managecode = "S" Or RS!managecode = "T" Or RS!managecode = "X" Then
      RS!adj205rat = -9999
      RS!skipme = True
      adj205r = False
   End If
   If adj205r = True Then
      RS!adj205rat = (RS!adj205wt / (allwt / allcalf)) * 100
   End If
   RS.Update
Return

avgdgain:
   avgdgain = True
   RS.Edit
   If RS!skipme = True Or RS!age_in_days <= 0 Or RS!managecode = "F" Or RS!birthwt <= 0 Or RS!actweight - RS!birthwt <= 0 Then
      RS!avgdailygain = -9999
      RS!skipme = True
      avgdgain = False
   End If
   If avgdgain = True Then
      RS!avgdailygain = (RS!actweight - RS!birthwt) / RS!age_in_days
   End If
   RS.Update
Return

wt2day:
   wt2day = True
   RS.Edit
   If RS!skipme = True Or RS!age_in_days < 0 Then
     RS!wt2daygain = -9999
     RS!skipme = True
     wt2day = False
     GoTo AgeTestDone
   End If
   If RS!actweight / RS!age_in_days <= 0 Or RS!managecode = "X" Or RS!managecode = "F" Then
      RS!wt2daygain = -9999
      RS!skipme = True
      wt2day = False
   End If
AgeTestDone:
   If wt2day = True Then
      RS!wt2daygain = RS!actweight / RS!age_in_days
   End If
   RS.Update
Return
End Sub



Private Sub CreateFormula(FORMULA$)
Dim theand$, t As Integer
If cowreps.lblMultiCows.Caption <> "All" Then
   If FrmSelect_Multi_Cows.lstCows.SelectedCount > 0 Then
      theand$ = "": FORMULA$ = ""
         For t = 0 To FrmSelect_Multi_Cows.lstCows.ListCount - 1
            If FrmSelect_Multi_Cows.lstCows.Tagged(t) = True Then
               FrmSelect_Multi_Cows.lstCows.ListIndex = t
               FORMULA$ = FORMULA$ & theand$ & " cowprof.cowid = '" & FrmSelect_Multi_Cows.lstCows.ColText & "'"
               theand$ = " or "
            End If
         Next t
   FORMULA$ = " where " & FORMULA$
   End If
 End If
End Sub

Public Sub CreateMPPA()
Dim DB As database, RS As Recordset, SQL$, FORMULA$, theand$, t As Integer
Dim repdb As database, RepRS As Recordset
SQL$ = "SELECT DISTINCTROW calfwean.HerdID, Count(calfwean.CalfID) AS CountOfCalfID, Sum(calfwean.ratio) AS Sum, calfbirth.CowID"
SQL$ = SQL$ & " FROM calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)"
If cowreps.lblMultiCows.Caption <> "All" Then
   If FrmSelect_Multi_Cows.lstCows.SelectedCount > 0 Then
      theand$ = " or ": FORMULA$ = ""
         For t = 0 To FrmSelect_Multi_Cows.lstCows.ListCount - 1
            If FrmSelect_Multi_Cows.lstCows.Tagged(t) = True Then
               FrmSelect_Multi_Cows.lstCows.ListIndex = t
               FORMULA$ = FORMULA$ & theand$ & " (((calfbirth.cowid) = '" & FrmSelect_Multi_Cows.lstCows.ColText & "'))"
            End If
         Next t
   End If
 End If
SQL$ = SQL$ & " where (((calfwean.actweight) > 0) And ((calfwean.Ratio) > 0))"
SQL$ = SQL$ & FORMULA
SQL$ = SQL$ & " GROUP BY calfwean.HerdID, calfbirth.CowID"
SQL$ = SQL$ & " ORDER BY calfbirth.CowID"
Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DB.OpenRecordset(SQL$, dbOpenSnapshot)
Set RepRS = repdb.OpenRecordset("cwlstsort", dbOpenTable)
With RS
   If Not RS.EOF Then .MoveFirst
   Do Until .EOF
      repdb.Execute "update cwlstsort set mppa = 100 + ((" & Field2Num(RS!countofcalfid) & " * 0.4) / (1 + (" & Field2Num(RS!countofcalfid) & " - 1) * 0.4)) * ((" & Field2Num(RS!Sum) & " / " & Field2Num(RS!countofcalfid) & ") - 100) where cowid = '" & Field2Str(!CowID) & "' and herdid = '" & Field2Str(!herdid) & "'"
      .MoveNext
   Loop
End With
RS.Close: Set RS = Nothing
SQL = "UPDATE DISTINCTROW CwLstSort INNER JOIN cowprof ON (CwLstSort.cowID = cowprof.cowID) AND (CwLstSort.HerdID = cowprof.HerdID) SET cowprof.mpda = [cwlstsort].[mppa]"
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub
Public Sub GetCowCndWnWt()
Dim DB As database, repdb As database, RS As Recordset, RepRS As Recordset
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
Set RepRS = repdb.OpenRecordset("select * from siresumtmp order by cowid", dbOpenDynaset)
Set RS = DB.OpenRecordset("cowbrd", dbOpenTable)
RS.Index = "primarykey"
'RS.MoveFirst: RepRS.MoveLast: RepRS.MoveFirst
Do Until RepRS.EOF
   RS.Seek "=", RepRS!herdid, RepRS!CowID, RepRS!dateweighed
   If Not RS.NoMatch Then
      RepRS.Edit: RepRS!weanwt = RS!weanwt: RepRS!weancond = RS!weancond: RepRS.Update
   End If
   RepRS.MoveNext
Loop
repdb.Close: Set repdb = Nothing
DB.Close: Set DB = Nothing
End Sub

Public Sub LoadCows(mhListBox As Mh3dList, pType$)
Dim DB As database, RS As Recordset, strSQL$
Screen.MousePointer = vbHourglass
If pType = "Active" Then pType = " and cowprof.active = 'A' "
If pType = "Culled" Then pType = " and cowprof.active = 'C' "
If pType = "Pedigree" Then pType = " and cowprof.active = 'P' "
 strSQL$ = "select cowid, herdid from cowprof where herdid = '" & herdid & "'" & pType
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset(strSQL$)
 mhListBox.Clear
 Do Until tbData.EOF
    mhListBox.AddItem Field2Str(tbData!CowID) & Chr$(9) & Field2Str(tbData!herdid)
    tbData.MoveNext
 Loop
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Sub Build_Cow_RPT_List(mhListBox As Mh3dList)
Dim DB As DAO.database, indx%
Dim pCowID$, pHerdID$
Set DB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from rptcows"
Do Until indx = mhListBox.ListCount
   If mhListBox.Tagged(indx) Then
      mhListBox.Col = 0
      pCowID = mhListBox.ColList(indx)
      mhListBox.Col = 1
      pHerdID = mhListBox.ColList(indx)
      DB.Execute "insert into rptcows (herdid, cowid) values ('" & pHerdID$ & "', '" & pCowID & "')"
   End If
   indx = indx + 1
Loop
DB.Close: Set DB = Nothing
End Sub
