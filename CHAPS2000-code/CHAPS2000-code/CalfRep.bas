
Attribute VB_Name = "CalfRep"
Option Explicit
Dim F As Double, P As Double 'Total for F managecodes, used in createrepperf
Dim TCE As Double 'Total for total cows kept for calving used in replacement calf calculation
Public TurnDate As Date 'Bull Turn Out Date
Public thirdcowdate As Date, ThirdCow As String 'Third cow string and date
Public denom As Double 'Denominator for Pregper, Calving Percentage, Calf Death Loss, Calf Crop or Wean Per, Replacement Rate
Dim SortOrder% 'Integer for Order By Clause Select Statement
Public DataFound(3) As Boolean
Public EndSub As Boolean 'No calves found
Public ValidTurnDate As Boolean
Public group As Boolean
Private Calculated As Boolean
Private BoolThirdCowDate As Boolean
Private ActTurnDate As Date
Private mTestDate As Date

Public Function Check_EID(eid As String, recordtype As String, orgid As String, herdid As String, calfidatbirth As String, rectype As String) As String
    Dim SQL As String
    Dim dbReps As database
    Dim rsReps As Recordset
    Set dbReps = DBEngine(0).OpenDatabase(dbfile$, False, False)

    Check_EID = True
    If eid = "" Then Exit Function
      
      'check the calfs
      SQL = "SELECT calfid, herdid FROM calfbirth where calfbirth.elecid = '" & eid & "' "
      If calfidatbirth <> "" Then
        SQL = SQL & " and calfid <> '" & calfidatbirth & "'"
      End If
            Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      While Not rsReps.EOF
         If recordtype = "Calf" Then
           If UCase(rsReps!calfid) = UCase(orgid) And UCase(rsReps!herdid) = UCase(herdid) Then
           Else
             Check_EID = False
             rsReps.Close: Set rsReps = Nothing
             dbReps.Close: Set dbReps = Nothing
             Exit Function
           End If
         Else
           Check_EID = False
           rsReps.Close: Set rsReps = Nothing
           dbReps.Close: Set dbReps = Nothing
           Exit Function
         End If
         rsReps.MoveNext
      Wend
      
      
      'check the cows
      SQL = "SELECT cowid, herdid FROM cowprof where cowprof.elecid = '" & eid & "' "
      If rectype = "Calf" Then
        SQL = SQL & " and calfid  <> '" & orgid & "'"
      End If
      Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      While Not rsReps.EOF
         If recordtype = "Cow" Then
           If UCase(rsReps!CowID) = UCase(orgid) And UCase(rsReps!herdid) = UCase(herdid) Then
           Else
             Check_EID = False
             rsReps.Close: Set rsReps = Nothing
             dbReps.Close: Set dbReps = Nothing
             Exit Function
           End If
         Else
           Check_EID = False
           rsReps.Close: Set rsReps = Nothing
           dbReps.Close: Set dbReps = Nothing
           Exit Function
         End If
         rsReps.MoveNext
      Wend
      
      
      
      'check the sires
      SQL = "SELECT Sireid, herdid FROM sireprof where sireprof.elecid = '" & eid & "' "
      If rectype = "Calf" Then
        SQL = SQL & " and calfid  <> '" & orgid & "'"
      End If
            Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      While Not rsReps.EOF
         If recordtype = "Sire" Then
           If UCase(rsReps!sireid) = UCase(orgid) And UCase(rsReps!herdid) = UCase(herdid) Then
           Else
             Check_EID = False
             rsReps.Close: Set rsReps = Nothing
             dbReps.Close: Set dbReps = Nothing
             Exit Function
           End If
         Else
           Check_EID = False
           rsReps.Close: Set rsReps = Nothing
           dbReps.Close: Set dbReps = Nothing
           Exit Function
         End If
         rsReps.MoveNext
        End

        rsReps.Close() : rsReps = Nothing
        dbReps.Close() : dbReps = Nothing


End Function


Public Sub CreateCalfReps(SortOrder%, BeginDate As Date, enddate As Date, Overwrite As Boolean)
Dim SQL$, dbChaps As database, dbReps As database, Sex As Single, mTotalWeight As Double, mTotalCalf As Double, mUncounted As Double, mUncountedWt As Double
Dim rsReps As Recordset, AvgAge As Double, RepRSEdit As Recordset, RepRSData As Recordset
Dim age As Double, rowcount As Double, avgwt As Double, allcalf As Double, order$, iResponse As Integer
'On Error Resume Next
EndSub = False
Call FindTurnDate
If EndSub Then Exit Sub
Set dbChaps = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set dbReps = DBEngine(0).OpenDatabase(repfile$, False, False)
dbReps.Execute ("delete * from calflist"), dbFailOnError
dbReps.Execute ("delete * from calflistavg")
dbReps.Execute ("delete * from calflistsum")

For Sex = 0 To 3 'iterate through 1 time for each sex
   SQL$ = "insert into calflist in '" & repfile$ & "'"
   SQL$ = SQL$ & " SELECT DISTINCTROW calfbirth.sex, calfbirth.HerdID, calfbirth.CalfID, calfbirth.birthdate, calfwean.dateweighed, calfbirth.birthwt, calfbirth.calvingease, calfwean.actweight, calfwean.managecode, calfwean.group, calfwean.misc1, calfbirth.CowID, calfbirth.CowAge, calfbirth.sireID, cowprof.breed AS cow_breed, sireprof.breed AS sire_breed, calfwean.score as cframe "
   SQL$ = SQL$ & " FROM sireprof INNER JOIN (cowprof INNER JOIN (calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON cowprof.cowID = calfbirth.CowID) ON sireprof.SireID = calfbirth.sireID where sex = '" & CStr(Sex) & "' and calfbirth.birthdate >= #" & BeginDate & "# and calfbirth.birthdate <= #" & enddate & "# and calfbirth.herdid = '" & herdid & "'"
   SQL = SQL & " GROUP BY calfbirth.sex, calfbirth.HerdID, calfbirth.CalfID, calfbirth.birthdate, calfwean.dateweighed, calfbirth.birthwt, calfbirth.calvingease, calfwean.actweight, calfwean.managecode, calfwean.group, calfwean.misc1, calfbirth.CowID, calfbirth.CowAge, calfbirth.sireID, cowprof.breed, sireprof.breed, calfwean.score "
   dbChaps.Execute (SQL$), dbFailOnError 'grab all calves
   SQL$ = "UPDATE DISTINCTROW CalfList SET CalfList.age_in_days = [calflist].[dateweighed]-[calflist].[birthdate] where dateweighed > 0 and birthdate > #01/01/1900# and managecode <> 'A' and managecode <> 'B' and managecode <> 'C' and managecode <> 'D' "
   dbReps.Execute (SQL$), dbFailOnError 'update age_in_days
      
   'update adj birth wt
   dbReps.Execute "update calflist set adjbirthwt = birthwt + 8 where cowage <= 2 and birthwt > 0"
   dbReps.Execute "update calflist set adjbirthwt = birthwt + 5 where cowage = 3 and birthwt > 0"
   dbReps.Execute "update calflist set adjbirthwt = birthwt + 2 where cowage = 4 and birthwt > 0"
   dbReps.Execute "update calflist set adjbirthwt = birthwt + 0 where cowage >= 5 and cowage <= 10 and birthwt > 0"
   dbReps.Execute "update calflist set adjbirthwt = birthwt + 3 where cowage >= 11 and birthwt > 0"
   dbReps.Execute (SQL$), dbFailOnError
   
   SQL$ = "update calflist set adjbirthwt = 75 where isnull(birthwt) or birthwt = 0 and sex = 1" 'update birthwt for adj205
   dbReps.Execute (SQL$), dbFailOnError
   SQL$ = "update calflist set adjbirthwt = 75 where isnull(birthwt) or birthwt = 0 and sex = 3" 'update birthwt for adj205
   dbReps.Execute (SQL$), dbFailOnError
   SQL$ = "update calflist set adjbirthwt = 70 where isnull(birthwt) or birthwt = 0 and sex = 2"
   dbReps.Execute SQL, dbFailOnError
   
   dbReps.Execute ("update calflist set dam = 54 where sex = 2 and cowage = 2"), dbFailOnError 'update sex 2 dam adj
   dbReps.Execute ("update calflist set dam = 36 where sex = 2 and cowage = 3"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 18 where sex = 2 and cowage = 4"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 18 where sex = 2 and cowage >= 11"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 60 where sex = 1 and cowage = 2"), dbFailOnError 'update sexes 1 dam adj
   dbReps.Execute ("update calflist set dam = 40 where sex = 1 and cowage = 3"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 20 where sex = 1 and cowage = 4"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 20 where sex = 1 and cowage >= 11"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 60 where sex = 3 and cowage = 2"), dbFailOnError 'update sexes 3 dam adj
   dbReps.Execute ("update calflist set dam = 40 where sex = 3 and cowage = 3"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 20 where sex = 3 and cowage = 4"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 20 where sex = 3 and cowage >= 11"), dbFailOnError
   dbReps.Execute ("update calflist set dam = 0 where isnull(dam)"), dbFailOnError
   'update irrcalf flag for irregular calves and managecodes
   If calfreps.optCont(0).Value Then
      SQL = "SELECT Sum(IIf(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X',calflist.age_in_days,0)) AS Sum_A205, Sum(iif(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X', 1, 0)) AS Calf_Count, GROUP FROM calflist where sex = " & Sex & " GROUP BY calflist.sex, calflist.group"
      Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      While Not rsReps.EOF
        On Error Resume Next
        AvgAge = 0
        If Not rsReps.EOF Then
          AvgAge = Field2Num(rsReps!sum_a205) / Field2Num(rsReps!calf_count)
        Else
          AvgAge = 0
        End If
        On Error GoTo 0
        SQL$ = "update calflist set irrcalf = true where (age_in_days > " & AvgAge + 45 & " or age_in_days < " & AvgAge - 45 & " or managecode='A' or managecode='B' or managecode='C' or managecode='D' or managecode='E' or managecode='F' or managecode='N' or managecode='K' or managecode='S' or managecode='T' or managecode='X') AND CALFLIST.GROUP = '" & Field2Str(rsReps!group) & "' AND sex = " & Sex
        dbReps.Execute SQL, dbFailOnError
        rsReps.MoveNext
                End
   Else
                SQL = "SELECT Sum(IIf(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X',calflist.age_in_days,0)) AS Sum_A205, Sum(iif(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X', 1, 0)) AS Calf_Count FROM calflist where sex = " & Sex & " GROUP BY calflist.sex"
                rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
                On Error Resume Next
                If Not rsReps.EOF Then
                    If Field2Num(rsReps!calf_count) <> 0 Then
                        AvgAge = Field2Num(rsReps!sum_a205) / Field2Num(rsReps!calf_count)
                    Else
                        AvgAge = 0
                    End If
                End If
                On Error GoTo 0
                SQL = "update calflist set irrcalf = true where (age_in_days > " & AvgAge + 45 & " or age_in_days < " & AvgAge - 45 & " or managecode='A' or managecode='B' or managecode='C' or managecode='D' or managecode='E' or managecode='F' or managecode='N' or managecode='K' or managecode='S' or managecode='T' or managecode='X') and sex = " & Sex
                dbReps.Execute(SQL, dbFailOnError)
   End If
   
      
   SQL$ = "Update calflist set adj205wt = (((actweight - adjbirthwt )/ age_in_days) * 205) + adjbirthwt + dam where age_in_days > 0 and actweight > 0" 'update adj205wt
   dbReps.Execute (SQL$), dbFailOnError
   SQL$ = "update calflist set avgdailygain = (actweight - birthwt) / age_in_days where age_in_days > 0 and birthwt > 0 and actweight - birthwt > 0 and managecode <> 'F'" 'update avgdailygain
   dbReps.Execute (SQL$), dbFailOnError
   'dbReps.Execute ("delete avgdailygain from calflist where avgdailygain < 0"), dbFailOnError
   SQL$ = "update calflist set wt2daygain = actweight / age_in_days where age_in_days > 0 and actweight > 0" 'update wt2day gain
   dbReps.Execute (SQL$), dbFailOnError
   'dbReps.Execute ("delete wt2daygain from calflist where wt2daygain < 0 or managecode = 'F' or managecode = 'X'"), dbFailOnError
   dbReps.Execute ("update calflist set calflist.group = 0 where calflist.group = null"), dbFailOnError
   
   If calfreps.optCont(0).Value Then
      SQL = "SELECT Sum(IIf(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X',calflist.adj205wt,0)) AS Sum_A205, Sum(iif(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X', 1, 0)) AS Calf_Count, group FROM calflist where sex = " & Sex & " and not irrcalf GROUP BY calflist.sex, calflist.group"
      Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      Do Until rsReps.EOF
         On Error Resume Next
         'here maybe
         dbReps.Execute "update calflist set adj205rat = (adj205wt / " & (Field2Num(rsReps!sum_a205) / Field2Num(rsReps!calf_count)) & ") * 100 where sex = " & Sex & " and irrcalf = false AND calflist.GROUP = '" & Field2Str(rsReps!group) & "'"
         On Error GoTo 0
         rsReps.MoveNext
      Loop
   Else
      'SQL = "SELECT Sum(IIf(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X',calflist.adj205wt,0)) AS Sum_A205, Sum(iif(managecode<>'A' and managecode<>'B' and managecode<>'C' and managecode<>'D' and managecode<>'E' and managecode<>'F' and managecode<>'N' and managecode<>'K' and managecode<>'S' and managecode<>'T' and managecode<>'X', 1, 0)) AS Calf_Count FROM calflist where sex = " & Sex & " GROUP BY calflist.sex"
            SQL = "SELECT Sum(calflist.adj205wt) AS Sum_A205, Sum( 1) AS Calf_Count FROM calflist where (((calflist.sex)=" & Sex & ")  and not    ) GROUP BY calflist.sex"
      Set rsReps = dbReps.OpenRecordset(SQL, dbOpenSnapshot)
      If Not rsReps.EOF Then
         On Error Resume Next
         dbReps.Execute "update calflist set adj205rat = (adj205wt / " & (Field2Num(rsReps!sum_a205) / Field2Num(rsReps!calf_count)) & ") * 100 where sex = " & Sex & " and irrcalf = false"
         On Error GoTo 0
      End If
   End If
   'grab average age in days for irregular calves
             
'Averages group by sireid

SQL$ = "insert into calflistavg " '
If group = True Then
  SQL$ = SQL$ & " SELECT DISTINCTROW calflist.sex, calflist.SireID, calflist.group, calflist.sire_breed, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS adj205wt, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS calvingease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS actweight, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS age_in_days, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS avgdailygain, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wt2daygain, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS cframe, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom FROM calflist where sex = " & Sex & " GROUP BY calflist.SireID, calflist.group, calflist.sire_breed, calflist.sex"
Else
   SQL = SQL & " SELECT DISTINCTROW calflist.sex, calflist.SireID, calflist.sire_breed, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS adj205wt, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS calvingease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS actweight, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS age_in_days, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS avgdailygain, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wt2daygain, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS cframe, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom FROM calflist where sex = " & Sex & " GROUP BY calflist.SireID, calflist.sire_breed, calflist.sex"
End If
dbReps.Execute (SQL$), dbFailOnError
'3/27/2000 code seems to be here twice once here and once after the next insert into.
'Set rsReps = dbReps.OpenRecordset("select * from calflistavg where sex = " & Sex & " and sireid <> ''", dbOpenDynaset)
'If rsReps.RecordCount > 0 Then
'Do Until rsReps.EOF
   'rsReps.Edit
   'On Error Resume Next
   'rsReps!adj205wt = Field2Num(rsReps!adj205wt) / rsReps!num
   'rsReps!birthwt = Field2Num(rsReps!birthwt) / rsReps!bw_denom
   'rsReps!calvingease = Field2Num(rsReps!calvingease) / rsReps!num
   'rsReps!actweight = Field2Num(rsReps!actweight) / rsReps!num
   'rsReps!age_in_days = Field2Num(rsReps!age_in_days) / rsReps!num
   'rsReps!cframe = Field2Num(rsReps!cframe) / rsReps!fr_denom
   'rsReps!avgdailygain = Field2Num(rsReps!avgdailygain) / rsReps!agd_denom
   'rsReps!wt2daygain = Field2Num(rsReps!wt2daygain) / rsReps!num
   'On Error GoTo 0
   'rsReps.Update
   'rsReps.MoveNext
'Loop
'End If

'Averages group by cow breed
SQL = " insert into calflistavg "
If group = True Then
   SQL$ = SQL$ & " SELECT DISTINCTROW calflist.sex, calflist.cow_breed, calflist.group, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS adj205wt, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS calvingease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS actweight, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS age_in_days, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS avgdailygain, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wt2daygain, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS cframe, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom From calflist where sex = " & Sex & " GROUP BY calflist.group, calflist.cow_breed, calflist.sex"
Else
   SQL$ = SQL$ & " SELECT DISTINCTROW calflist.sex, calflist.cow_breed, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS adj205wt, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS calvingease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS actweight, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS age_in_days, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS avgdailygain, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wt2daygain, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS cframe, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom From calflist where sex = " & Sex & " GROUP BY calflist.cow_breed, calflist.sex"
End If
dbReps.Execute SQL

 Set rsReps = dbReps.OpenRecordset("select * from calflistavg where sex = " & Sex, dbOpenDynaset)
   Do Until rsReps.EOF
      rsReps.Edit
      On Error Resume Next
      rsReps!adj205wt = Field2Num(rsReps!adj205wt) / rsReps!num
      rsReps!birthwt = Field2Num(rsReps!birthwt) / rsReps!bw_denom
      rsReps!calvingease = Field2Num(rsReps!calvingease) / rsReps!num
      rsReps!actweight = Field2Num(rsReps!actweight) / rsReps!num
      rsReps!age_in_days = Field2Num(rsReps!age_in_days) / rsReps!num
      If rsReps!fr_denom <> 0 Then rsReps!cframe = Field2Num(rsReps!cframe) / rsReps!fr_denom Else rsReps!cframe = 0
      rsReps!avgdailygain = Field2Num(rsReps!avgdailygain) / rsReps!agd_denom
      rsReps!wt2daygain = Field2Num(rsReps!wt2daygain) / rsReps!num
      On Error GoTo 0
      rsReps.Update
      rsReps.MoveNext
   Loop

'Group Summary
 SQL = "insert into calflistsum "
 If group = True Then
   SQL = SQL & "select calflist.group, calflist.sex, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS wt205, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS cease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS allwt, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS avgage, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS adg, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wdg, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS frscore, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom FROM calflist where sex = " & Sex & " GROUP BY calflist.sex, calflist.group"
 Else
   SQL = SQL & "select calflist.sex, Sum(IIf([irrcalf]=False,1,0)) AS num, Sum(IIf([irrcalf]=False,[calflist].[adj205wt],0)) AS wt205, Sum(IIf([irrcalf]=False,[calflist].[birthwt],0)) AS birthwt, Sum(IIf([irrcalf]=False,[calflist].[calvingease],0)) AS cease, Sum(IIf([irrcalf]=False,[calflist].[actweight],0)) AS allwt, Sum(IIf([irrcalf]=False,[calflist].[age_in_days],0)) AS avgage, Sum(IIf([irrcalf]=False,[calflist].[avgdailygain],0)) AS adg, Sum(IIf([irrcalf]=False,[calflist].[wt2daygain],0)) AS wdg, Sum(IIf([irrcalf]=False,[calflist].[cframe],0)) AS frscore, sum(iif(calflist.irrcalf = false and calflist.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(calflist.irrcalf = false and calflist.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(calflist.irrcalf = false and calflist.avgdailygain > 0, 1, 0)) as AGD_Denom FROM calflist where sex = " & Sex & " GROUP BY calflist.sex"
 End If
 dbReps.Execute SQL
 Set rsReps = dbReps.OpenRecordset("select * from calflistsum where sex = " & Sex, dbOpenDynaset)
 Do Until rsReps.EOF
      rsReps.Edit
      On Error Resume Next
      rsReps!wt205 = Field2Num(rsReps!wt205) / rsReps!num
      rsReps!birthwt = Field2Num(rsReps!birthwt) / rsReps!bw_denom
      rsReps!cease = Field2Num(rsReps!cease) / rsReps!num
      rsReps!allwt = Field2Num(rsReps!allwt) / rsReps!num
      rsReps!AvgAge = Field2Num(rsReps!AvgAge) / rsReps!num
      If rsReps!fr_denom <> 0 Then rsReps!frscore = Field2Num(rsReps!frscore) / rsReps!fr_denom Else rsReps!frscore = 0
      rsReps!adg = Field2Num(rsReps!adg) / rsReps!agd_denom
      rsReps!wdg = Field2Num(rsReps!wdg) / rsReps!num
      On Error GoTo 0
      rsReps.Update
      rsReps.MoveNext
Loop
NextSex:
Next

Call CreateCowSummary
Call CreateCalfDistReport(thirdcowdate, age, rowcount, avgwt, allcalf)
'age = 0
Call CreateCalfDistReport_12(thirdcowdate, age, rowcount, avgwt, allcalf)
Call CreateAvgActWnWt(thirdcowdate, age, rowcount, avgwt, allcalf)
Call CreateCritSuccFactors
Call CreatePerCalv
'Call CreateSireSum
Call CreateRepPerf
Call CreateFRRP

'create sire summary report
dbReps.Execute ("delete * from siresumtmp"), dbFailOnError
dbReps.Execute ("insert into siresumtmp select * from calflist"), dbFailOnError
'get all sire ids
dbReps.Execute ("delete * from siresum"), dbFailOnError
'If group = False Then
   dbReps.Execute "update siresumtmp set adj205wt = switch(sex = 2, adj205wt * 1.05, sex = 1, adj205wt * .95, sex = 3, adj205wt * 1.00, sex = 0, adj205wt * 1.00)"
   dbReps.Execute ("insert into siresum SELECT DISTINCTROW sum(iif(irrcalf = false, 1, 0)) AS num, Sum(iif(irrcalf = false, SireSumTmp.birthwt, 0)) AS birthwt, Sum(iif(irrcalf = false, SireSumTmp.calvingease,0)) AS calvingease, Sum(iif(irrcalf = false, SireSumTmp.actweight, 0)) AS actweight, Sum(iif(irrcalf = false, SireSumTmp.cframe, 0)) AS cframe, Sum(iif(irrcalf = false, SireSumTmp.age_in_days, 0)) AS age_in_days, " & _
      "Sum(iif(irrcalf = false, siresumtmp.adj205wt, 0)) AS adj205wt, Sum(iif(irrcalf = false, SireSumTmp.avgdailygain,0)) AS avgdailygain, Sum(iif(irrcalf = false, SireSumTmp.wt2daygain, 0)) AS wt2daygain, SireSumTmp.SireID, siresumtmp.sire_breed, sum(iif(siresumtmp.irrcalf = false and siresumtmp.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(siresumtmp.irrcalf = false and siresumtmp.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(siresumtmp.irrcalf = false and siresumtmp.avgdailygain > 0, 1, 0)) as AGD_Denom From SireSumTmp GROUP BY SireSumTmp.SireID,siresumtmp.sire_breed"), dbFailOnError
'Else
'   dbReps.Execute ("insert into siresum SELECT DISTINCTROW sum(iif(irrcalf = false, 1, 0)) AS num, Sum(iif(irrcalf = false, SireSumTmp.birthwt, 0)) AS birthwt, Sum(iif(irrcalf = false, SireSumTmp.calvingease,0)) AS calvingease, Sum(iif(irrcalf = false, SireSumTmp.actweight, 0)) AS actweight, Sum(iif(irrcalf = false, SireSumTmp.cframe, 0)) AS cframe, Sum(iif(irrcalf = false, SireSumTmp.age_in_days, 0)) AS age_in_days, Sum(iif(irrcalf = false, SireSumTmp.adj205wt, 0)) AS adj205wt, Sum(iif(irrcalf = false, SireSumTmp.avgdailygain,0)) AS avgdailygain, Sum(iif(irrcalf = false, SireSumTmp.wt2daygain, 0)) AS wt2daygain, SireSumTmp.SireID, siresumtmp.sire_breed, sum(iif(siresumtmp.irrcalf = false and siresumtmp.birthwt > 0, 1, 0)) as BW_Denom, sum(iif(siresumtmp.irrcalf = false and siresumtmp.cframe > 0, 1, 0)) as Fr_Denom, sum(iif(siresumtmp.irrcalf = false and siresumtmp.avgdailygain > 0, 1, 0)) as AGD_Denom From SireSumTmp GROUP BY SireSumTmp.SireID, siresumtmp.sire_breed, siresumtmp.group"), dbFailOnError
'   dbReps.Execute ("update siresum set siresum.group = 0 where siresum.group = null"), dbFailOnError
'End If
'update averages
dbReps.Execute ("update siresum set birthwt = birthwt / bw_denom, actweight = actweight / num, adj205wt = adj205wt / num, avgdailygain = avgdailygain / agd_denom, wt2daygain = wt2daygain / num, calvingease = calvingease / num, age_in_days = age_in_days / num, cframe = cframe / fr_denom")

'sort calfid's
Select Case SortOrder
   Case 0
      order$ = "calfid"
   Case 1
      order$ = "adj205wt desc"
   Case 2
      order$ = "sireid"
   Case 3
      order$ = "actweight"
   Case 4
      order$ = "age_in_days"
   Case 5
      order$ = "cowage"
   Case 6
      order$ = "birthwt"
   Case 7
      order$ = "cframe"
End Select
  
  dbReps.Execute ("delete * from calfsort"), dbFailOnError
  dbReps.Execute ("insert into calfsort select * from calflist order by " & order$), dbFailOnError
  dbReps.Execute ("delete * from calflist"), dbFailOnError
  dbReps.Execute ("insert into calflist select * from calfsort"), dbFailOnError
  dbReps.Execute ("delete * from calfsort")
  dbReps.Execute ("insert into calfsort select * from calflistavg order by " & order$), dbFailOnError
  dbReps.Execute ("delete * from calflistavg")
  dbReps.Execute ("insert into calflistavg select * from calfsort")

'update adj205 ratio in calfwean
'If IsDate(calfreps.txtWeighDate) Then
   'iResponse = MsgBox("Do You Wish To Overwrite The Adjusted 205 Ratio?", vbYesNo + vbQuestion)
   'If iResponse = vbYes Then
   If calfreps.chkOverwrite.Value = vbChecked Then
      Call CreateTableAttachment(dbfile$, repfile$, "calflist", "calflist")
      dbChaps.Execute ("update calfwean, calflist set calfwean.ratio = calflist.adj205rat, calfwean.score = calflist.cframe, calfwean.wt205 = calflist.adj205wt where calfwean.calfid = calflist.calfid and calflist.irrcalf = false"), dbFailOnError
      dbChaps.Execute "update calfwean, calflist set calfwean.ratio = 0, calfwean.score = calflist.cframe, calfwean.wt205 = calflist.adj205wt where calfwean.calfid = calflist.calfid and calflist.irrcalf = true"
      Call DeleteTableAttachment(dbfile$, "calflist")
   End If
'End If
rsReps.Close: Set rsReps = Nothing
dbChaps.Close: Set dbChaps = Nothing
dbReps.Close: Set dbReps = Nothing
End Sub

Public Function CreateCalfTables(ByVal Sex As Integer, group As Boolean)
On Error GoTo ehandle
Dim SQL, Col(2), fieldvar$(2), FORMULA$, hmfields%, strSQL$
Dim table$, DB As database, DBREP As database
Select Case Sex
   Case 0
      table$ = "misccalf"
   Case 1
      table$ = "steercalf"
   Case 2
      table$ = "heical"
   Case 3
      table$ = "bulcalf"
End Select
SQL = "INSERT INTO " & table$ & " (herdid, CalfID, birthdate, sex, birthwt, dateweighed, calvingease, actweight, "
SQL = SQL & " managecode, cframe, [group], misc1, CowID, sireID, age, cow_breed, sire_breed ) "
SQL = SQL & " IN '" & repfile$ & "'"
SQL = SQL & " SELECT DISTINCTROW calfbirth.herdid, calfbirth.CalfID, calfbirth.birthdate, calfbirth.sex, calfbirth.birthwt, "
SQL = SQL & " calfwean.dateweighed, calfbirth.calvingease, calfwean.actweight, calfwean.managecode, "
SQL = SQL & " calfwean.cframe, calfwean.group, calfbirth.misc1, calfbirth.CowID, calfbirth.sireID, calfbirth.CowAge, "
SQL = SQL & " cowprof.breed, sireprof.breed FROM ((cowprof INNER JOIN (sireprof RIGHT JOIN calfbirth "
SQL = SQL & " ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID)) ON "
SQL = SQL & " (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT "
SQL = SQL & " JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) "
SQL = SQL & " LEFT JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID)"
   
 '  hmfields% = 1
 ' Col(1) = 0
 ' fieldvar$(1) = "calfBIRTH.herdid"
 ' If calfreps.lblhow_many_herd <> "All" Then
 '  Call create_sql_selection(FrmSelect_Multi_Herds!lstherd, Col(), fieldvar$(), hmfields%, formula$)
 '  SQL = SQL & " WHERE (((calfbirth.sex)= '" & Sex & "' )) and " & formula$ & ""
'Else
   SQL = SQL & " WHERE (((calfbirth.sex)='" & Sex & "')) "
'End If
'SQL = SQL & " and  (((calfbirth.birthdate)>#" & calfreps.txtStartDate & "# And (calfbirth.birthdate)<#" & calfreps.txtEndDate & "#))"
If IsDate(calfreps.txtStartDate) Then
   SQL = SQL & "And ((calfbirth.birthdate) > #" & calfreps.txtStartDate.TEXT & "#) "
End If
If IsDate(calfreps.txtEndDate) Then
   SQL = SQL & "And ((calfbirth.birthdate) < #" & calfreps.txtEndDate.TEXT & "#) "
End If
'Contempary Groups
  If group = True Then SQL = SQL & " order by calfwean.group "
  
'Call ReturnSortOrder(strSQL$)
'SQL = SQL & strSQL$

Set DB = DBEngine(0).OpenDatabase(dbfile$, False, True)
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False)
DBREP.Execute ("delete * from " & table$)
DB.Execute (SQL)

If DB.RecordsAffected > 0 Then
   Select Case Sex
      Case 0
         DataFound(0) = True
         report.Setformulas("SuppressMisc") = False
      Case 1
         DataFound(1) = True
         report.Setformulas("SuppressSteers") = False 'Really the Bulls
      Case 2
         DataFound(2) = True
         report.Setformulas("SuppressHeifers") = False
      Case 3
         DataFound(3) = True
         report.Setformulas("SuppressBulls") = False 'Really the Steers
   End Select
End If



Exit Function
ehandle:
If Err.Number = 94 Then Resume Next
End Function

Public Function ReturnCalfOrder(SortOrder As Integer) As String
Dim SQL$
Select Case SortOrder
   Case 0
      SQL$ = " ORDER BY calfbirth.calfid"
   Case 1
      SQL$ = " ORDER BY calfbirth.calfid"
   Case 2
      SQL$ = " ORDER BY calfbirth.sireid"
   Case 3
      SQL$ = " ORDER BY calfwean.actweight"
   Case 4
      SQL$ = " ORDER BY calfbirth.birthdate"
   Case 5
      SQL$ = " ORDER BY calfbirth.cowage"
   Case 6
      SQL$ = " ORDER BY calfbirth.birthwt"
   Case 7
      SQL$ = " ORDER BY calfbirth.cframe"
End Select
End Function

Public Sub CreateAvgActWnWt(thirdcowdate As Date, avgdate As Double, rowcount As Double, avgwt As Double, allcalf As Double)
Dim DB As database, RS As Recordset
Dim period(5) As Double, AvgWn(5) As Double, no_calf As Double, lbsbeef As Double, T_Mean_Denom&
Dim TMEAN As Double, UScor As Double, pTestDate As Date
On Error GoTo ehandle

'If allcalf > 0 Then
'   UScor = avgwt / allcalf
'Else
'   UScor = 1
'End If

pTestDate = mTestDate

Set DB = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DB.OpenRecordset("siresum", dbOpenTable)
With RS
   Do Until .EOF
     If !actweight > 0 Then
     'T_Mean_Denom = T_Mean_Denom + 1
     'TMEAN = TMEAN + ((UScor - !actweight) * (UScor - !actweight))
      
      If !managecode <> "A" Or !managecode <> "B" And !skipme = False Then
         no_calf = no_calf + 1
         lbsbeef = lbsbeef + !actweight
      End If
      If !birthdate < pTestDate Then
         period(0) = period(0) + 1
         AvgWn(0) = AvgWn(0) + !actweight
      End If
      If !birthdate >= pTestDate And !birthdate <= pTestDate + 20 Then
         period(1) = period(1) + 1
         AvgWn(1) = AvgWn(1) + !actweight
      End If
      If !birthdate >= pTestDate + 21 And !birthdate <= pTestDate + 41 Then
         period(2) = period(2) + 1
         AvgWn(2) = AvgWn(2) + !actweight
      End If
      If !birthdate >= pTestDate + 42 And !birthdate <= pTestDate + 62 Then
         period(3) = period(3) + 1
         AvgWn(3) = AvgWn(3) + !actweight
      End If
      If !birthdate >= pTestDate + 63 And !birthdate <= pTestDate + 83 Then
         period(4) = period(4) + 1
         AvgWn(4) = AvgWn(4) + !actweight
      End If
      If !birthdate > pTestDate + 84 Then
         period(5) = period(5) + 1
         AvgWn(5) = AvgWn(5) + !actweight
      End If
      .MoveNext
    Else
      .MoveNext
    End If
   Loop
End With
DB.Execute ("delete * from avgactwnwt")

If no_calf <> 0 Then
  UScor = lbsbeef / no_calf
Else
  UScor = 0
End If
   'loop through for herd uniformity
   If Not RS.EOF And Not RS.BOF Then RS.MoveFirst
   Do Until RS.EOF
      If RS!actweight > 0 And RS!managecode <> "A" And RS!managecode <> "B" Then TMEAN = TMEAN + ((RS!actweight - UScor) * (RS!actweight - UScor))
      RS.MoveNext
   Loop
   

Set RS = DB.OpenRecordset("AvgActWnWt")
With RS
.AddNew
   If period(0) > 0 Then
      !avgact0 = AvgWn(0) / period(0) 'avg actweight for early
   Else
      !avgact0 = 0
   End If
   If period(1) > 0 Then
      !avgact1 = AvgWn(1) / period(1) 'avg actweight for 1st 21
   Else
      !avgact1 = 0
   End If
   If period(2) > 0 Then
      !avgact2 = AvgWn(2) / period(2) 'avg actweight for 2nd 21
   Else
      !avgact2 = 0
   End If
   If period(3) > 0 Then
      !avgact3 = AvgWn(3) / period(3)
   Else
      !avgact3 = 0
   End If
   If period(4) > 0 Then
      !avgact4 = AvgWn(4) / period(4)
   Else
      !avgact4 = 0
   End If
   If period(5) > 0 Then
      !avgact5 = AvgWn(5) / period(5)
   Else
      !avgact5 = 0
   End If
   
   !ThirdCow = ThirdCow 'third mature cow calving string
   !thirdcowdate = thirdcowdate 'third mature cow calving date
   !lbsbeef = lbsbeef 'lbs of beef produced at weaning
   !no_calf = no_calf 'denominator for lbsbeef avg
   !TurnDate = ActTurnDate  'bull turnout date
   !ReportDate = thirdcowdate - 285 'calculated turn out date
   If Calculated = False Then
      !ReportDate = !ReportDate & "*"
   Else
      !TurnDate = !TurnDate & "*"
   End If
   If no_calf <> 0 Then
     !avglbsbeef = lbsbeef / no_calf 'avg lbs beef produced at weaning
   Else
     !avglbsbeef = 0
   End If
   If rowcount <> 0 Then
     !avgdate = avgdate / rowcount 'avg birthdate
   Else
     !avgdate = Null
   End If
   !estimated = IIf(Calculated, "T", "F")
   !UScor = Sqr(TMEAN / (no_calf - 1))
   'If ValidTurnDate = False Then !estimated = "Estimated"
   .Update
End With
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub




Public Sub CreateFRRP()
On Error GoTo ehandle
Dim SQL As String, frrp(1) As Double
Dim RepRS As Recordset, repdb As database
Dim RS As Recordset, DB As database, pTestDate As Date
Dim tbCritSucc As DAO.Recordset
Dim RepCalv As Long

If Calculated = False Then pTestDate = thirdcowdate - 285 Else pTestDate = ActTurnDate

SQL = "select sum(iif(cowprof.active <> 'P' and cowprof.enteredherd >= #" & pTestDate & "# and cowprof.enteredherd <= #" & pTestDate + 365 & "# and cowprof.herdid = '" & herdid & "', 1, 0)) as cow_count  from cowprof"

Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DB.OpenRecordset("select spasource from prefspa", dbOpenSnapshot)
If Not RS.EOF Then
   repdb.Execute "update reproper set spa_header = '" & Field2Str(RS!spasource) & "'"
End If
Set RS = DB.OpenRecordset(SQL)
Set RepRS = repdb.OpenRecordset("herdcount")
'If rs.RecordCount > 0 And reprs.RecordCount > 0 Then
'frrp(0) = Field2Num(RepRS!noofcalves)
Set RepRS = repdb.OpenRecordset("cowcount", dbOpenSnapshot)
If Not RS.EOF Then
      If Field2Num(RS!cow_count) = 0 Then GoTo Skip
      frrp(0) = (Field2Num(RS!cow_count) / (RepRS!TCEXP - RepRS!h - RepRS!j - RepRS!L - RepRS!r - RepRS!y)) * 100
         Set RepRS = repdb.OpenRecordset("reproper", dbOpenTable)
         RepRS.Edit: RepRS!frrp = frrp(0): RepRS.Update
Else
Skip:
   Set tbCritSucc = repdb.OpenRecordset("select repcalv from critsuccfac", dbOpenSnapshot)
   If Not tbCritSucc.EOF Then RepCalv = Field2Num(tbCritSucc!RepCalv)
   tbCritSucc.Close: Set tbCritSucc = Nothing
      'frrp(1) = frrp(0)
      'frrp(0) = ((frrp(0) + P) / (RepRS!TCEXP - RepRS!h - RepRS!j - RepRS!L - RepRS!R - RepRS!Y)) * 100
         frrp(0) = ((RepCalv + P) / (RepRS!TCEXP - RepRS!h - RepRS!j - RepRS!L - RepRS!r - RepRS!y)) * 100
         Set RepRS = repdb.OpenRecordset("reproper", dbOpenTable)
         RepRS.Edit: RepRS!frrp = frrp(0): RepRS.Update
End If
'Set RepRS = repdb.OpenRecordset("reproper", dbOpenTable)
'RepRS.Edit: RepRS!frrp = frrp(0): RepRS.Update
'End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub

Public Sub CreateCalfDistReport_12(thirdcowdate As Date, avgdate As Double, rowcount As Double, avgwt As Double, allcalf As Double)
Dim RS As Recordset, DB As database, SQL
Dim codecount As Double, no_calves As Double, avgwgt As Double
Dim calf As Double, period(5) As Double, age As Double, pTestDate As Date
On Error GoTo ehandle
'if no third cow date then calculate test date for calving distribution table
'If BoolThirdCowDate = False Then pTestDate = TurnDate + 285 Else pTestDate = thirdcowdate
pTestDate = mTestDate
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
SQL = "select * from siresum where cowage > 11"
Set RS = DB.OpenRecordset(SQL)
With RS
  If RS.RecordCount > 0 Then
   .MoveFirst
   Do Until .EOF
      If !actweight > 0 Then
         avgwgt = avgwgt + !actweight
         calf = calf + 1
      End If
      If !managecode = "A" Or !managecode = "B" Then
         codecount = codecount + 1

      End If
      If !managecode <> "A" And !managecode <> "B" Then
         no_calves = no_calves + 1
      
         If !birthdate < pTestDate Then
           period(0) = period(0) + 1
         End If
         If !birthdate >= pTestDate And !birthdate <= pTestDate + 20 Then
           period(1) = period(1) + 1
         End If
         If !birthdate >= pTestDate + 21 And !birthdate <= pTestDate + 41 Then
           period(2) = period(2) + 1
         End If
         If !birthdate >= pTestDate + 42 And !birthdate <= pTestDate + 62 Then
           period(3) = period(3) + 1
         End If
         If !birthdate >= pTestDate + 63 And !birthdate <= pTestDate + 83 Then
           period(4) = period(4) + 1
         End If
         If !birthdate > pTestDate + 84 Then
           period(5) = period(5) + 1
         End If
'      If !age_in_days > 0 Then
         age = age + !birthdate
         'Debug.Print !birthdate & "     "; age
'      End If
   End If
   .MoveNext
   Loop
   End If
 
 Set RS = DB.OpenRecordset("herdcount", dbOpenTable)
   RS.AddNew
      RS!DamAge = 12
      RS!noofcalves = no_calves
      RS!period0 = period(0)
      RS!period1 = period(1)
      RS!period2 = period(2)
      RS!period3 = period(3)
      RS!period4 = period(4)
      RS!period5 = period(5)
      RS!openab = codecount
      If no_calves > 0 Then
         'RS!avgdate = age / no_calves
         'RS!avgdate = age / calf
         RS!avgdate = age / no_calves
      End If
      avgdate = avgdate + age
      If calf > 0 Then
         RS!avgwt = avgwgt / calf
      End If
      allcalf = allcalf + calf
      avgwt = avgwt + avgwgt
      rowcount = rowcount + no_calves
   RS.Update
End With
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
'Call CreateAvgActWnWt(thirdcowdate, age(1), tcc, avgwgt(1), calf(1))
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub

Public Sub CreateMiscCows(group As Boolean)
On Error GoTo ehandle
Dim RS As Recordset, GRS As Recordset, GSql$
Dim DB As database, DBREP As database
Dim SQL As String, agedays As Boolean, adj205 As Boolean, adj205r As Boolean, avgdgain As Boolean, wt2day As Boolean
Dim dam As Double, allcalf As Double, AvgAge As Double, x As Double, allwt As Double, IrrCalf As Double, y As Double
Dim wt205 As Double, birthwt As Double, cease As Double, frscor As Double, adg As Double, wdg As Double
Dim skipme As Double
'Dim formula$, Col(1), fieldvar$(1), hmfields%, irrcalfwt As Double

Call CreateCalfTables(0, False)
If DataFound(0) = False Then Exit Sub
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, True)
Set RS = DBREP.OpenRecordset("misccalf", dbOpenTable)

If RS.RecordCount > 0 Then
'DataFound(0) = True
Do Until RS.EOF
   GoSub agedays
   GoSub adj205
   GoSub avgdgain
   GoSub wt2day
   If RS!skipme = False And RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
      'allwt = allwt + rs!adj205wt
      x = x + RS!age_in_days
   Else
      skipme = skipme + 1
   End If
   RS.MoveNext
Loop

If RS.RecordCount <> skipme Then
   AvgAge = x / (RS.RecordCount - skipme)
Set RS = DBREP.OpenRecordset("misccalf")
RS.MoveFirst
x = 0
IrrCalf = 0
Do Until RS.EOF
      'If rs!calfID = "7013" Then
      '   Stop
      'End If
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
AvgAge = x / allcalf
RS.MoveFirst
Do Until RS.EOF
   GoSub adj205r
   RS.MoveNext
Loop

RS.Close: Set RS = Nothing
DBREP.Close: Set DBREP = Nothing
End If
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
         If RS!age = 2 Then
            dam = 54
         End If
         If RS!age = 3 Then
            dam = 36
         End If
      If RS!age = 4 Then
         dam = 18
      End If
      If RS!age > 11 Then
         dam = 18
      End If
   End If
   If RS!Sex = 1 Or RS!Sex = 3 Then
      If RS!age = 2 Then
         dam = 60
      End If
      If RS!age = 3 Then
         dam = 40
      End If
      If RS!age = 4 Then
         dam = 20
      End If
      If RS!age > 11 Then
         dam = 20
      End If
   End If
   RS!adj205wt = (((RS!actweight - RS!birthwt) / RS!age_in_days) * 205) + RS!birthwt + dam
   End If
   RS.Update
Return

adj205r:
If group = False Then
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
Else
   GSql$ = "SELECT Sum(heical.adj205wt) AS allwt, Count(heical.CalfID) AS allcalf, heical.group From heical where (((heical.adj205wt) > 0)) GROUP BY heical.group"
   Set GRS = DBREP.OpenRecordset(GSql$, dbOpenSnapshot)
   If GRS.RecordCount > 0 Then
      GRS.FindFirst Field2Str(GRS!group) = Field2Str(RS!group)
      RS.Edit: RS!adj205rat = (RS!adj205wt / (GRS!allwt / GRS!allcalf)) * 100: RS.Update
   End If
   GRS.Close: Set GRS = Nothing
End If
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
ehandle:
If Err.Number = 94 Then Resume Next
End Sub




Public Sub CreateSBulAvg(SortOrder%, group As Boolean)
Dim repdb As database, RepRS As Recordset, SQL$
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from sbulavg")
SQL$ = "insert into sbulavg "
If group = True Then
   SQL$ = SQL$ & " SELECT DISTINCTROW steercalf.SireID, steercalf.sire_breed, Count(steercalf.CalfID) AS num, Sum(steercalf.adj205wt) AS adj205wt, Sum(steercalf.birthwt) AS birthwt, Sum(steercalf.calvingease) AS calvingease, Sum(steercalf.actweight) AS actweight, Sum(steercalf.age_in_days) AS age_in_days, Sum(steercalf.avgdailygain) AS avgdailygain, Sum(steercalf.wt2daygain) AS wt2daygain, steercalf.group From steercalf"
   SQL$ = SQL$ & " where (((steercalf.skipme) = False)) GROUP BY steercalf.SireID, steercalf.sire_breed, steercalf.group"
Else
   SQL$ = SQL$ & " SELECT DISTINCTROW steercalf.SireID, steercalf.sire_breed, Count(steercalf.CalfID) AS num, Sum(steercalf.adj205wt) AS adj205wt, Sum(steercalf.birthwt) AS birthwt, Sum(steercalf.calvingease) AS calvingease, Sum(steercalf.actweight) AS actweight, Sum(steercalf.age_in_days) AS age_in_days, Sum(steercalf.avgdailygain) AS avgdailygain, Sum(steercalf.wt2daygain) AS wt2daygain"
   SQL$ = SQL$ & " From steercalf where (((steercalf.skipme) = False)) GROUP BY steercalf.SireID, steercalf.sire_breed"
End If
repdb.Execute (SQL$), dbFailOnError
Set RepRS = repdb.OpenRecordset("sbulavg", dbOpenTable)
If RepRS.RecordCount > 0 Then
Do Until RepRS.EOF
   RepRS.Edit
   RepRS!adj205wt = Field2Num(RepRS!adj205wt) / RepRS!num
   RepRS!birthwt = Field2Num(RepRS!birthwt) / RepRS!num
   RepRS!calvingease = Field2Num(RepRS!calvingease) / RepRS!num
   RepRS!actweight = Field2Num(RepRS!actweight) / RepRS!num
   RepRS!age_in_days = Field2Num(RepRS!age_in_days) / RepRS!num
   RepRS!cframe = Field2Num(RepRS!cframe) / RepRS!num
   RepRS!avgdailygain = Field2Num(RepRS!avgdailygain) / RepRS!num
   RepRS!wt2daygain = Field2Num(RepRS!wt2daygain) / RepRS!num
   RepRS.Update
   RepRS.MoveNext
Loop
End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
End Sub


Public Sub CreateSHeiAvg(SortOrder%, group As Boolean)
Dim SQL$, repdb As database, RepRS As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from sheiavg")
If group = True Then
   repdb.Execute ("insert into sheiavg SELECT steercalf.cow_breed, Count(steercalf.CalfID) AS Num, Sum(steercalf.adj205wt) AS adj205wt, Sum(steercalf.birthwt) AS birthwt, steercalf.group, Sum(steercalf.calvingease) AS calvingease, Sum(steercalf.actweight) AS actweight, Sum(steercalf.age_in_days) AS age_in_days, Sum(steercalf.cframe) AS cframe, Sum(steercalf.avgdailygain) AS avgdailygain, Sum(steercalf.wt2daygain) AS wt2daygain From steercalf GROUP BY steercalf.cow_breed, steercalf.group, steercalf.skipme HAVING (((steercalf.skipme)=False))"), dbFailOnError
Else
   repdb.Execute ("insert into sheiavg SELECT steercalf.cow_breed, Count(steercalf.CalfID) AS Num, Sum(steercalf.adj205wt) AS adj205wt, Sum(steercalf.birthwt) AS birthwt, Sum(steercalf.calvingease) AS calvingease, Sum(steercalf.actweight) AS actweight, Sum(steercalf.age_in_days) AS age_in_days, Sum(steercalf.cframe) AS cframe, Sum(steercalf.avgdailygain) AS avgdailygain, Sum(steercalf.wt2daygain) AS wt2daygain From steercalf GROUP BY steercalf.cow_breed, steercalf.skipme HAVING (((steercalf.skipme)=False))"), dbFailOnError
End If
Set RepRS = repdb.OpenRecordset("sheiavg", dbOpenTable)
If RepRS.RecordCount > 0 Then
   Do Until RepRS.EOF
      RepRS.Edit
      RepRS!adj205wt = Field2Num(RepRS!adj205wt) / RepRS!num
      RepRS!birthwt = Field2Num(RepRS!birthwt) / RepRS!num
      RepRS!calvingease = Field2Num(RepRS!calvingease) / RepRS!num
      RepRS!actweight = Field2Num(RepRS!actweight) / RepRS!num
      RepRS!age_in_days = Field2Num(RepRS!age_in_days) / RepRS!num
      RepRS!cframe = Field2Num(RepRS!cframe) / RepRS!num
      RepRS!avgdailygain = Field2Num(RepRS!avgdailygain) / RepRS!num
      RepRS!wt2daygain = Field2Num(RepRS!wt2daygain) / RepRS!num
      RepRS.Update
      RepRS.MoveNext
   Loop
End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
End Sub


Public Sub CreateSteerSum(group As Boolean)
On Local Error GoTo ehandle
Dim repdb As database, RepRSData As Recordset, RepRSEdit As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from steersum")
If group = True Then
   'Get all group identifiers for steersum group averages
   repdb.Execute ("insert into steersum select group from heical group by group order by group")
 If repdb.RecordsAffected > 0 Then
   repdb.Execute ("update steersum set steersum.group = 0 where isnull(group)")
   repdb.Execute ("update heical set heical.group = 0 where isnull(group)")
   Set RepRSEdit = repdb.OpenRecordset("steersum", dbOpenTable)
   'update calvingease
   Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num, heical.group From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4)) GROUP BY heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!cease = RepRSData!cease / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num, heical.group from heical where heical.birthwt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num, heical.group from heical where heical.adj205wt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num, heical.group from heical where heical.actweight > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num, heical.group from heical where heical.age_in_days > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num, heical.group from heical where heical.cframe > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!frscor = RepRSData!frscor / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num, heical.group from heical where heical.avgdailygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!adg = RepRSData!adg / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num, heical.group from heical where heical.wt2daygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wdg = RepRSData!wda / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
 End If
Else
  Set RepRSEdit = repdb.OpenRecordset("steersum", dbOpenTable)
  RepRSEdit.AddNew
   'update calvingease
  Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4))", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!cease = RepRSData!cease / RepRSData!num
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num from heical where heical.birthwt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num from heical where heical.adj205wt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num from heical where heical.actweight > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num from heical where heical.age_in_days > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num from heical where heical.cframe > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!frscore = RepRSData!frscor / RepRSData!num
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num from heical where heical.avgdailygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!adg = RepRSData!adg / RepRSData!num
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num from heical where heical.wt2daygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wdg = RepRSData!wda / RepRSData!num
   RepRSEdit.Update
 End If
RepRSData.Close: Set RepRSData = Nothing
RepRSEdit.Close: Set RepRSEdit = Nothing
repdb.Close: Set repdb = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next Else RepRSData.Close: Set RepRSData = Nothing: RepRSEdit.Close: Set RepRSEdit = Nothing: repdb.Close: Set repdb = Nothing: MsgBox Err.Description

End Sub
Public Sub CreateSteerCalf(group As Boolean)
On Error GoTo ehandle
Dim RS As Recordset, GRS As Recordset, GSql$
Dim DB As database, DBREP As database
Dim SQL As String, agedays As Boolean, adj205 As Boolean, adj205r As Boolean, avgdgain As Boolean, wt2day As Boolean
Dim dam As Double, allcalf As Double, AvgAge As Double, x As Double, allwt As Double, IrrCalf As Double, y As Double
Dim wt205 As Double, birthwt As Double, cease As Double, frscor As Double, adg As Double, wdg As Double
Dim skipme As Double
'Dim formula$, Col(1), fieldvar$(1), hmfields%, irrcalfwt As Double

Call CreateCalfTables(1, False)
If DataFound(1) = False Then Exit Sub
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DBREP.OpenRecordset("steercalf", dbOpenTable)

If RS.RecordCount > 0 Then
'DataFound(1) = True
Do Until RS.EOF
   GoSub agedays
   GoSub adj205
   GoSub avgdgain
   GoSub wt2day
   If RS!skipme = False And RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
      'allwt = allwt + rs!adj205wt
      x = x + RS!age_in_days
   Else
      skipme = skipme + 1
   End If
   RS.MoveNext
Loop
If RS.RecordCount <> skipme Then
   AvgAge = x / (RS.RecordCount - skipme)

Set RS = DBREP.OpenRecordset("steercalf")
RS.MoveFirst
x = 0
IrrCalf = 0
Do Until RS.EOF
      'If rs!calfID = "7013" Then
      '   Stop
      'End If
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
AvgAge = x / allcalf
RS.MoveFirst
Do Until RS.EOF
   GoSub adj205r
   RS.MoveNext
Loop

RS.Close: Set RS = Nothing
DBREP.Close: Set DBREP = Nothing

End If
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
         If RS!age = 2 Then
            dam = 54
         End If
         If RS!age = 3 Then
            dam = 36
         End If
      If RS!age = 4 Then
         dam = 18
      End If
      If RS!age > 11 Then
         dam = 18
      End If
   End If
   If RS!Sex = 1 Or RS!Sex = 3 Then
      If RS!age = 2 Then
         dam = 60
      End If
      If RS!age = 3 Then
         dam = 40
      End If
      If RS!age = 4 Then
         dam = 20
      End If
      If RS!age > 11 Then
         dam = 20
      End If
   End If
   RS!adj205wt = (((RS!actweight - RS!birthwt) / RS!age_in_days) * 205) + RS!birthwt + dam
   End If
   RS.Update
Return

adj205r:
If group = False Then
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
Else
   GSql$ = "SELECT Sum(heical.adj205wt) AS allwt, Count(heical.CalfID) AS allcalf, heical.group From heical where (((heical.adj205wt) > 0)) GROUP BY heical.group"
   Set GRS = DBREP.OpenRecordset(GSql$, dbOpenSnapshot)
   If GRS.RecordCount > 0 Then
      GRS.FindFirst Field2Str(GRS!group) = Field2Str(RS!group)
      RS.Edit: RS!adj205rat = (RS!adj205wt / (GRS!allwt / GRS!allcalf)) * 100: RS.Update
   End If
   GRS.Close: Set GRS = Nothing
End If
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
ehandle:
If Err.Number = 94 Then Resume Next
End Sub

Public Sub CreateSireSum()
On Error GoTo ehandle
Dim DBREP As database, dbrs As Recordset
Dim SQL As String, sireid As String, calf As Double, wt205 As Double, birthwt As Double, CalfEase As Double
Dim actwt As Double, agedays As Double, frscor As Double, adg As Double, wda As Double, sire_breed As String
Dim varbookmark As Variant, skipme As Double, SkipRow As Integer, SortOrder$

Set DBREP = DBEngine(0).OpenDatabase(repfile$, False)
DBREP.Execute ("delete * from siresumtmp")
SQL = "insert into siresumtmp select * from calflist"
DBREP.Execute (SQL)
'SQL = "insert into siresumtmp select * from bulcalf"
'DBREP.Execute (SQL)
'SQL = "insert into siresumtmp select * from steercalf"
'DBREP.Execute (SQL)
'SQL = "insert into siresumtmp select * from misccalf"
'DBREP.Execute (SQL)
'QL = "insert into siresum select * from siresumtmp where skipme = false order by sireid, sire_breed"
DBREP.Execute ("delete * from siresum")
DBREP.Execute ("insert into siresum select * from siresumtmp where sire_breed <> ''"), dbFailOnError

Set dbrs = DBREP.OpenRecordset("siresum", dbOpenTable)

If dbrs.RecordCount > 0 Then
dbrs.MoveFirst
With dbrs
   sireid = !sireid
   sire_breed = !sire_breed
   Do Until .EOF
      varbookmark = .Bookmark
         wt205 = 0
         calf = 0
         birthwt = 0
         CalfEase = 0
         actwt = 0
         agedays = 0
         frscor = 0
         adg = 0
         wda = 0
         skipme = 0
NextRow:
   While !sireid = sireid And !sire_breed = sire_breed
         calf = calf + 1
            Select Case !Sex
               Case 1
                  wt205 = wt205 + (!adj205wt * 0.95)
               Case 2
                  wt205 = wt205 + (!adj205wt * 1.05)
               Case 3
                  wt205 = wt205 + !adj205wt
            End Select
            birthwt = birthwt + !birthwt
            CalfEase = CalfEase + !calvingease
            actwt = actwt + !actweight
            agedays = agedays + !age_in_days
            frscor = frscor + !cframe
            adg = adg + !avgdailygain
            wda = wda + !wt2daygain
         .MoveNext
        If .EOF Then GoTo save
      Wend
      
save:
      
      .Bookmark = varbookmark
       While !sireid = sireid And !sire_breed = sire_breed
         .Edit
         If calf > 0 Then
            !adj205wt = wt205 / (calf - skipme)
            !birthwt = birthwt / (calf - skipme)
            !calvingease = CalfEase / (calf - skipme)
            !actweight = actwt / (calf - skipme)
            !age_in_days = agedays / (calf - skipme)
            !cframe = frscor / (calf - skipme)
            !avgdailygain = adg / (calf - skipme)
            !wt2daygain = wda / (calf - skipme)
            !num = calf
         .Update
         End If
         .MoveNext
         If .EOF Then Exit Do
       Wend
      If Not .EOF Then
         sireid = !sireid
         sire_breed = !sire_breed
      End If
   Loop
End With

End If

dbrs.Close: Set dbrs = Nothing
DBREP.Close: Set DBREP = Nothing

Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub
Public Sub CreateBHeiAvg(SortOrder%, group As Boolean)
Dim SQL$, repdb As database, RepRS As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from bheiavg")
If group = True Then
   repdb.Execute ("insert into bheiavg SELECT bulcalf.cow_breed, Count(bulcalf.CalfID) AS Num, Sum(bulcalf.adj205wt) AS adj205wt, Sum(bulcalf.birthwt) AS birthwt, bulcalf.group, Sum(bulcalf.calvingease) AS calvingease, Sum(bulcalf.actweight) AS actweight, Sum(bulcalf.age_in_days) AS age_in_days, Sum(bulcalf.cframe) AS cframe, Sum(bulcalf.avgdailygain) AS avgdailygain, Sum(bulcalf.wt2daygain) AS wt2daygain From bulcalf GROUP BY bulcalf.cow_breed, bulcalf.group, bulcalf.skipme HAVING (((bulcalf.skipme)=False))"), dbFailOnError
Else
   repdb.Execute ("insert into bheiavg SELECT bulcalf.cow_breed, Count(bulcalf.CalfID) AS Num, Sum(bulcalf.adj205wt) AS adj205wt, Sum(bulcalf.birthwt) AS birthwt, Sum(bulcalf.calvingease) AS calvingease, Sum(bulcalf.actweight) AS actweight, Sum(bulcalf.age_in_days) AS age_in_days, Sum(bulcalf.cframe) AS cframe, Sum(bulcalf.avgdailygain) AS avgdailygain, Sum(bulcalf.wt2daygain) AS wt2daygain From bulcalf GROUP BY bulcalf.cow_breed, bulcalf.skipme HAVING (((bulcalf.skipme)=False))"), dbFailOnError
End If
Set RepRS = repdb.OpenRecordset("bheiavg", dbOpenTable)
If RepRS.RecordCount > 0 Then
   Do Until RepRS.EOF
      RepRS.Edit
      RepRS!adj205wt = Field2Num(RepRS!adj205wt) / RepRS!num
      RepRS!birthwt = Field2Num(RepRS!birthwt) / RepRS!num
      RepRS!calvingease = Field2Num(RepRS!calvingease) / RepRS!num
      RepRS!actweight = Field2Num(RepRS!actweight) / RepRS!num
      RepRS!age_in_days = Field2Num(RepRS!age_in_days) / RepRS!num
      RepRS!cframe = Field2Num(RepRS!cframe) / RepRS!num
      RepRS!avgdailygain = Field2Num(RepRS!avgdailygain) / RepRS!num
      RepRS!wt2daygain = Field2Num(RepRS!wt2daygain) / RepRS!num
      RepRS.Update
      RepRS.MoveNext
   Loop
End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
End Sub

Public Sub CreateCowSummary()
On Error GoTo ehandle
Dim repdb As database, ResRS As Recordset, SQL As String
Dim DB As database, RS As Recordset
Dim agedays As Double, allcalf(3) As Double, Sex(2) As Double, born(2) As Double
Dim notwd(2) As Double, totdays As Double, w2d As Double, bwt As Double, adj205 As Double, dam As Double
Dim G As Double, h As Double, j As Double, K As Double, L As Double, r As Double, y As Double, NextRow As Boolean
Dim pTestDate As Date
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
DB.Execute ("delete * from siresum")
DB.Execute ("delete * from siresumtmp")
DB.Execute ("insert into siresumtmp select * from calflist")
DB.Execute ("insert into siresum select * from siresumtmp order by sex")

Set RS = DB.OpenRecordset("siresum", dbOpenTable)
If RS.RecordCount > 0 Then
With RS

   '.MoveFirst
  Do Until RS.EOF
   Select Case !Sex
      Case 1
       If !managecode <> "A" And !managecode <> "B" Then
          Sex(0) = Sex(0) + 1
      End If
      If !managecode <> "A" And !managecode <> "B" And !actweight > 0 Then
         born(0) = born(0) + 1
      End If
      If !managecode = "X" Then
         notwd(0) = notwd(0) + 1
      End If
      Case 2
       If !managecode <> "A" And !managecode <> "B" Then
         Sex(1) = Sex(1) + 1
      End If
      If !managecode <> "A" And !managecode <> "B" And !actweight > 0 Then
         born(1) = born(1) + 1
      End If
      If !managecode = "X" Then
         notwd(1) = notwd(1) + 1
      End If
      Case 3
       If !managecode <> "A" And !managecode <> "B" Then
         Sex(2) = Sex(2) + 1
      End If
      If !managecode <> "A" And !managecode <> "B" And !actweight > 0 Then
         born(2) = born(2) + 1
      End If
      If !managecode = "X" Then
         notwd(2) = notwd(2) + 1
      End If
   End Select
   If !managecode = "F" Then
      F = F + 1
   End If
   If !managecode = "P" Then
      P = P + 1
   End If
   .MoveNext
   Loop
End With

If Calculated = False Then
   pTestDate = thirdcowdate - 285
Else
   pTestDate = ActTurnDate
End If


SQL = "SELECT DISTINCTROW cowprof.reasonculled, cowprof.dateculled"
SQL = SQL & " From cowprof WHERE (((cowprof.dateculled)>#" & pTestDate & "#-1 "
SQL = SQL & " And (cowprof.dateculled)<#" & pTestDate & "# + 366)) and herdid = '" & herdid & "'"
'sql = "select dateculled, reasonculled from cowprof where dateculled > #" & ptestdate - 1 & "#"
Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
Set RS = DB.OpenRecordset(SQL, dbOpenSnapshot)

With RS
   '.MoveFirst
    Do Until .EOF
      
      Select Case !reasonculled
      Case "G"
         G = G + 1
      Case "H"
         h = h + 1
      Case "J"
         j = j + 1
      Case "K"
         K = K + 1
      Case "L"
         L = L + 1
      Case "R"
         r = r + 1
      Case "Y"
         y = y + 1
      End Select
      
      .MoveNext
   Loop
End With
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
DB.Execute ("delete * from siresum")
'db.Execute ("insert into siresum select * from siresumtmp where skipme = false")
Set RS = DB.OpenRecordset("siresumtmp", dbOpenSnapshot)
With RS
  If .RecordCount > 0 Then
   '.MoveFirst
   Do Until .EOF
            If !wt2daygain > 0 And !managecode <> "A" And !managecode <> "B" And !actweight > 0 Then
               w2d = w2d + !wt2daygain
               allcalf(0) = allcalf(0) + 1
            End If
            If !age_in_days > 0 And !managecode <> "A" And !managecode <> "B" Then
                  agedays = agedays + !age_in_days
                  allcalf(3) = allcalf(3) + 1
            End If
            If !birthwt > 0 Then
               bwt = bwt + !birthwt
               allcalf(1) = allcalf(1) + 1
            End If
            If !adj205wt > 0 And !skipme = False Then
               allcalf(2) = allcalf(2) + 1
            Select Case !Sex
               Case 1
                  adj205 = adj205 + (!adj205wt * 0.95)
               Case 2
                  adj205 = adj205 + (!adj205wt * 1.05)
               Case 3
                  adj205 = adj205 + !adj205wt
            End Select
            End If
            NextRow = True
         If NextRow = True Then
            .MoveNext
         End If
    Loop
   End If
End With

Set ResRS = DB.OpenRecordset("cowcount", dbOpenTable)
     With ResRS
       .Edit
         !G = G
         !h = h
         !j = j
         !K = K
         !L = L
         !r = r
         !y = y
       .Update
    End With

DB.Execute ("delete * from herdcount2")
Set ResRS = DB.OpenRecordset("herdcount2", dbOpenTable)
With ResRS
   .AddNew
   'sd12 fix put the if on the four statements below.
   If allcalf(3) <> 0 Then !avgdays = agedays / allcalf(3)
   If allcalf(0) <> 0 Then !wtperday = w2d / allcalf(0)
   If allcalf(1) <> 0 Then !avgbirthwt = bwt / allcalf(1)
   If allcalf(2) <> 0 Then !avg205 = adj205 / allcalf(2)
   !calvesbornbulls = Sex(0)
   !calvesbornheifers = Sex(1)
   !calvesbornsteers = Sex(2)
   !calvesweighedbulls = born(0)
   !calvesweighedheifers = born(1)
   !calvesweighedsteers = born(2)
   !calvesnotweighedbulls = notwd(0)
   !calvesnotweighedheifers = notwd(1)
   !calvesnotweighedsteers = notwd(2)
   .Update
End With

RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
'ResRS.Close: Set ResRS = Nothing
'repdb.Close: Set repdb = Nothing
End If
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
If Err.Number = 3021 Then
   ResRS.AddNew
   Resume Next
End If
End Sub

Public Sub CreateCritSuccFactors()
On Error GoTo ehandle
Dim DB As database, RS As Recordset, table%, SQL As String
Dim agedays As Double, w2d As Double, birthwt As Double, adg As Double, wt205 As Double, Sex(2) As Double
Dim frscor As Double, allcalf(9) As Double, actwt(2) As Double, cowage As Double
Dim cowwnwt As Double, cowbrdcond As Double, RepCow As Double

Set DB = DBEngine(0).OpenDatabase(repfile$, False)
DB.Execute ("delete * from prefcsf")
Set DB = DBEngine(0).OpenDatabase(dbfile$, False)
SQL = "insert into prefcsf in '" & repfile$ & "'"
SQL = SQL & " select * from prefcsf"
DB.Execute (SQL)
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
DB.Execute ("delete * from siresumtmp")
DB.Execute ("delete * from siresum")

DB.Execute ("insert into siresumtmp select * from calflist")
Set RS = DB.OpenRecordset("siresumtmp", dbOpenSnapshot)
'If RS.RecordCount > 0 Then
   With RS
      Do Until .EOF
               'If !cowage = 2 And !managecode <> "A" And !managecode <> "B" Then
               '   RepCow = RepCow + 1
               'End If
               If !cowage < 3 Then RepCow = RepCow + 1
               If !managecode <> "A" And !managecode <> "B" And !avgdailygain > 0 And !birthwt > 0 And !actweight > 0 Then
                  adg = adg + !avgdailygain
                  allcalf(3) = allcalf(3) + 1
               End If
               If !cframe > 0 Then
                  frscor = frscor + !cframe
                  allcalf(5) = allcalf(5) + 1
               End If
               If !cowage > 0 Then
                  cowage = cowage + !cowage
                  allcalf(6) = allcalf(6) + 1
               End If
               If !actweight > 0 Then
                  Select Case !Sex
                     Case 0
                     Case 1
                        actwt(0) = actwt(0) + !actweight
                        Sex(0) = Sex(0) + 1
                     Case 2
                        actwt(1) = actwt(1) + !actweight
                        Sex(1) = Sex(1) + 1
                     Case 3
                        actwt(2) = actwt(2) + !actweight
                        Sex(2) = Sex(2) + 1
                  End Select
               End If
         .MoveNext
       Loop
   End With
Set DB = DBEngine(0).OpenDatabase(dbfile$, True)
Set RS = DB.OpenRecordset("select * from cowbrd where calfdate = #" & ActTurnDate & "#", dbOpenSnapshot)
With RS
   Do Until .EOF
      If !weanwt > 0 Then
         allcalf(8) = allcalf(8) + 1
         cowwnwt = cowwnwt + !weanwt
      End If
      If !weancond > 0 Then
         allcalf(9) = allcalf(9) + 1
         cowbrdcond = cowbrdcond + !weancond
      End If
      .MoveNext
   Loop
End With
Set DB = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DB.OpenRecordset("herdcount2")
With RS
   agedays = !avgdays
   w2d = !wtperday
   birthwt = !avgbirthwt
   wt205 = !avg205
End With
DB.Execute ("delete * from critsuccfac")
Set RS = DB.OpenRecordset("critsuccfac", dbOpenTable)
With RS
   .AddNew
      !cpt = agedays
      If w2d > 0 Then
         !wt2daygain = w2d
      End If
      !birthwt = birthwt
      If adg > 0 Then
         !adg = adg / allcalf(3)
      End If
      If cowage > 0 Then
         !avgcowage = cowage / allcalf(6)
      End If
      !wt205 = wt205
      If cowwnwt > 0 Then
         !cowwtwean = cowwnwt / allcalf(8)
      End If
      If cowbrdcond > 0 Then
         !cowconscore = cowbrdcond / allcalf(9)
      End If
      If actwt(0) > 0 And Sex(0) > 0 Then
         !actwtbul = actwt(0) / Sex(0)
      End If
      If actwt(1) > 0 And Sex(1) > 0 Then
         !actwtheif = actwt(1) / Sex(1)
      End If
      If actwt(2) > 0 And Sex(2) > 0 Then
         !actwtsteer = actwt(2) / Sex(2)
      End If
      If frscor > 0 Then
         !cframe = frscor / allcalf(5)
      End If
      !RepCalv = RepCow
   .Update
End With
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
'End If
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub

Public Sub CreateHeiAvg(SortOrder%, group As Boolean)
Dim SQL$, repdb As database, RepRS As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from heiavg")
If group = True Then
   repdb.Execute ("insert into heiavg SELECT heical.cow_breed, Count(heical.CalfID) AS Num, Sum(heical.adj205wt) AS adj205wt, Sum(heical.birthwt) AS birthwt, heical.group, Sum(heical.calvingease) AS calvingease, Sum(heical.actweight) AS actweight, Sum(heical.age_in_days) AS age_in_days, Sum(heical.cframe) AS cframe, Sum(heical.avgdailygain) AS avgdailygain, Sum(heical.wt2daygain) AS wt2daygain From heical GROUP BY heical.cow_breed, heical.group, heical.skipme HAVING (((heical.skipme)=False))"), dbFailOnError
Else
   repdb.Execute ("insert into heiavg SELECT heical.cow_breed, Count(heical.CalfID) AS Num, Sum(heical.adj205wt) AS adj205wt, Sum(heical.birthwt) AS birthwt, Sum(heical.calvingease) AS calvingease, Sum(heical.actweight) AS actweight, Sum(heical.age_in_days) AS age_in_days, Sum(heical.cframe) AS cframe, Sum(heical.avgdailygain) AS avgdailygain, Sum(heical.wt2daygain) AS wt2daygain From heical GROUP BY heical.cow_breed, heical.skipme HAVING (((heical.skipme)=False))"), dbFailOnError
End If
Set RepRS = repdb.OpenRecordset("heiavg", dbOpenTable)
If RepRS.RecordCount > 0 Then
   Do Until RepRS.EOF
      RepRS.Edit
      RepRS!adj205wt = Field2Str(RepRS!adj205wt) / RepRS!num
      RepRS!birthwt = Field2Str(RepRS!birthwt) / RepRS!num
      RepRS!calvingease = Field2Str(RepRS!calvingease) / RepRS!num
      RepRS!actweight = Field2Str(RepRS!actweight) / RepRS!num
      RepRS!age_in_days = Field2Str(RepRS!age_in_days) / RepRS!num
      RepRS!cframe = Field2Str(RepRS!cframe) / RepRS!num
      RepRS!avgdailygain = Field2Str(RepRS!avgdailygain) / RepRS!num
      RepRS!wt2daygain = Field2Str(RepRS!wt2daygain) / RepRS!num
      RepRS.Update
      RepRS.MoveNext
   Loop
End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
End Sub
Public Sub CreateHBulAvg(SortOrder%, group As Boolean)
Dim repdb As database, RepRS As Recordset, SQL$
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from hbulavg")
SQL$ = "insert into hbulavg "
If group = True Then
   SQL$ = SQL$ & " SELECT DISTINCTROW HeiCal.SireID, HeiCal.sire_breed, Count(HeiCal.CalfID) AS num, Sum(HeiCal.adj205wt) AS adj205wt, Sum(HeiCal.birthwt) AS birthwt, Sum(HeiCal.calvingease) AS calvingease, Sum(HeiCal.actweight) AS actweight, Sum(HeiCal.age_in_days) AS age_in_days, Sum(HeiCal.avgdailygain) AS avgdailygain, Sum(HeiCal.wt2daygain) AS wt2daygain, HeiCal.group From HeiCal"
   SQL$ = SQL$ & " where (((HeiCal.skipme) = False)) GROUP BY HeiCal.SireID, HeiCal.sire_breed, HeiCal.group"
Else
   SQL$ = SQL$ & " SELECT DISTINCTROW HeiCal.SireID, HeiCal.sire_breed, Count(HeiCal.CalfID) AS num, Sum(HeiCal.adj205wt) AS adj205wt, Sum(HeiCal.birthwt) AS birthwt, Sum(HeiCal.calvingease) AS calvingease, Sum(HeiCal.actweight) AS actweight, Sum(HeiCal.age_in_days) AS age_in_days, Sum(HeiCal.avgdailygain) AS avgdailygain, Sum(HeiCal.wt2daygain) AS wt2daygain"
   SQL$ = SQL$ & " From HeiCal where (((HeiCal.skipme) = False)) GROUP BY HeiCal.SireID, HeiCal.sire_breed"
End If
repdb.Execute (SQL$), dbFailOnError
Set RepRS = repdb.OpenRecordset("hbulavg", dbOpenTable)
If RepRS.RecordCount > 0 Then
Do Until RepRS.EOF
   RepRS.Edit
   RepRS!adj205wt = Field2Num(RepRS!adj205wt) / RepRS!num
   RepRS!birthwt = Field2Num(RepRS!birthwt) / RepRS!num
   RepRS!calvingease = Field2Num(RepRS!calvingease) / RepRS!num
   RepRS!actweight = Field2Num(RepRS!actweight) / RepRS!num
   RepRS!age_in_days = Field2Num(RepRS!age_in_days) / RepRS!num
   RepRS!cframe = Field2Num(RepRS!cframe) / RepRS!num
   RepRS!avgdailygain = Field2Num(RepRS!avgdailygain) / RepRS!num
   RepRS!wt2daygain = Field2Num(RepRS!wt2daygain) / RepRS!num
   RepRS.Update
   RepRS.MoveNext
Loop
End If
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
End Sub
Public Sub CreateHeiCalf(SortOrder As Integer, group As Boolean)
On Error GoTo ehandle
Dim RS As Recordset, strSQL$, GSql$, GRS As Recordset
Dim DB As database, DBREP As database
Dim SQL As String, agedays As Boolean, adj205 As Boolean, adj205r As Boolean, avgdgain As Boolean, wt2day As Boolean
Dim dam As Double, allcalf As Double, AvgAge As Double, x As Double, allwt As Double, IrrCalf As Double, y As Double
Dim wt205 As Double, birthwt As Double, cease As Double, frscor As Double, adg As Double, wdg As Double
Dim skipme As Double
'Dim formula$, Col(1), fieldvar$(1), hmfields%, irrcalfwt As Double

Call CreateCalfTables(2, group)
If DataFound(2) = False Then Exit Sub
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)

Set RS = DBREP.OpenRecordset("heical", dbOpenTable)

If RS.RecordCount > 0 Then
'DataFound(2) = True
Do Until RS.EOF
   GoSub agedays
   GoSub adj205
   GoSub avgdgain
   GoSub wt2day
   If RS!skipme = False And RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
      'allwt = allwt + rs!adj205wt
      x = x + RS!age_in_days
   Else
      skipme = skipme + 1
   End If
   RS.MoveNext
Loop
If RS.RecordCount <> skipme Then
   AvgAge = x / (RS.RecordCount - skipme)

Set RS = DBREP.OpenRecordset("heical")
RS.MoveFirst
x = 0
IrrCalf = 0
Do Until RS.EOF
      'If rs!calfID = "7013" Then
      '   Stop
      'End If
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
AvgAge = x / allcalf
RS.MoveFirst
Do Until RS.EOF
   GoSub adj205r
   RS.MoveNext
Loop

RS.Close: Set RS = Nothing
DBREP.Close: Set DBREP = Nothing
End If
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
         If RS!age = 2 Then
            dam = 54
         End If
         If RS!age = 3 Then
            dam = 36
         End If
      If RS!age = 4 Then
         dam = 18
      End If
      If RS!age > 11 Then
         dam = 18
      End If
   End If
   If RS!Sex = 1 Or RS!Sex = 3 Then
      If RS!age = 2 Then
         dam = 60
      End If
      If RS!age = 3 Then
         dam = 40
      End If
      If RS!age = 4 Then
         dam = 20
      End If
      If RS!age > 11 Then
         dam = 20
      End If
   End If
   RS!adj205wt = (((RS!actweight - RS!birthwt) / RS!age_in_days) * 205) + RS!birthwt + dam
   End If
   RS.Update
Return

adj205r:
If group = False Then
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
Else
   GSql$ = "SELECT Sum(heical.adj205wt) AS allwt, Count(heical.CalfID) AS allcalf, heical.group From heical where (((heical.adj205wt) > 0)) GROUP BY heical.group"
   Set GRS = DBREP.OpenRecordset(GSql$, dbOpenSnapshot)
   If GRS.RecordCount > 0 Then
      GRS.FindFirst Field2Str(GRS!group) = Field2Str(RS!group)
      RS.Edit: RS!adj205rat = (RS!adj205wt / (GRS!allwt / GRS!allcalf)) * 100: RS.Update
   End If
   GRS.Close: Set GRS = Nothing
End If
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
ehandle:
If Err.Number = 94 Then Resume Next
End Sub
Public Sub CreateBulSum(group As Boolean)
On Local Error GoTo ehandle
Dim repdb As database, RepRSData As Recordset, RepRSEdit As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from bulsum")
If group = True Then
   'Get all group identifiers for bulsum group averages
   repdb.Execute ("insert into bulsum select group from heical group by group order by group")
 If repdb.RecordsAffected > 0 Then
   repdb.Execute ("update bulsum set bulsum.group = 0 where isnull(group)")
   repdb.Execute ("update heical set heical.group = 0 where isnull(group)")
   Set RepRSEdit = repdb.OpenRecordset("bulsum", dbOpenTable)
   'update calvingease
   Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num, heical.group From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4)) GROUP BY heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!cease = RepRSData!cease / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num, heical.group from heical where heical.birthwt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num, heical.group from heical where heical.adj205wt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num, heical.group from heical where heical.actweight > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num, heical.group from heical where heical.age_in_days > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num, heical.group from heical where heical.cframe > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!frscor = RepRSData!frscor / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num, heical.group from heical where heical.avgdailygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!adg = RepRSData!adg / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num, heical.group from heical where heical.wt2daygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wdg = RepRSData!wda / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
 End If
Else
  Set RepRSEdit = repdb.OpenRecordset("bulsum", dbOpenTable)
  RepRSEdit.AddNew
   'update calvingease
  Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4))", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!cease = RepRSData!cease / RepRSData!num
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num from heical where heical.birthwt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num from heical where heical.adj205wt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num from heical where heical.actweight > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num from heical where heical.age_in_days > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num from heical where heical.cframe > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!frscore = RepRSData!frscor / RepRSData!num
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num from heical where heical.avgdailygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!adg = RepRSData!adg / RepRSData!num
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num from heical where heical.wt2daygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wdg = RepRSData!wda / RepRSData!num
   RepRSEdit.Update
 End If
RepRSData.Close: Set RepRSData = Nothing
RepRSEdit.Close: Set RepRSEdit = Nothing
repdb.Close: Set repdb = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next Else RepRSData.Close: Set RepRSData = Nothing: RepRSEdit.Close: Set RepRSEdit = Nothing: repdb.Close: Set repdb = Nothing: MsgBox Err.Description

End Sub
Public Sub CreateBulCalf(group As Boolean)
On Error GoTo ehandle
Dim RS As Recordset, GRS As Recordset, GSql$
Dim DB As database, DBREP As database
Dim SQL As String, agedays As Boolean, adj205 As Boolean, adj205r As Boolean, avgdgain As Boolean, wt2day As Boolean
Dim dam As Double, allcalf As Double, AvgAge As Double, x As Double, allwt As Double, IrrCalf As Double, y As Double
Dim wt205 As Double, birthwt As Double, cease As Double, frscor As Double, adg As Double, wdg As Double
Dim skipme As Double
'Dim formula$, Col(1), fieldvar$(1), hmfields%, irrcalfwt As Double
Call CreateCalfTables(3, False)
If DataFound(3) = False Then Exit Sub
Set DBREP = DBEngine(0).OpenDatabase(repfile$, False)
Set RS = DBREP.OpenRecordset("bulcalf", dbOpenTable)

If RS.RecordCount > 0 Then
'DataFound(3) = True
Do Until RS.EOF
   GoSub agedays
   GoSub adj205
   GoSub avgdgain
   GoSub wt2day
   If RS!skipme = False And RS!managecode <> "A" And RS!managecode <> "B" And RS!managecode <> "C" And RS!managecode <> "D" And RS!managecode <> "E" And RS!managecode <> "F" And RS!managecode <> "K" And RS!managecode <> "N" And RS!managecode <> "P" And RS!managecode <> "S" And RS!managecode <> "T" And RS!managecode <> "X" Then
      'allwt = allwt + rs!adj205wt
      x = x + RS!age_in_days
   Else
      skipme = skipme + 1
   End If
   RS.MoveNext
Loop
If RS.RecordCount <> skipme Then
   AvgAge = x / (RS.RecordCount - skipme)

Set RS = DBREP.OpenRecordset("bulcalf")
RS.MoveFirst
x = 0
IrrCalf = 0
Do Until RS.EOF
      'If rs!calfID = "7013" Then
      '   Stop
      'End If
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
AvgAge = x / allcalf
RS.MoveFirst
Do Until RS.EOF
   GoSub adj205r
   RS.MoveNext
Loop

RS.Close: Set RS = Nothing
DBREP.Close: Set DBREP = Nothing

End If
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
         If RS!age = 2 Then
            dam = 54
         End If
         If RS!age = 3 Then
            dam = 36
         End If
      If RS!age = 4 Then
         dam = 18
      End If
      If RS!age > 11 Then
         dam = 18
      End If
   End If
   If RS!Sex = 1 Or RS!Sex = 3 Then
      If RS!age = 2 Then
         dam = 60
      End If
      If RS!age = 3 Then
         dam = 40
      End If
      If RS!age = 4 Then
         dam = 20
      End If
      If RS!age > 11 Then
         dam = 20
      End If
   End If
   RS!adj205wt = (((RS!actweight - RS!birthwt) / RS!age_in_days) * 205) + RS!birthwt + dam
   End If
   RS.Update
Return

adj205r:
If group = False Then
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
Else
   GSql$ = "SELECT Sum(heical.adj205wt) AS allwt, Count(heical.CalfID) AS allcalf, heical.group From heical where (((heical.adj205wt) > 0)) GROUP BY heical.group"
   Set GRS = DBREP.OpenRecordset(GSql$, dbOpenSnapshot)
   If GRS.RecordCount > 0 Then
      GRS.FindFirst Field2Str(GRS!group) = Field2Str(RS!group)
      RS.Edit: RS!adj205rat = (RS!adj205wt / (GRS!allwt / GRS!allcalf)) * 100: RS.Update
   End If
   GRS.Close: Set GRS = Nothing
End If
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
ehandle:
If Err.Number = 94 Then Resume Next
End Sub
Public Sub CreateHeiSum(group As Boolean)
On Local Error GoTo ehandle
Dim repdb As database, RepRSData As Recordset, RepRSEdit As Recordset
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from heisum")
If group = True Then
   'Get all group identifiers for heisum group averages
   repdb.Execute ("insert into heisum select group from heical group by group order by group")
 If repdb.RecordsAffected > 0 Then
   repdb.Execute ("update heisum set heisum.group = 0 where isnull(group)")
   repdb.Execute ("update heical set heical.group = 0 where isnull(group)")
   Set RepRSEdit = repdb.OpenRecordset("heisum", dbOpenTable)
   'update calvingease
   Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num, heical.group From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4)) GROUP BY heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!cease = RepRSData!cease / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num, heical.group from heical where heical.birthwt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num, heical.group from heical where heical.adj205wt > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num, heical.group from heical where heical.actweight > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num, heical.group from heical where heical.age_in_days > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num, heical.group from heical where heical.cframe > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!frscor = RepRSData!frscor / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num, heical.group from heical where heical.avgdailygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!adg = RepRSData!adg / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num, heical.group from heical where heical.wt2daygain > 0 group by heical.group order by heical.group", dbOpenSnapshot)
   If RepRSData.RecordCount > 0 Then
   RepRSEdit.MoveFirst
   Do Until RepRSData.EOF
      'RepRsEdit.FindFirst RepRsEdit!Group = RepRsData!Group
      RepRSEdit.Edit
      RepRSEdit!wdg = RepRSData!wda / RepRSData!num
      RepRSEdit.Update
      RepRSData.MoveNext
      RepRSEdit.MoveNext
   Loop
   End If
 End If
Else
  Set RepRSEdit = repdb.OpenRecordset("heisum", dbOpenTable)
  RepRSEdit.AddNew
   'update calvingease
  Set RepRSData = repdb.OpenRecordset("SELECT Sum(heical.calvingease) AS cease, Count(heical.CalfID) AS num From heical where (((heical.calvingease) >= 0 And (heical.calvingease) <= 4))", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!cease = RepRSData!cease / RepRSData!num
   'update birthwt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.birthwt) as birthwt, count(heical.calfid) as num from heical where heical.birthwt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!birthwt = RepRSData!birthwt / RepRSData!num
   'update 205 wt
   Set RepRSData = repdb.OpenRecordset("select sum(heical.adj205wt) as adj205, count(heical.calfid) as num from heical where heical.adj205wt > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wt205 = RepRSData!adj205 / RepRSData!num
   'update actual wean weight
   Set RepRSData = repdb.OpenRecordset("select sum(heical.actweight) as actweight, count(heical.calfid) as num from heical where heical.actweight > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!allwt = RepRSData!actweight / RepRSData!num
   'update age-days
   Set RepRSData = repdb.OpenRecordset("select sum(heical.age_in_days) as age_days, count(heical.calfid) as num from heical where heical.age_in_days > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!AvgAge = RepRSData!age_days / RepRSData!num
   'update frame score
   Set RepRSData = repdb.OpenRecordset("select sum(heical.cframe) as frscor, count(calfid) as num from heical where heical.cframe > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!frscore = RepRSData!frscor / RepRSData!num
   'update ADG
   Set RepRSData = repdb.OpenRecordset("select sum(heical.avgdailygain) as adg, count(calfid) as num from heical where heical.avgdailygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!adg = RepRSData!adg / RepRSData!num
   'update WDA
   Set RepRSData = repdb.OpenRecordset("select sum(heical.wt2daygain) as wda,count(calfid) as num from heical where heical.wt2daygain > 0", dbOpenSnapshot)
   If Not IsNull(RepRSData!num) Then RepRSEdit!wdg = RepRSData!wda / RepRSData!num
   RepRSEdit.Update
 End If
RepRSData.Close: Set RepRSData = Nothing
RepRSEdit.Close: Set RepRSEdit = Nothing
repdb.Close: Set repdb = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next Else RepRSData.Close: Set RepRSData = Nothing: RepRSEdit.Close: Set RepRSEdit = Nothing: repdb.Close: Set repdb = Nothing: MsgBox Err.Description
End Sub



Public Sub CreateCalfDistReport(thirdcowdate As Date, avgdate As Double, rowcount As Double, avgwgt As Double, calf As Double)
    On Error GoTo ehandle
    Dim repdb As database, ResRS As Recordset
    Dim DB As database, RS As Recordset
    Dim TCA As Double, TCO As Double, TlC As Double, allcalf As Double
    Dim TCEXP As Double, tcc As Double, CWC As Double, DamAge As Double
    Dim SQL As String, age As Double, period(5) As Double, exitcode%, SaveCode%
    Dim codecount As Double, countme As Boolean, SaveData As Integer, no_calves As Double, avgwt As Double
    Dim table%, avgbdate As Double, age12 As Boolean
    Dim pTestDate As Date, pTestDate2 As Date
    DB = DBEngine(0).OpenDatabase(dbfile$, False)
    repdb = DBEngine(0).OpenDatabase(repfile$, False)
    'if no third cow date then calculate test date for calving distribution table
    'If BoolThirdCowDate = False Then pTestDate = TurnDate + 285 Else pTestDate = thirdcowdate
    If Calculated = True Then pTestDate = ActTurnDate Else pTestDate = thirdcowdate - 285

    pTestDate2 = mTestDate
    TCE = 0
    repdb.Execute("delete * from siresumtmp")
    repdb.Execute("delete * from siresum")
    'repdb.Execute ("insert into siresumtmp select * from heical")
    'repdb.Execute ("insert into siresumtmp select * from bulcalf")
    'repdb.Execute ("insert into siresumtmp select * from misccalf")
    'repdb.Execute ("insert into siresumtmp select * from steercalf")
    repdb.Execute("insert into siresumtmp select * from calflist")
    RS = repdb.OpenRecordset("siresumtmp", dbOpenSnapshot)
    If RS.RecordCount > 0 Then
        With RS
            '.MoveFirst
            Do Until .EOF
                If !managecode = "C" Or !managecode = "D" Or !managecode = "F" Or !managecode = "K" Then
                    TlC = TlC + 1 'Total Cows losing calf
                End If
                If !managecode = "B" Then
                    TCA = TCA + 1 'Total Cows aborted
                End If
                If !managecode = "A" Then
                    TCO = TCO + 1 'Total cows open
                End If
                If !managecode = "T" Then TCE = TCE + 0.5 Else TCE = TCE + 1
                .MoveNext()
            Loop
        End With

        'TCE = RS.RecordCount 'Total Cows kept for calving

        RS.Close() : RS = Nothing

        'SQL = "select dateculled from cowprof where dateculled > #" & TurnDate - 1 & "#"
        'SQL = "SELECT DISTINCTROW cowprof.reasonculled, cowprof.dateculled"
        'SQL = SQL & " From cowprof WHERE (((cowprof.dateculled)>#" & pTestDate2 & "#-1 "
        'SQL = SQL & " And (cowprof.dateculled)<#" & pTestDate2 & "# + 366)) and cowprof.herdid = '" & herdid & "'"
        SQL = "SELECT DISTINCTROW Count(cowprof.dateculled) AS CowCount From cowprof WHERE (((cowprof.dateculled) Between #" & pTestDate & "# And #" & pTestDate + 365 & "#)) GROUP BY cowprof.HerdID HAVING (((cowprof.HerdID)='" & herdid & "'))"
        RS = DB.OpenRecordset(SQL)
        If Not RS.EOF Then
            TCEXP = Field2Num(RS!cowcount) + TCE 'Total Cows Exposed
        Else
            TCEXP = TCE
        End If

        tcc = TCE - TCA - TCO 'number of cows calving
        CWC = TCE - TCA - TCO - TlC 'number of cows weaning calves

        SaveData = 2
        exitcode% = 2
GoSub SaveData

        RS.Close() : RS = Nothing
        repdb.Execute("insert into siresum select * from siresumtmp order by cowage ")
        SQL = "SELECT DISTINCTROW SireSum.CowID, SireSum.birthdate From SireSum WHERE (((SireSum.birthdate)=#" & thirdcowdate & "#))"
        RS = repdb.OpenRecordset(SQL)

        RS = repdb.OpenRecordset("siresum", dbOpenTable)
        SaveCode% = 0

        If RS.RecordCount > 0 Then
            With RS
                '.MoveFirst
                DamAge = !cowage
                Do While !cowage < 12
                    Do While !cowage = DamAge And !cowage < 12
                        If !managecode = "A" Or !managecode = "B" Then
                            codecount = codecount + 1
                        End If
                        If !managecode <> "A" And !managecode <> "B" Then
                            no_calves = no_calves + 1
                            countme = True
                            If !actweight > 0 Then
                                avgwt = avgwt + !actweight
                                allcalf = allcalf + 1
                            End If
                            If !birthdate < pTestDate2 Then
                                period(0) = period(0) + 1
                                countme = False
                            End If
                            If !birthdate >= pTestDate2 And !birthdate <= pTestDate2 + 20 Then
                                period(1) = period(1) + 1
                                countme = False
                            End If
                            If !birthdate >= pTestDate2 + 21 And !birthdate <= pTestDate2 + 41 Then
                                period(2) = period(2) + 1
                                countme = False
                            End If
                            If !birthdate >= pTestDate2 + 42 And !birthdate <= pTestDate2 + 62 Then
                                period(3) = period(3) + 1
                                countme = False
                            End If
                            If !birthdate >= pTestDate2 + 63 And !birthdate <= pTestDate2 + 83 Then
                                period(4) = period(4) + 1
                                countme = False
                            End If
                            If !birthdate > pTestDate2 + 84 Then
                                period(5) = period(5) + 1
                                countme = False
                            End If
                            'If countme = False Then
                            age = age + !birthdate
                            'Debug.Print age & "       " & !birthdate
                            'End If
                            .MoveNext()
                            If Not .EOF Then
                                'If age12 = False Then
                                If !cowage <> DamAge Then
                                    SaveData = 1
                                    exitcode% = 2
            GoSub SaveData
                                    DamAge = !cowage
                                End If
                                'Else
                                '    .MoveNext
                                'End If
                            Else
                                SaveData = 1
                                exitcode% = 1
         GoSub SaveData
         GoSub CloseDB
                            End If
                        Else
                            .MoveNext()
                            If Not .EOF Then
                                'If age12 = False Then
                                If !cowage <> DamAge Then
                                    SaveData = 1
                                    exitcode% = 2
               GoSub SaveData
                                    DamAge = !cowage
                                End If
                                'Else
                                '   .MoveNext
                                'End If
                            Else
                                SaveData = 1
                                exitcode% = 1
         GoSub SaveData
         GoSub CloseDB
                            End If
                        End If
                    Loop
                Loop
            End With
        Else
   GoSub CloseDB
        End If
    End If
GoSub CloseDB
    'avg birthdate, third cow calving string and date are passed to createavgactwnwt for report.setformula
    'Call CreateCalfDistReport_12(thirdcowdate, age(1), TCC, avgwt(1), allcalf(1))
    Exit Sub

CloseDB:
    RS.Close() : RS = Nothing
    DB.Close() : DB = Nothing
    If exitcode% = 1 Then Exit Sub
    Return

SaveData:

    Select Case SaveData
        Case 1
            If no_calves > 0 Then
                'Set repdb = DBEngine(0).OpenDatabase(repfile$, , False)
                ResRS = repdb.OpenRecordset("herdcount", dbOpenTable)
                If SaveCode% = 0 Then
                    repdb.Execute("delete * from herdcount")
                    SaveCode% = 1
                End If
                With ResRS
                    .AddNew()
                    !DamAge = DamAge
                    !noofcalves = no_calves
                    !period0 = period(0)
                    !period1 = period(1)
                    !period2 = period(2)
                    !period3 = period(3)
                    !period4 = period(4)
                    !period5 = period(5)
                    !openab = codecount
                    If no_calves > 0 Then
                        !avgdate = age / no_calves
                    End If
                    avgdate = avgdate + age
                    If allcalf > 0 Then
                        !avgwt = avgwt / allcalf
                    End If
                    calf = calf + allcalf
                    avgwgt = avgwgt + avgwt
                    rowcount = rowcount + no_calves
                    .Update()
                End With


                ResRS.Close() : ResRS = Nothing
                'repdb.Close: Set repdb = Nothing

                allcalf = 0
                age = 0
                avgwt = 0
                period(0) = 0
                period(1) = 0
                period(2) = 0
                period(3) = 0
                period(4) = 0
                period(5) = 0
                codecount = 0
                no_calves = 0
            End If
        Case 2
            'Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
            ResRS = repdb.OpenRecordset("cowcount", dbOpenTable)
            With ResRS
                If Not ResRS.EOF Then
                    .Edit()
                    !TCEXP = TCEXP
                    !TCE = TCE
                    !TCA = TCA
                    !TCO = TCO
                    !tcc = tcc
                    !TCL = TlC
                    !CWC = CWC
                    .Update()
                End If
            End With
            ResRS.Close() : ResRS = Nothing
            'repdb.Close: Set repdb = Nothing
    End Select
    'If exitcode% = 1 Then
    'If table% <> 1 Then
    'avg birthdate, third cow calving string and date are passed to createavgactwnwt for report.setformula
    'Call CreateAvgActWnWt(thirdcowdate, age(1), TCC, avgwt(1), allcalf(1))
    'Exit Sub
    'Else
    'table% = 0
    'GoTo Table
    'End If
    'End If
    Return
    Exit Sub
ehandle:
    If Err.Number = 94 Then Resume Next
End Sub

Public Sub CreateBulAvg(SortOrder%, group As Boolean)
    Dim repdb As database, RepRS As Recordset, SQL$
    repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
    repdb.Execute("delete * from bulavg")
    SQL$ = "insert into bulavg "
    If group = True Then
        SQL$ = SQL$ & " SELECT DISTINCTROW bulcalf.SireID, bulcalf.sire_breed, Count(bulcalf.CalfID) AS num, Sum(bulcalf.adj205wt) AS adj205wt, Sum(bulcalf.birthwt) AS birthwt, Sum(bulcalf.calvingease) AS calvingease, Sum(bulcalf.actweight) AS actweight, Sum(bulcalf.age_in_days) AS age_in_days, Sum(bulcalf.avgdailygain) AS avgdailygain, Sum(bulcalf.wt2daygain) AS wt2daygain, bulcalf.group From bulcalf"
        SQL$ = SQL$ & " where (((bulcalf.skipme) = False)) GROUP BY bulcalf.SireID, bulcalf.sire_breed, bulcalf.group"
    Else
        SQL$ = SQL$ & " SELECT DISTINCTROW bulcalf.SireID, bulcalf.sire_breed, Count(bulcalf.CalfID) AS num, Sum(bulcalf.adj205wt) AS adj205wt, Sum(bulcalf.birthwt) AS birthwt, Sum(bulcalf.calvingease) AS calvingease, Sum(bulcalf.actweight) AS actweight, Sum(bulcalf.age_in_days) AS age_in_days, Sum(bulcalf.avgdailygain) AS avgdailygain, Sum(bulcalf.wt2daygain) AS wt2daygain"
        SQL$ = SQL$ & " From bulcalf where (((bulcalf.skipme) = False)) GROUP BY bulcalf.SireID, bulcalf.sire_breed"
    End If
repdb.Execute (SQL$), dbFailOnError
    RepRS = repdb.OpenRecordset("bulavg", dbOpenTable)
    If RepRS.RecordCount > 0 Then
        Do Until RepRS.EOF
            RepRS.Edit()
            RepRS!adj205wt = Field2Num(RepRS!adj205wt) / RepRS!num
            RepRS!birthwt = Field2Num(RepRS!birthwt) / RepRS!num
            RepRS!calvingease = Field2Num(RepRS!calvingease) / RepRS!num
            RepRS!actweight = Field2Num(RepRS!actweight) / RepRS!num
            RepRS!age_in_days = Field2Num(RepRS!age_in_days) / RepRS!num
            RepRS!cframe = Field2Num(RepRS!cframe) / RepRS!num
            RepRS!avgdailygain = Field2Num(RepRS!avgdailygain) / RepRS!num
            RepRS!wt2daygain = Field2Num(RepRS!wt2daygain) / RepRS!num
            RepRS.Update()
            RepRS.MoveNext()
        Loop
    End If
    RepRS.Close() : RepRS = Nothing
    repdb.Close() : repdb = Nothing
End Sub



Public Sub CreatePerCalv()
On Error GoTo ehandle
Dim RS As Recordset, DB As database
Dim allcalf As Double, allcalf2 As Double, period(5) As Double, period2(5) As Double
Set DB = DBEngine(0).OpenDatabase(repfile$, False, False)
Set RS = DB.OpenRecordset("herdcount", dbOpenSnapshot)
If RS.RecordCount > 0 Then
With RS
   .MoveFirst
   Do Until .EOF
      If !DamAge = 2 Then
         allcalf2 = allcalf2 + !noofcalves
         period(0) = period(0) + !period0
         period(1) = period(1) + !period1
         period(2) = period(2) + !period2
         period(3) = period(3) + !period3
         period(4) = period(4) + !period4
         period(5) = period(5) + !period5
      End If
         
         allcalf = allcalf + !noofcalves
         period2(0) = period2(0) + !period0
         period2(1) = period2(1) + !period1
         period2(2) = period2(2) + !period2
         period2(3) = period2(3) + !period3
         period2(4) = period2(4) + !period4
         period2(5) = period2(5) + !period5
      
      .MoveNext
   Loop
End With
Set RS = DB.OpenRecordset("critsuccfac", dbOpenTable)
With RS
   .Edit
      'allcalf = Period(0) + Period(1) + Period(2) + Period(3) + Period(4) + Period(5)
      'allcalf2 = period2(1) + period2(2) + period2(3) + period2(4) + period2(5)
      !percalvearly = (period(0) / allcalf2) * 100
      !percalv21 = ((period(0) + period(1)) / allcalf2) * 100
      !percalv42 = ((period(2) + period(0) + period(1)) / allcalf2) * 100
      allcalf = allcalf - allcalf2
      If allcalf <> 0 Then
        !permatcalv21 = ((period2(0) + period2(1) - period(0) - period(1)) / allcalf) * 100
        !permatcalv42 = ((period2(0) + period2(1) + period2(2) - period(0) - period(1) - period(2)) / allcalf) * 100
      Else
        !permatcalv21 = 0
        !permatcalv42 = 0
      End If
   .Update
End With
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End If
Exit Sub
ehandle:
Resume Next
End Sub


Public Sub CreateRepPerf()
On Error GoTo ehandle
Dim SQL As String, DB As database, RS  As Recordset, repdb As database, RepRS As Recordset
Dim PLP As Double, PregPer As Double, CP As Double, CDL As Double, WeanPer As Double, frrp(1) As Double
Dim CalfDeath As Double, allcalf(2) As Double, period(5) As Double, allwt As Double
Dim actwt As Double, CalfBorn As Double, ActWtCount As Double, x As Double, A As Double, b As Double
Dim denom2 As Double, pTestDate As Date, pTestDate2  As Date

'If BoolThirdCowDate = False Then pTestDate = TurnDate + 285 Else pTestDate = thirdcowdate
'If Calculated = True Then pTestDate2 = ActTurnDate + 285 Else pTestDate2 = thirdcowdate - 285
pTestDate2 = mTestDate
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from prefspa")
SQL = "insert into prefspa in '" & repfile$ & "'"
SQL = SQL & " select * from prefspa"
DB.Execute (SQL)
F = 0
'P = 0
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
Set RS = repdb.OpenRecordset("cowcount", dbOpenSnapshot)

With RS
   denom = !TCEXP - !h - !j - !L - !r - !y 'Denominator for Pregper, Calving Percentage, Calf Death Loss, Calf Crop or Wean Per, Replacement Rate
   PLP = !TCA / (!TCE - !TCO) * 100
   PregPer = ((!TCE - !TCO) / denom) * 100
   CP = ((!TCE - !TCA - !TCO) / denom) * 100
   CalfDeath = !TCL
   WeanPer = !TCEXP
End With

Set RS = repdb.OpenRecordset("siresumtmp", dbOpenTable)

With RS
   Do Until .EOF
      If !managecode <> "A" And !managecode <> "B" And !actweight > 0 Then
         allwt = allwt + !actweight
         allcalf(1) = allcalf(1) + 1
      End If
      If !actweight > 0 Then
         actwt = actwt + !actweight
         allcalf(2) = allcalf(2) + 1
      End If
   .MoveNext
   Loop
End With

Set RS = repdb.OpenRecordset("siresumtmp", dbOpenSnapshot)

With RS
   Do Until .EOF
      If !managecode = "A" Then
         A = A + 1
      End If
      If !managecode = "B" Then
         b = b + 1
      End If
      If !managecode <> "A" And !managecode <> "B" Then
      If !birthdate < pTestDate2 Then
         period(0) = period(0) + 1

      End If
      If !birthdate >= pTestDate2 And !birthdate <= pTestDate2 + 20 Then
         period(1) = period(1) + 1

      End If
      If !birthdate >= pTestDate2 + 21 And !birthdate <= pTestDate2 + 41 Then
         period(2) = period(2) + 1
         
      End If
      If !birthdate >= pTestDate2 + 42 And !birthdate <= pTestDate2 + 62 Then
         period(3) = period(3) + 1
         
      End If
      If !birthdate >= pTestDate2 + 63 And !birthdate <= pTestDate2 + 83 Then
         period(4) = period(4) + 1
         
      End If
      If !birthdate > pTestDate2 + 84 Then
         period(5) = period(5) + 1
         
      End If
      End If
      .MoveNext
   Loop
End With
CDL = CalfDeath / denom
If (RS.RecordCount - A - b) <> 0 Then
  CalfDeath = (CalfDeath / (RS.RecordCount - A - b)) * 100
Else
  CalfDeath = 0
End If

SQL = "select actweight, managecode from siresumtmp"
Set RS = repdb.OpenRecordset(SQL, dbOpenSnapshot)

With RS
   Do Until .EOF
   If !actweight > 0 Then
      ActWtCount = ActWtCount + 1
   End If
   If !managecode = "F" Then
      F = F + 1
   End If
   If !managecode = "X" Then
      x = x + 1
   End If
   .MoveNext
   Loop
End With
DB.Close: Set DB = Nothing
repdb.Execute ("delete * from reproper")
Set RS = repdb.OpenRecordset("reproper", dbOpenTable)
Set RepRS = repdb.OpenRecordset("cowcount", dbOpenSnapshot)

With RS
   .AddNew
   allcalf(0) = period(0) + period(1) + period(2) + period(3) + period(4) + period(5)
   If allcalf(0) <> 0 Then
     !first21 = ((period(1) + period(0)) / allcalf(0)) * 100
     !first42 = ((period(2) + period(0) + period(1)) / (allcalf(0))) * 100
     !first63 = ((period(2) + period(0) + period(1) + period(3)) / (allcalf(0))) * 100
   Else
     !first21 = 0
     !first42 = 0
     !first63 = 0
   End If
   !after63 = 100 - !first63
   !PregPer = PregPer
   !PLP = PLP
   !CP = CP
   !CDL = CDL * 100
   !WeanPr = ((ActWtCount + x - F) / denom) * 100
   !CalfDeath = CalfDeath
   If (denom) <> 0 Then
     !pndperexpfem = allwt / (denom)
   Else
     !pndperexpfem = 0
   End If
   If allcalf(1) <> 0 Then
     !avgweanwt = allwt / allcalf(1)
   Else
     !avgweanwt = 0
   End If
   .Update
End With

RS.Close: Set RS = Nothing
RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
Exit Sub
ehandle:
If Err.Number = 94 Then Resume Next
End Sub

Public Sub FindTurnDate()
Dim SQL$, DB As database, RS As Recordset, pPassed As Boolean
   SQL = "SELECT DISTINCTROW cowprof.enteredherd, calfbirth.CowID, calfwean.managecode, calfbirth.birthdate, calfbirth.CowAge FROM cowprof INNER JOIN (calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID) where calfwean.managecode <> 'P' And calfbirth.birthdate <> #1/1/1900# "
   If IsDate(calfreps.txtStartDate) Then SQL = SQL & "And calfbirth.birthdate >= #" & calfreps.txtStartDate.TEXT & "# "
   If IsDate(calfreps.txtEndDate) Then SQL = SQL & "And calfbirth.birthdate <= #" & calfreps.txtEndDate.TEXT & "# "
    SQL = SQL & " and calfwean.managecode <> 'A' and calfwean.managecode <> 'B' and calfwean.managecode <> 'P' And calfbirth.cowage > 2 "
ORDER BY calfbirth.birthdate"
   ActTurnDate = TurnDate
   TurnDate = #10:00:00 AM#
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
   
   Set RS = DB.OpenRecordset(SQL, dbOpenDynaset)
   
   F = 0
   P = 0
   TCE = 0
   
   If Not RS.EOF Then
      RS.MoveNext
      If Not RS.EOF Then RS.MoveNext
      If RS.EOF Then GoTo Failed_Test
      Do Until pPassed = True
         If RS!enteredherd < ActTurnDate Or IsNull(RS!enteredherd) Then pPassed = True: Exit Do
         RS.MoveNext
      Loop
      thirdcowdate = Field2Date(RS!birthdate)
      ThirdCow = Field2Str(RS!CowID)
         If DateDiff("D", CDate(Left(calfreps.cboyear.TEXT, 10)), thirdcowdate) > 295 Or DateDiff("D", CDate(Left(calfreps.cboyear.TEXT, 10)), thirdcowdate) < 275 Then
            TurnDate = thirdcowdate - 285
            Calculated = False
         Else
            Calculated = True
         End If
         BoolThirdCowDate = True
   Else
Failed_Test:
      Calculated = True
      BoolThirdCowDate = False
    End If

CloseDB:
   RS.Close: Set RS = Nothing
   DB.Close: Set DB = Nothing

   If Calculated = True Then
      mTestDate = ActTurnDate + 285
   Else
      mTestDate = TurnDate + 285
   End If

End Sub

Public Sub UpdateTurnOutDate(NewDate As Date, Herd As String)
Dim tbCowBrd As Recordset
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set tbCowBrd = DB.OpenRecordset("misc", dbOpenTable)
   tbCowBrd.Index = "primarykey"
   tbCowBrd.Seek "=", "TurnDate" & Herd
If tbCowBrd.NoMatch Then
   tbCowBrd.AddNew
   tbCowBrd!thekey = "TurnDate" & Herd
   tbCowBrd!thetext = CStr(NewDate)
   tbCowBrd.Update
Else
   tbCowBrd.Edit
   tbCowBrd!thekey = "TurnDate" & Herd
   tbCowBrd!thetext = CStr(NewDate)
   tbCowBrd.Update
End If
tbCowBrd.Close: Set tbCowBrd = Nothing
DB.Close: Set DB = Nothing
End Sub

Public Sub SortCalfData(SortOrder%)
Dim repdb As database, SQL$, order$
'If SortOrder% = 5 Or SortOrder% = 0 Then Exit Sub
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
Select Case SortOrder
   Case 0
      order$ = ".calfid"
   Case 1
      order$ = ".adj205wt"
   Case 2
      order$ = ".sireid"
   Case 3
      order$ = ".actweight"
   Case 4
      order$ = ".age_in_days"
   Case 5
      order$ = ".age"
   Case 6
      order$ = ".birthwt"
   Case 7
      order$ = ".cframe"
End Select
If Not SortOrder% = 5 And Not SortOrder% = 0 Then
'Heifer - bulls averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
'repdb.Execute ("insert into calfsort SELECT DISTINCTROW HBulAvg.SireID From HBulAvg GROUP BY HBulAvg.SireID, HBulAvg" & Order$ & " ORDER BY HBulAvg" & Order$), dbFailOnError
'repdb.Execute ("update calfsort, hbulavg set calfsort.num = hbulavg.num, calfsort.sire_breed = hbulavg.sire_breed, calfsort.adj205wt = hbulavg.adj205wt, calfsort.birthwt = hbulavg.birthwt, calfsort.calvingease = hbulavg.calvingease, calfsort.actweight = hbulavg.actweight, calfsort.age_in_days = hbulavg.age_in_days, calfsort.cframe = hbulavg.cframe, calfsort.avgdailygain = hbulavg.avgdailygain, calfsort.wt2daygain = hbulavg.wt2daygain where calfsort.sireid = hbulavg.sireid"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW hbulavg.sireid From hbulavg GROUP BY hbulavg.sireid, hbulavg" & order$ & " ORDER BY hbulavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, hbulavg set calfsort.num = hbulavg.num, calfsort.sireid = hbulavg.sireid, calfsort.adj205wt = hbulavg.adj205wt, calfsort.birthwt = hbulavg.birthwt, calfsort.calvingease = hbulavg.calvingease, calfsort.actweight = hbulavg.actweight, calfsort.age_in_days = hbulavg.age_in_days, calfsort.cframe = hbulavg.cframe, calfsort.avgdailygain = hbulavg.avgdailygain, calfsort.wt2daygain = hbulavg.wt2daygain where calfsort.sireid = hbulavg.sireid"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW hbulavg.sireid, hbulavg.group From hbulavg GROUP BY hbulavg.sireid, hbulavg.group, hbulavg" & order$ & " ORDER BY hbulavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW hbulavg SET hbulavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, hbulavg set calfsort.sire_breed = hbulavg.sire_breed, calfsort.num = hbulavg.num, calfsort.sireid = hbulavg.sireid, calfsort.adj205wt = hbulavg.adj205wt, calfsort.birthwt = hbulavg.birthwt, calfsort.calvingease = hbulavg.calvingease, calfsort.actweight = hbulavg.actweight, calfsort.age_in_days = hbulavg.age_in_days, calfsort.cframe = hbulavg.cframe, calfsort.avgdailygain = hbulavg.avgdailygain, calfsort.wt2daygain = hbulavg.wt2daygain where calfsort.sireid = hbulavg.sireid and calfsort.group = hbulavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from hbulavg")
repdb.Execute ("insert into hbulavg select * from calfsort"), dbFailOnError
'Heifer - cowbreed averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW heiavg.cow_breed From heiavg GROUP BY heiavg.cow_breed, heiavg" & order$ & " ORDER BY heiavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, heiavg set calfsort.num = heiavg.num, calfsort.cow_breed = heiavg.cow_breed, calfsort.adj205wt = heiavg.adj205wt, calfsort.birthwt = heiavg.birthwt, calfsort.calvingease = heiavg.calvingease, calfsort.actweight = heiavg.actweight, calfsort.age_in_days = heiavg.age_in_days, calfsort.cframe = heiavg.cframe, calfsort.avgdailygain = heiavg.avgdailygain, calfsort.wt2daygain = heiavg.wt2daygain where calfsort.cow_breed = heiavg.cow_breed"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW heiavg.cow_breed, heiavg.group From heiavg GROUP BY heiavg.cow_breed, heiavg.group, heiavg" & order$ & " ORDER BY heiavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW heiavg SET heiavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, heiavg set calfsort.num = heiavg.num, calfsort.cow_breed = heiavg.cow_breed, calfsort.adj205wt = heiavg.adj205wt, calfsort.birthwt = heiavg.birthwt, calfsort.calvingease = heiavg.calvingease, calfsort.actweight = heiavg.actweight, calfsort.age_in_days = heiavg.age_in_days, calfsort.cframe = heiavg.cframe, calfsort.avgdailygain = heiavg.avgdailygain, calfsort.wt2daygain = heiavg.wt2daygain where calfsort.cow_breed = heiavg.cow_breed and calfsort.group = heiavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from heiavg")
repdb.Execute ("insert into heiavg select * from calfsort"), dbFailOnError
'Bulls - bull averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
'repdb.Execute ("insert into calfsort SELECT DISTINCTROW bulavg.SireID From bulavg GROUP BY bulavg.SireID, bulavg" & Order$ & " ORDER BY bulavg" & Order$), dbFailOnError
'repdb.Execute ("update calfsort, bulavg set calfsort.num = bulavg.num, calfsort.sire_breed = bulavg.sire_breed, calfsort.adj205wt = bulavg.adj205wt, calfsort.birthwt = bulavg.birthwt, calfsort.calvingease = bulavg.calvingease, calfsort.actweight = bulavg.actweight, calfsort.age_in_days = bulavg.age_in_days, calfsort.cframe = bulavg.cframe, calfsort.avgdailygain = bulavg.avgdailygain, calfsort.wt2daygain = bulavg.wt2daygain where calfsort.sireid = bulavg.sireid"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW bulavg.sireid From bulavg GROUP BY bulavg.sireid, bulavg" & order$ & " ORDER BY bulavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, bulavg set calfsort.num = bulavg.num, calfsort.sireid = bulavg.sireid, calfsort.adj205wt = bulavg.adj205wt, calfsort.birthwt = bulavg.birthwt, calfsort.calvingease = bulavg.calvingease, calfsort.actweight = bulavg.actweight, calfsort.age_in_days = bulavg.age_in_days, calfsort.cframe = bulavg.cframe, calfsort.avgdailygain = bulavg.avgdailygain, calfsort.wt2daygain = bulavg.wt2daygain where calfsort.sireid = bulavg.sireid"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW bulavg.sireid, bulavg.group From bulavg GROUP BY bulavg.sireid, bulavg.group, bulavg" & order$ & " ORDER BY bulavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW bulavg SET bulavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, bulavg set calfsort.sire_breed = bulavg.sire_breed, calfsort.num = bulavg.num, calfsort.sireid = bulavg.sireid, calfsort.adj205wt = bulavg.adj205wt, calfsort.birthwt = bulavg.birthwt, calfsort.calvingease = bulavg.calvingease, calfsort.actweight = bulavg.actweight, calfsort.age_in_days = bulavg.age_in_days, calfsort.cframe = bulavg.cframe, calfsort.avgdailygain = bulavg.avgdailygain, calfsort.wt2daygain = bulavg.wt2daygain where calfsort.sireid = bulavg.sireid and calfsort.group = bulavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from bulavg")
repdb.Execute ("insert into bulavg select * from calfsort"), dbFailOnError
'Bull - cowbreed averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
'repdb.Execute ("insert into calfsort SELECT DISTINCTROW bheiavg.cow_breed From bheiavg GROUP BY bheiavg.cow_breed, bheiavg" & Order$ & " ORDER BY bheiavg" & Order$), dbFailOnError
'repdb.Execute ("update calfsort, bheiavg set calfsort.num = bheiavg.num, calfsort.cow_breed = bheiavg.cow_breed, calfsort.adj205wt = bheiavg.adj205wt, calfsort.birthwt = bheiavg.birthwt, calfsort.calvingease = bheiavg.calvingease, calfsort.actweight = bheiavg.actweight, calfsort.age_in_days = bheiavg.age_in_days, calfsort.cframe = bheiavg.cframe, calfsort.avgdailygain = bheiavg.avgdailygain, calfsort.wt2daygain = bheiavg.wt2daygain where calfsort.cow_breed = bheiavg.cow_breed"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW bheiavg.cow_breed From bheiavg GROUP BY bheiavg.cow_breed, bheiavg" & order$ & " ORDER BY bheiavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, bheiavg set calfsort.num = bheiavg.num, calfsort.cow_breed = bheiavg.cow_breed, calfsort.adj205wt = bheiavg.adj205wt, calfsort.birthwt = bheiavg.birthwt, calfsort.calvingease = bheiavg.calvingease, calfsort.actweight = bheiavg.actweight, calfsort.age_in_days = bheiavg.age_in_days, calfsort.cframe = bheiavg.cframe, calfsort.avgdailygain = bheiavg.avgdailygain, calfsort.wt2daygain = bheiavg.wt2daygain where calfsort.cow_breed = bheiavg.cow_breed"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW bheiavg.cow_breed, bheiavg.group From bheiavg GROUP BY bheiavg.cow_breed, bheiavg.group, bheiavg" & order$ & " ORDER BY bheiavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW bheiavg SET bheiavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, bheiavg set calfsort.num = bheiavg.num, calfsort.cow_breed = bheiavg.cow_breed, calfsort.adj205wt = bheiavg.adj205wt, calfsort.birthwt = bheiavg.birthwt, calfsort.calvingease = bheiavg.calvingease, calfsort.actweight = bheiavg.actweight, calfsort.age_in_days = bheiavg.age_in_days, calfsort.cframe = bheiavg.cframe, calfsort.avgdailygain = bheiavg.avgdailygain, calfsort.wt2daygain = bheiavg.wt2daygain where calfsort.cow_breed = bheiavg.cow_breed and calfsort.group = bheiavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from bheiavg")
repdb.Execute ("insert into bheiavg select * from calfsort"), dbFailOnError
'Steers - bull averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
'repdb.Execute ("insert into calfsort SELECT DISTINCTROW sbulavg.SireID From sbulavg GROUP BY sbulavg.SireID, sbulavg" & Order$ & " ORDER BY sbulavg" & Order$), dbFailOnError
'repdb.Execute ("update calfsort, sbulavg set calfsort.num = sbulavg.num, calfsort.sire_breed = sbulavg.sire_breed, calfsort.adj205wt = sbulavg.adj205wt, calfsort.birthwt = sbulavg.birthwt, calfsort.calvingease = sbulavg.calvingease, calfsort.actweight = sbulavg.actweight, calfsort.age_in_days = sbulavg.age_in_days, calfsort.cframe = sbulavg.cframe, calfsort.avgdailygain = sbulavg.avgdailygain, calfsort.wt2daygain = sbulavg.wt2daygain where calfsort.sireid = sbulavg.sireid"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW sbulavg.sireid From sbulavg GROUP BY sbulavg.sireid, sbulavg" & order$ & " ORDER BY sbulavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, sbulavg set calfsort.num = sbulavg.num, calfsort.sireid = sbulavg.sireid, calfsort.adj205wt = sbulavg.adj205wt, calfsort.birthwt = sbulavg.birthwt, calfsort.calvingease = sbulavg.calvingease, calfsort.actweight = sbulavg.actweight, calfsort.age_in_days = sbulavg.age_in_days, calfsort.cframe = sbulavg.cframe, calfsort.avgdailygain = sbulavg.avgdailygain, calfsort.wt2daygain = sbulavg.wt2daygain where calfsort.sireid = sbulavg.sireid"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW sbulavg.sireid, sbulavg.group From sbulavg GROUP BY sbulavg.sireid, sbulavg.group, sbulavg" & order$ & " ORDER BY sbulavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW sbulavg SET sbulavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, sbulavg set calfsort.sire_breed = sbulavg.sire_breed, calfsort.num = sbulavg.num, calfsort.sireid = sbulavg.sireid, calfsort.adj205wt = sbulavg.adj205wt, calfsort.birthwt = sbulavg.birthwt, calfsort.calvingease = sbulavg.calvingease, calfsort.actweight = sbulavg.actweight, calfsort.age_in_days = sbulavg.age_in_days, calfsort.cframe = sbulavg.cframe, calfsort.avgdailygain = sbulavg.avgdailygain, calfsort.wt2daygain = sbulavg.wt2daygain where calfsort.sireid = sbulavg.sireid and calfsort.group = sbulavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from sbulavg")
repdb.Execute ("insert into sbulavg select * from calfsort"), dbFailOnError
'Steers - cowbreed averages
repdb.Execute ("delete * from calfsort"), dbFailOnError
'repdb.Execute ("insert into calfsort SELECT DISTINCTROW sheiavg.cow_breed From sheiavg GROUP BY sheiavg.cow_breed, sheiavg" & Order$ & " ORDER BY sheiavg" & Order$), dbFailOnError
'repdb.Execute ("update calfsort, sheiavg set calfsort.num = sheiavg.num, calfsort.cow_breed = sheiavg.cow_breed, calfsort.adj205wt = sheiavg.adj205wt, calfsort.birthwt = sheiavg.birthwt, calfsort.calvingease = sheiavg.calvingease, calfsort.actweight = sheiavg.actweight, calfsort.age_in_days = sheiavg.age_in_days, calfsort.cframe = sheiavg.cframe, calfsort.avgdailygain = sheiavg.avgdailygain, calfsort.wt2daygain = sheiavg.wt2daygain where calfsort.cow_breed = sheiavg.cow_breed"), dbFailOnError
If group = False Then
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW sheiavg.cow_breed From sheiavg GROUP BY sheiavg.cow_breed, sheiavg" & order$ & " ORDER BY sheiavg" & order$), dbFailOnError
   repdb.Execute ("update calfsort, sheiavg set calfsort.num = sheiavg.num, calfsort.cow_breed = sheiavg.cow_breed, calfsort.adj205wt = sheiavg.adj205wt, calfsort.birthwt = sheiavg.birthwt, calfsort.calvingease = sheiavg.calvingease, calfsort.actweight = sheiavg.actweight, calfsort.age_in_days = sheiavg.age_in_days, calfsort.cframe = sheiavg.cframe, calfsort.avgdailygain = sheiavg.avgdailygain, calfsort.wt2daygain = sheiavg.wt2daygain where calfsort.cow_breed = sheiavg.cow_breed"), dbFailOnError
Else
   repdb.Execute ("insert into calfsort SELECT DISTINCTROW sheiavg.cow_breed, sheiavg.group From sheiavg GROUP BY sheiavg.cow_breed, sheiavg.group, sheiavg" & order$ & " ORDER BY sheiavg" & order$), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW CalfSort SET CalfSort.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("UPDATE DISTINCTROW sheiavg SET sheiavg.[group] = '0' where isnull(group)"), dbFailOnError
   repdb.Execute ("update calfsort, sheiavg set calfsort.num = sheiavg.num, calfsort.cow_breed = sheiavg.cow_breed, calfsort.adj205wt = sheiavg.adj205wt, calfsort.birthwt = sheiavg.birthwt, calfsort.calvingease = sheiavg.calvingease, calfsort.actweight = sheiavg.actweight, calfsort.age_in_days = sheiavg.age_in_days, calfsort.cframe = sheiavg.cframe, calfsort.avgdailygain = sheiavg.avgdailygain, calfsort.wt2daygain = sheiavg.wt2daygain where calfsort.cow_breed = sheiavg.cow_breed and calfsort.group = sheiavg.group"), dbFailOnError
End If
repdb.Execute ("delete * from sheiavg")
repdb.Execute ("insert into sheiavg select * from calfsort"), dbFailOnError
'Sire summary order
repdb.Execute ("delete * from siresumtmp")
repdb.Execute ("insert into siresumtmp select * from siresum order by siresum" & order$), dbFailOnError
repdb.Execute ("delete * from siresum")
repdb.Execute ("insert into siresum select * from siresumtmp")
End If
'Calf id list tables
'heifers
repdb.Execute ("delete * from calfsort")
repdb.Execute ("insert into calfsort select * from heical order by heical" & order$), dbFailOnError
repdb.Execute ("delete * from heical")
repdb.Execute ("insert into heical select calfid, birthdate,birthwt,calvingease,actweight, dateweighed, managecode, cframe, group, misc1, age, cowid, age_in_days, adj205wt, adj205rat, avgdailygain, wt2daygain, sex, cow_breed, sire_breed, irrcalf, skipme, herdid, sireid from calfsort"), dbFailOnError
'bulls
repdb.Execute ("delete * from calfsort")
repdb.Execute ("insert into calfsort select * from bulcalf order by bulcalf" & order$), dbFailOnError
repdb.Execute ("delete * from bulcalf")
repdb.Execute ("insert into bulcalf select calfid, birthdate,birthwt,calvingease,actweight, dateweighed, managecode, cframe, group, misc1, age, cowid, age_in_days, adj205wt, adj205rat, avgdailygain, wt2daygain, sex, cow_breed, sire_breed, irrcalf, skipme, herdid, sireid from calfsort"), dbFailOnError
'steers
repdb.Execute ("delete * from calfsort")
repdb.Execute ("insert into calfsort select * from steercalf order by steercalf" & order$), dbFailOnError
repdb.Execute ("delete * from steercalf")
repdb.Execute ("insert into steercalf select calfid, birthdate,birthwt,calvingease,actweight, dateweighed, managecode, cframe, group, misc1, age, cowid, age_in_days, adj205wt, adj205rat, avgdailygain, wt2daygain, sex, cow_breed, sire_breed, irrcalf, skipme, herdid, sireid from calfsort"), dbFailOnError
'misc
repdb.Execute ("delete * from calfsort")
repdb.Execute ("insert into calfsort select * from misccalf order by misccalf" & order$), dbFailOnError
repdb.Execute ("delete * from misccalf")
repdb.Execute ("insert into misccalf select calfid, birthdate,birthwt,calvingease,actweight, dateweighed, managecode, cframe, group, misc1, age, cowid, age_in_days, adj205wt, adj205rat, avgdailygain, wt2daygain, sex, cow_breed, sire_breed, irrcalf, skipme, herdid, sireid from calfsort"), dbFailOnError
repdb.Close: Set repdb = Nothing
End Sub


