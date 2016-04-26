Attribute VB_Name = "YearlingReps"
Option Explicit

Public Sub Create365Reports(orderby$, StartDate As String, enddate As String, EndDayDate As String, StartDayDate As String)
'On Error Resume Next
Dim Sex(3) As Single
Dim SQL$, X As Integer, dam As Double, adj205 As Double
Dim DB As database, RS As Recordset, RepRS As Recordset, repdb As database
Dim reprsavg As Recordset
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set repdb = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
repdb.Execute ("delete * from yearlings")
For X = 0 To 3
 SQL = " INSERT INTO yearlings ( birthdate, DaysTest, AgeOff, datetest, wgton, wgtoff, ADGOn, WDAOff, frscore, misc, sireID, sex, yearid, BirthWt, Age, ScrotumCir, Pelvic, Managecode, ActWeight ) IN '" & repfile$ & "' "
 SQL$ = SQL$ & " SELECT DISTINCTROW calfbirth.birthdate, Sum([calfrep].[daydate]-[calfwean].[dateweighed]) AS Expr1, Sum([calfrep].[daydate]-[calfbirth].[birthdate]) AS Expr2, calfwean.dateweighed, calfwean.actweight, calfrep.daywt, IIf(calfrep.daywt>0,Sum((([calfrep].[daywt]-[calfwean].[actweight])/([calfrep].[daydate]-[calfwean].[dateweighed]))),0) AS Expr3, Sum(([calfrep]![daywt]/([calfrep]![daydate]-[calfbirth]![birthdate]))) AS Expr4, calfrep.dayscore AS frscore, calfwean.misc1, calfbirth.sireID, calfbirth.sex, calfrep.CalfID, calfbirth.birthwt, calfbirth.CowAge, calfrep.scrotumcir, calfrep.pelvic, calfwean.managecode, calfwean.actweight "
 SQL$ = SQL$ & " FROM (calfbirth INNER JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID) "
 SQL$ = SQL$ & " WHERE calfbirth.sex='" & X & "' and calfbirth.herdid = '" & herdid & "' "
If StartDate <> "--/--/----" Then
   SQL$ = SQL$ & " and calfbirth.birthdate >= #" & StartDate & "# "
End If
If enddate <> "--/--/----" Then
   SQL$ = SQL$ & " and calfbirth.birthdate <= #" & enddate & "#"
End If
If StartDayDate <> "--/--/----" Then
   SQL$ = SQL$ & " and calfrep.daydate >= #" & StartDayDate & "#"
End If
If EndDayDate <> "--/--/----" Then
   SQL$ = SQL$ & " and calfrep.daydate <= #" & EndDayDate & "#"
End If
SQL$ = SQL$ & "  GROUP BY calfbirth.birthdate, calfwean.dateweighed, calfwean.actweight, calfrep.daywt, calfrep.dayscore, calfwean.misc1, calfbirth.sireID, calfbirth.sex, calfrep.CalfID, calfbirth.birthwt, calfbirth.CowAge, calfrep.scrotumcir, calfrep.pelvic, calfwean.managecode ORDER BY calfrep.CalfID"
DB.Execute (SQL$), dbFailOnError
repdb.Execute "delete * from yearlings where managecode = 'A'"
repdb.Execute "delete * from yearlings where managecode = 'B'"
repdb.Execute "delete * from yearlings where managecode = 'C'"
repdb.Execute "delete * from yearlings where managecode = 'D'"
Set RepRS = repdb.OpenRecordset("Select Count(YearID) as Num from yearlings", dbOpenSnapshot)
If Not RepRS.EOF Then
   If Field2Num(RepRS!num) = 0 Then GoTo NextRow
End If
RepRS.Close: Set RepRS = Nothing
' Find WDA Avg for each sex from yearling table
   SQL$ = "SELECT DISTINCTROW Avg(Yearlings.WDAOff) AS Avg, Yearlings.sex From Yearlings GROUP BY Yearlings.sex HAVING (((Yearlings.sex)='" & CStr(X) & "'))"
   Set RepRS = repdb.OpenRecordset(SQL$, dbOpenSnapshot)
   If Not RepRS.EOF Then Sex(X) = IIf(IsNull(RepRS!Avg), 1, RepRS!Avg)
   SQL$ = "Update Yearlings Set WDARatio = ((WDAOff/" & Sex(X) & ")*100) Where Sex = '" & CStr(X) & "' and wdaoff > 0"
   repdb.Execute (SQL$)
NextRow:
Next X
   Set RepRS = repdb.OpenRecordset("yearlings", dbOpenTable)
   With RepRS
      Do Until .EOF
         .Edit
         If !datetest - !birthdate < 250 And !datetest - !birthdate > 160 Then
                If !birthwt = "" Or IsNull(!birthwt) Then
                     If !Sex = 1 Or !Sex = 3 Then
                           !birthwt = 75
                     End If
                     If !Sex = 2 Then
                           !birthwt = 70
                     End If
                  End If
               adj205 = (!wgton - !birthwt) / (!datetest - !birthdate) * 205 + !birthwt + FindDam(CInt(!Sex), CInt(!age))
               !Wgt365 = (((!wgtoff - !wgton) / !DaysTest) * 160) + adj205
         Else
            !Wgt365 = ((!WDAoff * 365) + FindDam(CInt(!Sex), CInt(!age)))
         End If
         .Update
         .MoveNext
      Loop
   End With
For X = 0 To 3
SQL$ = "SELECT DISTINCTROW Yearlings.DaysTest, Yearlings.YearID, Yearlings.AgeOff, Avg(Yearlings.Wgt365) AS Avg, Yearlings.sex From Yearlings"
SQL$ = SQL$ & " GROUP BY Yearlings.DaysTest, Yearlings.YearID, Yearlings.AgeOff, Yearlings.sex, yearlings.herdid"
SQL$ = SQL$ & " HAVING (((Yearlings.DaysTest)>111) AND ((Yearlings.AgeOff)>320 And (Yearlings.AgeOff)<410) AND ((Yearlings.sex)='" & CStr(X) & "'));"
Set RepRS = repdb.OpenRecordset(SQL$, dbOpenSnapshot)
  If RepRS.RecordCount < 1 Then GoTo NextRowA
   'Sex(X) = IIf(IsNull(RepRS!Avg), 1, RepRS!Avg)
   SQL$ = " SELECT  Avg(Yearlings.Wgt365) AS AvgOfWgt365"
   SQL$ = SQL$ & " From Yearlings GROUP BY  Yearlings.Sex HAVING (((Yearlings.Sex)='" & CStr(X) & "') AND ((Avg(Yearlings.Wgt365))>0));"
   Set reprsavg = repdb.OpenRecordset(SQL$, dbOpenSnapshot)
   If Not reprsavg.EOF Then
     SQL$ = "Update Yearlings Set 365Ratio = wgt365/" & Field2Num(reprsavg!AvgOfWgt365) & " * 100 Where Sex = '" & CStr(X) & "' and wgt365 > 0"
     'SQL$ = "Update Yearlings Set WDARatio = ((WDAOff/" & Sex(X) & ")*100) Where Sex = '" & CStr(X) & "' and wdaoff > 0"
   End If
   repdb.Execute (SQL$)
NextRowA:
Next X
'repdb.Execute ("delete * from yrgrpavg")
'For X = 0 To 3
'SQL$ = "INSERT INTO YrGrpAvg SELECT DISTINCTROW Avg(Yearlings.ADGOn) AS ADGOn, Avg(Yearlings.WDAOff) AS WDAOff, Avg(Yearlings.Wgt365) AS Wgt365, Avg(Yearlings.FrScore) AS FrScore, herdid, sex"
'SQL$ = SQL$ & " From Yearlings GROUP BY Yearlings.sex, yearlings.herdid HAVING (((Yearlings.sex)='" & CStr(X) & "'))"
'repdb.Execute (SQL$)
'Next X
'repdb.Execute ("delete * from yrsireavgs")
'For X = 0 To 3
'SQL$ = "Insert into yrsireavgs"
'SQL$ = SQL$ & " SELECT DISTINCTROW Yearlings.sex, Yearlings.SireID, Count(Yearlings.YearID) AS NoYrlngs, Avg(Yearlings.ADGOn) AS ADGOn, Avg(Yearlings.WDAOff) AS WDAOff, Avg(Yearlings.Wgt365) AS Wgt365, Avg(Yearlings.FrScore) AS FrScore From Yearlings"
'SQL$ = SQL$ & " GROUP BY Yearlings.sex, Yearlings.SireID HAVING (((Yearlings.sex)='" & CStr(X) & "'));"
'repdb.Execute (SQL$)
'Next X
'repdb.Execute ("delete * from yearlingssort")
'For X = 0 To 3
'  repdb.Execute ("insert into yearlingssort select * from yearlings where sex = '" & CStr(X) & "' order by " & orderby$)
'Next X
'repdb.Execute ("delete * from yearlings")
'repdb.Execute ("INSERT INTO yearlings ( Sex, YearID, HerdID, BirthDate, DateTest, DaysTest, AgeOff, WgtOn, WgtOff, ADGOn, WDAOff, WDARatio, Wgt365, 365Ratio, FrScore, Misc, SireID, SireBrd, BirthWt, Age ) SELECT yearlingssort.Sex, yearlingssort.YearID, yearlingssort.HerdID, yearlingssort.BirthDate, yearlingssort.DateTest, yearlingssort.DaysTest, yearlingssort.AgeOff, yearlingssort.WgtOn, yearlingssort.WgtOff, yearlingssort.ADGOn, yearlingssort.WDAOff, yearlingssort.WDARatio, yearlingssort.Wgt365, yearlingssort.[365Ratio], yearlingssort.FrScore, yearlingssort.Misc, yearlingssort.SireID, yearlingssort.SireBrd, yearlingssort.BirthWt, yearlingssort.Age FROM YearInclude INNER JOIN yearlingssort ON (yearlingssort.HerdID = YearInclude.IncludeHerdID) AND (YearInclude.IncludeCalfID = yearlingssort.YearID)"), dbFailOnError
'RepRS.Close: Set RepRS = Nothing
repdb.Close: Set repdb = Nothing
DB.Close: Set DB = Nothing
End Sub

Public Function FindDam(Sex As Integer, age As Integer) As Double
      If Sex = 2 Then
         If age = 2 Then
            FindDam = 54
         End If
         If age = 3 Then
            FindDam = 36
         End If
      If age = 4 Then
         FindDam = 18
      End If
      If age > 11 Then
         FindDam = 18
      End If
   End If
   If Sex = 1 Or Sex = 3 Then
      If age = 2 Then
         FindDam = 60
      End If
      If age = 3 Then
         FindDam = 40
      End If
      If age = 4 Then
         FindDam = 20
      End If
      If age > 11 Then
         FindDam = 20
      End If
   End If
End Function
