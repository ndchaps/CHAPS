Attribute VB_Name = "DaoRoutines"
Option Explicit
 Dim dbpm As database

Public Function DatabaseTableExists(DB As database, TableName As String) As Boolean
'*************************************************
' Programmer: Lynn Owens , David Boys
' Date Created: 01-23-1998
'
' Modified Dates:
'
' Description: This routine will check to see if a table exists in a database
'
'
' Parameters In:
'           db: Name of the Database
'    tableName: Name of the table to search for
'
' Parameters Out
'        DatabaseTableExists: True if table exists.  False otherwise
'
' Example:
'   ret = DatabaseTableExists(dbagvacct, "Invoice")
'
'******************************************************************************
 
 Dim RS As Recordset
 On Local Error GoTo LeHandle
 Set RS = DB.OpenRecordset(TableName, dbOpenTable)
 RS.Close: Set RS = Nothing
 DatabaseTableExists = True
Exit Function
 
LeHandle:
' MsgBox Err & " " & Error$
' Resume
 If Err = 3011 Then DatabaseTableExists = False: Exit Function
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "DoaRoutines - DatabaseTableExists"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())
 
End Function

Public Function data_exists2(DbName$, tbname$, hmkeys%, key$()) As Boolean
'function for use in validity seeks
'  send in tbname$ - name of table to seek in
'          hmkeys% - how many keys used on the seek (1-5)
'          key$() - array of up to 5 keys (0-4)
'
'data_exists = true - seek matched
'data_exists = false - seek did not match
'
 Dim dbtarget As database
 Dim rec As Recordset
 
 data_exists2 = True
 Set dbtarget = DBEngine(0).OpenDatabase(DbName$, 0, 0)
 Set rec = dbtarget.OpenRecordset(tbname$, dbOpenTable)
 rec.Index = "primarykey"
 DBEngine.Idle (dbFreeLocks)
 Select Case hmkeys%
   Case "1"
     rec.Seek "=", key$(0)
   Case "2"
     rec.Seek "=", key$(0), key$(1)
   Case "3"
     rec.Seek "=", key$(0), key$(1), key$(2)
   Case "4"
     rec.Seek "=", key$(0), key$(1), key$(2), key$(3)
   Case "5"
     rec.Seek "=", key$(0), key$(1), key$(2), key$(3), key$(4)
 End Select
 If rec.NoMatch Then data_exists2 = False
 rec.Close: Set rec = Nothing
 dbtarget.Close: Set dbtarget = Nothing
End Function

Public Sub CreateQueryDef(database As String, Defname As String, sqltext)
'A public routine to create a database QueryDef in a specified Database obj.
'
'  Inputs:  Database - string name of database including path etc... ie: acctdb$ or dbname$
'           Defname  - string name of QueryDef to be created
'           sqltext  - variant string of QueryDef SQL statement
'
' Programmer: Mark Barrow
'    Created: 6/3/97

 Dim dbs As database
 Dim qdfNew As QueryDef
 Dim SaveMessage$
 
 On Local Error GoTo LeHandle
 Call DeleteQueryDef(database, Defname)
 Screen.MousePointer = vbHourglass
    Set dbs = OpenDatabase(database, exclusiveyn%, readonlyyn%)
    With dbs
        ' Create QueryDef.
        Set qdfNew = .CreateQueryDef(Defname, sqltext)
        .Close
    End With
    Set dbs = Nothing
    Screen.MousePointer = vbNormal
Exit Sub

LeHandle:
' if table is locked or permission denied wait and do again
 If Err.Number = 3260 Or Err.Number = 3187 Or Err.Number = 3186 Or Err.Number = 3356 Or Err.Number = 3189 Then
   'SaveMessage$ = DisplayStatus(4, "Waiting For Database Access....")
   Call WaitSecs(2)
   'SaveMessage$ = DisplayStatus(4, SaveMessage$)
   Resume
 End If
 TEXT$(1) = "Database: " & database
 TEXT$(2) = "DefName: " & Defname
 TEXT$(3) = "SQL: " & sqltext
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "DaoRoutines - CreateQueryDef"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())
End Sub

Public Sub DeleteQueryDef(database As String, Defname As String)
'A public routine to delete a database QueryDef in a specified Database obj.
'
'  Inputs:  Database - string name of database including path etc... ie: acctdb$ or dbname$
'           Defname  - string name of QueryDef to be deleted
'
' Programmer: Mark Barrow
'    Created: 6/3/97

 Dim dbs As database
 Dim qdfNew As QueryDef
 Dim SaveMessage$
 
 On Local Error GoTo LeHandle
1 Screen.MousePointer = vbHourglass
2    Set dbs = OpenDatabase(database, exclusiveyn%, readonlyyn%)
3    With dbs
5        .QueryDefs.Delete Defname
6        .Close
7        End With
8    Set dbs = Nothing
9    Screen.MousePointer = vbNormal
Exit Sub

LeHandle:
' if table is locked or permission denied wait and do again
 If Erl = 5 And Err.Number = 3265 Then Resume 6
 If Err.Number = 3260 Or Err.Number = 3187 Or Err.Number = 3186 Or Err.Number = 3356 Or Err.Number = 3189 Then
   'SaveMessage$ = DisplayStatus(4, "Waiting For Database Access....")
   Call WaitSecs(2)
   'SaveMessage$ = DisplayStatus(4, SaveMessage$)
   Resume
 End If
 TEXT$(1) = "Database: " & database
 TEXT$(2) = "DefName: " & Defname
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "DaoRoutines - DeleteQueryDef"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub


Function Field2Num(thefield As Field) As Variant
 'this sub converts a null number field in the database to a 0
 '
 'inputs
 '   thefield = the field to be checked from the database
  
 If Not IsNull(thefield) Then
   Field2Num = Val(thefield)
  Else
   Field2Num = 0
 End If
End Function

Function Field2Str(thefield As Field) As String
 'this sub converts a null text field in the database to ""
 '
 'inputs
 '   thefield = the field to be checked from the database
  
 If Not IsNull(thefield) Then
   Field2Str = thefield
  Else
   Field2Str = ""
  End If
End Function

Public Sub CheckBox2Boolean(thefield As Field, TheCheckBox As Control)
 'this sub converts a checkbox to the database values of true of false
 '
 'inputs
 '   thefield = the field in the database where the value is to be saved
 '   TheCheckBox = the name of the check box the data is to be saved from
 '
 '
 If TheCheckBox.Value = vbChecked Then
   thefield = True
  Else
   thefield = False
 End If
End Sub

Public Function Field2Boolean(thefield As Field) As Boolean
 'this sub converts a null boolean field in the database to false
 '
 'inputs
 '   thefield = the field to be checked from the database
 
  
 If Not IsNull(thefield) Then
   If thefield Then
     Field2Boolean = True
    Else
     Field2Boolean = False
   End If
  Else
   Field2Boolean = False
  End If
End Function

Public Function Field2Date(thefield As Field) As String
 'this sub converts a null field in the database to a blank date in the form --/--/----
 '
 'inputs
 '   thefield = the field to be checked from the database
 If Not IsNull(thefield) Then
   Field2Date = Format(CDate(thefield), "mm/dd/yyyy")
   If Len(Field2Date) < 10 Then
     Field2Date = "--/--/----"
   End If
  Else
   Field2Date = Format$("--/--/----", "mm/dd/yyyy")
  End If
End Function
Public Function Field2time(thefield As Field) As String
'this sub converts a null time field in the database to a blank in the form --:-- --
 '
 'inputs
 '   thefield = the field to be checked from the database
  
 If Not IsNull(thefield) Then
   Field2time = Format$(Field2Str(thefield), "hh:mm ampm")
  Else
   Field2time = Format$("--:-- --", "hh:mm ampm")
  End If
End Function

Public Sub Date2Field(thefield As Field, TheDate As String)
 'this sub make sure a date is a valid date before it is saved in the database
 '
 'inputs
 '   thefield = the name of the field in the database where the date is to be saved
 '   thedate = is the date to check for validity
 '
 
 If IsDate(TheDate) Then
   thefield = CDate(TheDate)
 Else
   thefield = Null
 End If
End Sub

Public Sub CheckID(DbName As String, TableName As String, ID As String, TheTableName$())
 Dim DB As database
 Dim RECSET As Recordset
 Dim count As Integer
 Dim index1 As Integer
 Dim index2 As Integer
 count = 0
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(DbName, exclusiveyn%, readonlyyn%)
 For index2 = 0 To DB.Relations.count - 1
  If UCase$(DB.Relations(index2).Table) = UCase$(TableName$) Then
    For index1 = 0 To DB.Relations(index2).Fields.count - 1
     Set RECSET = DB.OpenRecordset("SELECT * FROM " & DB.Relations(index2).ForeignTable & " WHERE " & DB.Relations(index2).Fields(index1).ForeignName & "  = '" & ID$ & "'", dbOpenDynaset)
     DBEngine.Idle (dbFreeLocks)
     If Not RECSET.EOF And Not RECSET.BOF Then
       count = count + 1
       TheTableName$(count) = DB.Relations(index2).ForeignTable
     End If
    Next index1
  End If
 Next index2
 DB.Close: Set DB = Nothing
 TheTableName$(0) = Str$(count)
 Screen.MousePointer = vbDefault
End Sub

Public Sub ClearTableAttachemnts(DbName As String, UserID As String)
'*************************************************
' Programmer: Lynn Owens
' Date Created: 04-30-1997
'
' Modified Dates: 05-15-1997 Lynn Owens
'
' Description: This routine will DELETE All table attachment for a database
'
' Parameters In:
'  DbName: The Database where the attachment are
'  UserID:  the id of the user to clear attachments for
'   send in "" for all users
'
' Parameters Out: (none)
'
' Example:
'  call ClearTableAttachments("c:\avdata\agvance.mdb")
'   this would clear all table attcahments in c:\avdata\agvance.mdb  if their are any
'
'******************************************************************************
 Dim DB As database
 Dim TD As TableDef
 Dim TableCount As Integer
 Dim DeleteAttach As Boolean
 Dim SaveMessage$
 On Local Error GoTo LeHandle
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(DbName, exclusiveyn%, readonlyyn%)
' For Each TD In db.TableDefs
'   If TD.Connect <> "" Then
'    TD.Connect = ";Database=" & dbname
'    db.TableDefs.Delete TD.name
'  End If
' Next
 For TableCount = DB.TableDefs.count - 1 To 0 Step -1
  Set TD = DB.TableDefs(TableCount)
  If TD.Connect <> "" Then
    If Len(UserID) Then
      If Right$(TD.Name, Len(UserID)) = UserID Then
        DeleteAttach = True
       Else
        DeleteAttach = False
      End If
     Else
      DeleteAttach = True
    End If
    If DeleteAttach Then
      TD.Connect = ";Database=" & DbName
      DB.TableDefs.Delete TD.Name
    End If
 End If
 Next TableCount
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
 Exit Sub
 
LeHandle:
 Screen.MousePointer = vbDefault
 ' if table is locked or permission denied wait and do again
 If Err.Number = 3260 Or Err.Number = 3187 Or Err.Number = 3186 Or Err.Number = 3356 Or Err.Number = 3189 Then
   'SaveMessage$ = DisplayStatus(4, "Waiting For Database Access....")
   Call WaitSecs(2)
   'SaveMessage$ = DisplayStatus(4, SaveMessage$)
   Resume
 End If
 If Err.Number = 3343 Then
  Load frmrepairdatabase
  frmrepairdatabase.Tag = "Repair"
  frmrepairdatabase.DB = DbName
  frmrepairdatabase.Show vbModal
  Unload frmrepairdatabase
  Resume
 End If
 TEXT$(1) = "OrigDatabase: " & DbName
 TEXT$(2) = "User: " & UserID
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "DaoRoutines - CreateTableAttachment"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub

Public Sub CreateTableAttachment(OrigDatabase As String, AttachDatabase As String, AttachTable As String, TableName As String)
'*************************************************
' Programmer: Lynn Owens
' Date Created: 03-13-1997
'
' Modified Dates: 03-14-1997 Programmer: Lynn Owens
'                 03-18-1997 Programmer: Lynn Owens
'                 03-24-1997 Programmer: Lynn Owens
'
' Description: This routine will create an attachment
'  from a database to a table in another database
'
' Parameters In:
'  OrigDatabase: The Main Database where the attachment goes to
'  AttachDatabase:  The Database that contains the table to attach from
'  AttachTable:  the name of the table in the origdatabase (to)
'  TableName:  The name of the table to attach (from)
'
' Parameters Out: (none)
'
' Example:
' this example will attach custar from the agvance accounting database to the agvance precision pack database
' call CreateTableAttachment("d:\avdata\agvance.mdb",d:\avdata\agvacct.mdb","Custar","Custar")
'
'******************************************************************************
 Dim dbpmCAtt As database
 Dim SaveMessage$
 Dim tbdef As TableDef
 On Local Error GoTo LeHandle
 Call DeleteTableAttachment(OrigDatabase, AttachTable)
 Screen.MousePointer = vbHourglass
' MsgBox AttachTable$
 Set dbpmCAtt = DBEngine(0).OpenDatabase(OrigDatabase, exclusiveyn%, readonlyyn%)
 Set tbdef = dbpmCAtt.CreateTableDef(AttachTable$)
 tbdef.Connect = ";Database=" & AttachDatabase
 tbdef.SourceTableName = TableName
 dbpmCAtt.TableDefs.Append tbdef
 dbpmCAtt.Close: Set dbpmCAtt = Nothing
 Screen.MousePointer = vbDefault
 Exit Sub
 
LeHandle:
' if table is locked or permission denied wait and do again
'3167 is not really a member of family below - it is a symptom
'we solved by including it for the pause effect
 If Err.Number = 3167 Or Err.Number = 3624 Or Err.Number = 3260 Or Err.Number = 3187 Or Err.Number = 3186 Or Err.Number = 3356 Or Err.Number = 3189 Or Err.Number = 3045 Then
   'SaveMessage$ = DisplayStatus(4, "Waiting For Database Access....")
   Call WaitSecs(2)
   'SaveMessage$ = DisplayStatus(4, SaveMessage$)
   Resume
 End If
 TEXT$(1) = "OrigDatabase: " & OrigDatabase
 TEXT$(2) = "AttachDatabase: " & AttachDatabase
 TEXT$(3) = "AttachTable: " & AttachTable
 TEXT$(4) = "tablename: " & TableName
 TEXT$(5) = ""
 GMODNAME$ = "DaoRoutines - CreateTableAttachment"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub

Public Sub DeleteTableAttachment(OrigDatabase As String, TableName As String)
'*************************************************
' Programmer: Lynn Owens
' Date Created: 03-13-1997
'
' Modified Dates: 03-14-1997 Programmer: Lynn Owens
'                 03-18-1997 Programmer: Lynn Owens
'                 03-24-1997 Programmer: Lynn Owens
'                 06-17-1997 programmer: Lynn Owens
'                 11-13-1997 programmer: Lynn Owens
'
' Description: This routine will delete an attachment
'  from a database to a table in another database
'
' Parameters In:
'  OrigDatabase: The Main Database where the attachment goes from
'  TableName:  The name of the table to un-attach
'
' Parameters Out: (none)
'
' Example:
' this example will attach cust ar from the agvance accounting database to the agvance percision pack database
' call DeleteTableAttachment("d:\avdata\agvance.mdb","Custar")
'
'******************************************************************************
 Dim dbpmAtt As database
 Dim SaveMessage$
 On Local Error GoTo LeHandle
1 Screen.MousePointer = vbHourglass
2 Set dbpmAtt = DBEngine(0).OpenDatabase(OrigDatabase, exclusiveyn%, readonlyyn%)
3 dbpmAtt.TableDefs.Delete TableName
4 dbpmAtt.Close: Set dbpmAtt = Nothing
5 Screen.MousePointer = vbDefault
 Exit Sub
 
LeHandle:
' if table is locked or permission denied wait and do again
'MsgBox Error$
 If Erl = 3 And Err.Number = 3265 Then Resume 4
 '3167 is not really a member of family below - it is a symptom
 'we solved by including it for the pause effect
 If Err.Number = 3167 Or Err.Number = 3624 Or Err.Number = 3260 Or Err.Number = 3187 Or Err.Number = 3186 Or Err.Number = 3356 Or Err.Number = 3189 Then
   'SaveMessage$ = DisplayStatus(4, "Waiting For Database Access....")
   Call WaitSecs(2)
   'SaveMessage$ = DisplayStatus(4, SaveMessage$)
   Resume
 End If
 TEXT$(1) = "OrigDatabase: " & OrigDatabase
 TEXT$(2) = "Tablename: " & TableName
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = "DaoRoutines - DeleteTableAttachment"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())
End Sub


