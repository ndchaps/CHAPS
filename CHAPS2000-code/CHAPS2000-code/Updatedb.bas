Attribute VB_Name = "Update_Databases"
Option Explicit
' Dim ws As Workspace
 Dim DB As database
 Public Enum UpdateAction
  AddField = 1
  EditField = 2
  DeleteField = 3
  AddIndex = 4
  EditIndex = 5
  DeleteIndex = 6
  AddFieldToIndex = 7
  DeleteFieldInIndex = 8
  AddRelation = 9
  DeleteRelation = 10
  AddTableAndField = 11
  DeleteTable = 12
  Findrelation = 13
  ChangeLengthOfTextField = 14
 End Enum
Private Sub add_Field(TheTable$, thefield$, prop$(), exitcode%)
 Dim MyTable As TableDef
 Set MyTable = DB.TableDefs(TheTable$)
 Call create_field_def(MyTable, thefield$, prop(), exitcode%)
End Sub

Private Sub Add_Field_Index(TheTable$, indexname$, indexfieldname$, ReturnCode%)
 Dim FieldName$(50)
 Dim MyIndex As Index
 Dim Myfield As Field
 Dim MyTable As TableDef
 Dim t As Integer, hmfield As Integer
 hmfield = 0
 Set MyTable = DB.TableDefs(TheTable$)
 Set MyIndex = MyTable.CreateIndex(indexname$)
 
 For Each Myfield In MyIndex.Fields
    hmfield = hmfield + 1
    FieldName$(t) = Myfield.Name
 Next
 MyTable.Indexes.Delete indexname$
 
 Set Myfield = MyIndex.CreateField(indexfieldname$)
 MyIndex.Fields.Append Myfield
 
 For t = 1 To hmfield
  Set Myfield = MyIndex.CreateField(FieldName$(t))
  MyIndex.Fields.Append Myfield
 Next t

 MyTable.Indexes.Append MyIndex
End Sub
 
 Sub Add_Relation(reltype$, table1$, table2$, field1$(), Field2$(), PropArray$(), ReturnCode%)
  Dim i%
  Dim Myfield As Field
  Dim MyField2 As Field
  Dim myrelation As Relation
  On Local Error Resume Next
  Set myrelation = DB.CreateRelation(reltype$)
  If Err.Number <> 0 Then ReturnCode% = Err.Number: Exit Sub
  myrelation.Table = table1$
  myrelation.ForeignTable = table2$
  myrelation.Attributes = Val(PropArray$(0))
  For i% = 1 To Val(field1$(0))
    Set Myfield = myrelation.CreateField(field1$(i%))
    Myfield.ForeignName = Field2$(i%)
    myrelation.Fields.Append Myfield
    If Err.Number <> 0 Then ReturnCode% = Err.Number: Exit Sub
  Next i%
  DB.Relations.Append myrelation
  If Err.Number <> 0 Then ReturnCode% = Err.Number
End Sub

Private Sub ChangeFieldLength(TableName$, FieldName$, PropArray$(), ReturnCode%)
 Dim fld As Field
 Dim NewPropertyArray(5) As String
 ReturnCode% = 0
 Set fld = DB.TableDefs(TableName$).Fields(FieldName$)
 If fld.Type <> dbText Then Exit Sub
 NewPropertyArray(0) = "Text"
 If fld.AllowZeroLength Then
   NewPropertyArray(1) = "-1"
  Else
   NewPropertyArray(1) = "-1"
 End If
 If fld.Required Then
   NewPropertyArray(2) = "-1"
  Else
   NewPropertyArray(2) = "0"
 End If
 NewPropertyArray(3) = PropArray$(3)
 Call add_Field(TableName$, "SSITemp", NewPropertyArray(), ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' copy the data from the oldfield to the ssitempfield
 '
 DB.Execute "UPDATE " & TableName$ & " SET " & TableName$ & ".SsiTemp = [" & TableName$ & "]![" & FieldName$ & "];", dbFailOnError
 '
 ' delete the old field from the table
 '
 Call delete_Field(TableName$, FieldName$, ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' add the old field back in with the new length
 '
 Call add_Field(TableName$, FieldName$, NewPropertyArray(), ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' copy the data from the ssitemp field to the origional field
 '
 DB.Execute "UPDATE " & TableName$ & " SET " & TableName$ & "." & FieldName$ & " = [" & TableName$ & "]![ssitemp];", dbFailOnError
 '
 ' delete the ssitemp field
 '
 Call delete_Field(TableName$, "SSITemp", ReturnCode%)
End Sub

Private Sub ChangeFieldType(TableName$, FieldName$, PropArray$(), ReturnCode%)
 Dim fld As Field
 Dim NewPropertyArray(5) As String
 ReturnCode% = 0
 Set fld = DB.TableDefs(TableName$).Fields(FieldName$)
 NewPropertyArray(0) = PropArray$(0)
 If fld.AllowZeroLength Then
   NewPropertyArray(1) = "-1"
  Else
   NewPropertyArray(1) = "-1"
 End If
 If fld.Required Then
   NewPropertyArray(2) = "-1"
  Else
   NewPropertyArray(2) = "0"
 End If
 NewPropertyArray(3) = PropArray$(3)
 Call add_Field(TableName$, "SSITemp", NewPropertyArray(), ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' copy the data from the oldfield to the ssitempfield
 '
 DB.Execute "UPDATE " & TableName$ & " SET " & TableName$ & ".SsiTemp = [" & TableName$ & "]![" & FieldName$ & "];", dbFailOnError
 '
 ' delete the old field from the table
 '
 Call delete_Field(TableName$, FieldName$, ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' add the old field back in with the new length
 '
 Call add_Field(TableName$, FieldName$, NewPropertyArray(), ReturnCode%)
 If ReturnCode% <> 0 Then Exit Sub
 '
 ' copy the data from the ssitemp field to the origional field
 '
 DB.Execute "UPDATE " & TableName$ & " SET " & TableName$ & "." & FieldName$ & " = [" & TableName$ & "]![ssitemp];", dbFailOnError
 '
 ' delete the ssitemp field
 '
 Call delete_Field(TableName$, "SSITemp", ReturnCode%)
End Sub


Private Sub Delete_Relation(RELNAME$, ReturnCode%)
 On Local Error Resume Next
 DB.Relations.Delete RELNAME$
 ReturnCode% = Err.Number
End Sub

Private Sub Delete_Field_Index(TheTable$, indexname$, indexfieldname$, ReturnCode%)
 Dim FieldName$(50)
 Dim MyIndex As Index
 Dim newfield As Field
 Dim MyTable As TableDef
 Dim t As Integer, hmfield As Integer
 Set MyTable = DB.TableDefs(TheTable$)
 Set MyIndex = MyTable.Indexes(Trim$(indexname$))
 hmfield = 0
 For Each newfield In MyIndex.Fields
  If newfield.Name <> indexfieldname$ Then
    hmfield = hmfield + 1
    FieldName$(t) = newfield.Name
  End If
 Next
 'MyIndex.field indexfieldname$
 MyTable.CreateIndex (indexname$)
 For t = 1 To hmfield
  Set newfield = MyIndex.CreateField(FieldName$(t))
  MyIndex.Fields.Append newfield
 Next t
 'Set newfield = myindex.CreateField(indexfieldname$)
 'myindex.Fields.Append newfield
 DB.TableDefs(TheTable$).Indexes.Refresh
End Sub

Private Sub add_index(TableName$, indexname$, FieldName$(), PropArray$(), exitcode%)
 Dim newindex As Index
 Dim newfield As Field
 Dim t As Integer
 On Local Error Resume Next
 Set newindex = DB.TableDefs(TableName$).CreateIndex(Trim$(indexname$))
 If Err.Number <> 0 Then exitcode% = Err.Number: Exit Sub
 newindex.Primary = Val(PropArray$(0))
 newindex.Unique = Val(PropArray$(1))
 newindex.Required = Val(PropArray$(2))
 newindex.IgnoreNulls = Val(PropArray$(3))
 For t = 1 To Val(FieldName$(0))
  Set newfield = newindex.CreateField(FieldName$(t))
  newindex.Fields.Append newfield
  If Err <> 0 Then exitcode% = Err.Number: Exit Sub
 Next t
 If Err <> 0 Then exitcode% = Err: Exit Sub
 DB.TableDefs(TableName$).Indexes.Append newindex
 If Err <> 0 Then exitcode% = Err
End Sub

Private Sub create_field_def(TableName As TableDef, thefield$, PropArray$(), ReturnCode%)
 Dim newfield As Field
 Dim typevar
 On Local Error Resume Next
  Select Case PropArray$(0)
  Case "Text"
   typevar = dbText
  Case "Memo"
   typevar = dbMemo
  Case "Yes/No"
   typevar = dbBoolean
  Case "Currency"
   typevar = dbCurrency
  Case "Date/Time"
   typevar = dbDate
  Case "Integer"
   typevar = dbInteger
  Case "Long Integer"
   typevar = dbLong
  Case "Double"
   typevar = dbDouble
  Case "Single"
   typevar = dbSingle
  Case "Byte"
   typevar = dbByte
 End Select

 Set newfield = TableName.CreateField(Trim$(thefield$), typevar)
 Select Case PropArray$(0)
  Case "Text"
   newfield.size = Val(PropArray$(3))
   newfield.AllowZeroLength = Val(PropArray$(1))
   newfield.Required = Val(PropArray$(2))
  Case "Memo"
   newfield.AllowZeroLength = Val(PropArray$(1))
   newfield.Required = Val(PropArray$(2))
  Case "Yes/No"
    newfield.Required = Val(PropArray$(2))
  Case "Currency"
   newfield.Required = Val(PropArray$(2))
  Case "Date/Time"
   newfield.Required = Val(PropArray$(2))
  Case "Integer"
   newfield.Required = Val(PropArray$(2))
  Case "Long Integer"
   newfield.Required = Val(PropArray$(2))
  Case "Double"
   newfield.Required = Val(PropArray$(2))
  Case "Single"
   newfield.Required = Val(PropArray$(2))
  Case "Byte"
   newfield.Required = Val(PropArray$(2))
 End Select
 TableName.Fields.Append newfield
 If Err <> 0 Then ReturnCode% = Err
End Sub

Private Sub delete_Field(TheTable$, thefield$, exitcode%)
 Dim MyTable As TableDef
 On Local Error Resume Next
 Set MyTable = DB.TableDefs(TheTable$)
 MyTable.Fields.Delete thefield$
 If Err.Number <> 0 Then exitcode% = Err.Number: Exit Sub
 DB.TableDefs(TheTable$).Fields.Refresh
 exitcode% = Err.Number
End Sub

Private Sub delete_Index(TheTable$, theindex$, exitcode%)
 Dim MyTable As TableDef
 Set MyTable = DB.TableDefs(TheTable$)
 MyTable.Indexes.Delete theindex$
 DB.TableDefs(TheTable$).Indexes.Refresh
End Sub

Private Sub Delete_table(TableName$, exitcode%)
 Dim MyTabledef As TableDef
 Set MyTabledef = DB.CreateTableDef(TableName$)
 On Local Error Resume Next
 DB.Execute ("drop table " & TableName$)
 If Err <> 0 Then exitcode% = Err
End Sub

Private Sub Edit_Field(TheTable$, thefield$, prop$(), exitcode%)
 Dim newfield As Field
 On Local Error Resume Next
 Set newfield = DB.TableDefs(TheTable$).Fields(Trim$(thefield$))
 If Err <> 0 Then exitcode% = Err.Number: Exit Sub
 If prop$(5) <> thefield$ Then newfield.Name = prop$(5)
 Select Case prop$(0)
   Case "Text"
'   newfield.Type = dbText
   'newfield.Size = Val(prop$(3))
   newfield.AllowZeroLength = Val(prop$(1))
   newfield.Required = Val(prop$(2))
  ' newfield.OrdinalPosition = Val(prop$(4))
  Case "Memo"
   'newfield.Type = dbMemo
   newfield.AllowZeroLength = Val(prop$(1))
   newfield.Required = Val(prop$(2))
  Case "Yes/No"
    'newfield.Type = dbBoolean
    newfield.Required = Val(prop$(2))
  Case "Currency"
   'newfield.Type = dbCurrency
   newfield.Required = Val(prop$(2))
  Case "Date/Time"
   'newfield.Type = dbDate
   newfield.Required = Val(prop$(2))
  Case "Integer"
   'newfield.Type = dbInteger
   newfield.Required = Val(prop$(2))
  Case "Long Integer"
   'newfield.Type = dbLong
   newfield.Required = Val(prop$(2))
  Case "Double"
   'newfield.Type = dbDouble
   newfield.Required = Val(prop$(2))
  Case "Single"
   newfield.Type = dbSingle
   newfield.Required = Val(prop$(2))
  Case "Byte"
   'newfield.Type = dbByte
   newfield.Required = Val(prop$(2))
 End Select
 DB.TableDefs(TheTable$).Fields.Refresh
 If Err <> 0 Then exitcode% = Err.Number
End Sub

Private Sub Edit_Index(TheTable$, indexname$, PropArray$(), ReturnCode%)
 Dim newindex As Index
 On Local Error Resume Next
 Set newindex = DB.TableDefs(TheTable$).Indexes(Trim$(indexname$))
 newindex.Unique = Val(PropArray$(1))
 newindex.Required = PropArray$(2)
 newindex.IgnoreNulls = PropArray$(3)
 DB.TableDefs(TheTable$).Fields.Refresh
 If Err <> 0 Then ReturnCode% = Err.Number
End Sub

Public Function FIND_RELATION(MAINTABLE$, RELTable$, MAINFIELDS$(), RELFIELDS$()) As String
 Dim myrelation As Relation
 Dim Myfield As Field
 Dim t As Integer
 For Each myrelation In DB.Relations
  If UCase$(myrelation.Table) = UCase$(MAINTABLE$) And UCase$(myrelation.ForeignTable) = UCase$(RELTable$) Then
    FIND_RELATION = myrelation.Name
    t = 1
    For Each Myfield In myrelation.Fields
     If t > Val(MAINFIELDS$(0)) Then
       FIND_RELATION = ""
       Exit For
     End If
     If UCase$(Myfield.Name) <> UCase$(MAINFIELDS$(t)) Or UCase$(Myfield.ForeignName) <> UCase$(RELFIELDS$(t)) Then
       FIND_RELATION = ""
       Exit For
     End If
     t = t + 1
    Next
    If FIND_RELATION <> "" Then Exit For
  End If
 Next
End Function

Public Sub Update_Database(DbName$, TableName$, FieldName$, indexname$, RELNAME$, tablename2$, Field$(), Field2$(), PropArray$(), ActionCode%, ReturnCode%, Connect As String)
 ' this procedure will preform a action on a database
 ' dbname$      - name of the database ex. d:\avdata\agvance.mdb
 ' tablename$   - name of the table ex. grower
 ' fieldname$   - name of a field in the table ex. growid
 ' indexname$   - name of a index ex. PrimaryKey
 ' RELNAME$     - name of the relation(sent in and sent back depending on action) ex. fldplantoplansplt
 ' tablename2$  - name of the secon table in a relation ex. if plansplt was the second table in a relation it would be sent in here
 ' field$()     - name of fields in a relation or index ex. (0) = the number of fields and the (1-hmfields) are the field names (index or relations)
 ' field2$()    - name of second table fields in a relation or index ex. (0) = the number of fields and the (1-hmfields) are the field names (relations)
 ' proparray$() - a array of values sent in for properties or attributes of the action(0-hmproperties) list of each property or action below
 ' actioncode%  - The Type of action (listed below are the valid actions
 ' returncode%  - If A Error Occurs The VB error number is returned in this var
 '
 ' valid actions
 '
 ' 1 - Add A Field To A Table
 '     Send In:
 '             dbname$
 '             Tablename$
 '             FieldName$
 '             proparray$()
 '
 ' 2 - Edit A field In A Table
 '     Send In:
 '             dbname$
 '             Tablename$
 '             Fieldname$
 '             proparray$()
 '
 ' 3 - Delete A Field From Table
 '     Send In:
 '             Dbname$
 '             tablename$
 '             fieldname$
 ' 4 - Add A Index To A Table
 '     Send In:
 '             Dbname$
 '             Tablename$
 '             IndexName$
 '             field$()   - a list of fields in the index( sub 0 is hmfields) need at leat one
 ' 5 - Edit A INdex In Table
 '     Send In:
 '             Dbname$
 '             tablename$
 '             Indexname$
 '             proparray$()
 ' 6 - Delete A Index In A Table
 '     Send In:
 '             dbname$
 '             tablename$
 '             Indexname$
 ' 7 - Add A Field To A Index
 '     Send in:
 '             Dbname$
 '             tablename$
 '             indexname$
 '             fieldname$ - for index
 ' 8 - delete a field in an index
 '     send in:
 '             dbname$
 '             tablename$
 '             fieldname$
 '
 ' 9 - add a relation
 '     send in:
 '              dbname$
 '              tablename$ - main table of the relation
 '              relname$ - name of the relation to delete
 '              tabalename2$ - the foreign tables name
 '              field$()     - a list of fields from the main table to link to the foreign table (sub 0 is hmfields)
 '              field2$()     - a list of fields from the foreign table to link to the foreign table
 ' 10 - delate a relation
 '      send in:
 '              dbanme$
 '              relname$ - name of the relation to delete
 ' 11 - Add A Table
 '      Send in:
 '              dbname$
 '              Tablename$
 '              fieldname$  - name of one field to add(need one to add a table)
 '              proparray() - field properties
 ' 12 - delete a table
 '      Send in:
 '              dbname$
 '              tablename$
 ' 13 - Find A Relation and return the relation name
 '      send in:
 '               dbname$
 '               tablename$  - main table
 '               relname$    - sent back the relation name or a blank if it does not exist
 '               tablename2$ - the foreign table in the relation
 '
 ' 14 - Change the length of a text field
 '      send in:
 '              dbname$
 '              tablename$ - The table that contains the field we want to change the length
 '              proparray() - the array of properties (3) is the size
 ' 15 - Change the length of a text field
 '      send in:
 '              dbname$
 '              tablename$ - The table that contains the field we want to change the length
 '              proparray() - the array of properties (3) is the size
 '
 ' PropArray(0) = "type"
 '          (1) =
 '          (2) =
 '          (3) = Size
 '
 

 On Local Error GoTo LeHandle
 ReturnCode% = 0
 Screen.MousePointer = vbHourglass
' Set ws = DBEngine(0)
 Set DB = DBEngine(0).OpenDatabase(DbName$, True, False)
 Select Case ActionCode%
  Case 1
   Call add_Field(TableName$, FieldName$, PropArray$(), ReturnCode%)
  Case 2
   Call Edit_Field(TableName$, FieldName$, PropArray$(), ReturnCode%)
  Case 3
   Call delete_Field(TableName$, FieldName$, ReturnCode%)
  Case 4
   Call add_index(TableName$, indexname$, Field$(), PropArray$(), ReturnCode%)
  Case 5
   Call Edit_Index(TableName$, indexname$, PropArray$(), ReturnCode%)
  Case 6
   Call delete_Index(TableName$, indexname$, ReturnCode%)
  Case 7
   Call Add_Field_Index(TableName$, indexname$, FieldName$, ReturnCode%)
  Case 8
   Call Delete_Field_Index(TableName$, indexname$, FieldName$, ReturnCode%)
  Case 9
   Call Add_Relation(RELNAME$, TableName$, tablename2$, Field$(), Field2$(), PropArray$(), ReturnCode)
  Case 10
   Call Delete_Relation(RELNAME$, ReturnCode)
  Case 11
   Call Add_Table(TableName$, FieldName$, PropArray$(), ReturnCode%)
  Case 12
   Call Delete_table(TableName$, ReturnCode%)
  Case 13
   RELNAME$ = FIND_RELATION(TableName$, tablename2$, Field$(), Field2$())
  Case 14
   Call ChangeFieldLength(TableName$, FieldName$, PropArray$(), ReturnCode%)
  Case 15
   Call ChangeFieldType(TableName$, FieldName$, PropArray$(), ReturnCode%)

 End Select
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
 Exit Sub
LeHandle:
 Screen.MousePointer = vbDefault
 ReturnCode% = Err.Number
End Sub
Private Sub Add_Table(TableName$, FieldName$, PropArray$(), exitcode%)
 'dim mydatabase As Database
 Dim MyTabledef As TableDef
 Dim Myfield As Field
 On Local Error Resume Next
 Set MyTabledef = DB.CreateTableDef(TableName$)
 Call create_field_def(MyTabledef, FieldName$, PropArray$(), exitcode%)
 DB.TableDefs.Append MyTabledef
 If Err <> 0 Then exitcode% = Err
End Sub

