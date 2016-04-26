VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmdisplay_database 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Database Information"
   ClientHeight    =   4710
   ClientLeft      =   210
   ClientTop       =   2235
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   9150
   Begin FPSpread.vaSpread GridFields 
      Height          =   4095
      Left            =   1620
      OleObjectBlob   =   "frmdispl.frx":0000
      TabIndex        =   1
      Top             =   45
      Width           =   7515
   End
   Begin VB.ListBox LstTables 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   1530
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   385
      Left            =   8130
      TabIndex        =   2
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Label lblnumrecs 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   4290
      TabIndex        =   4
      Top             =   4320
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Records = "
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   4335
      Width           =   1485
   End
End
Attribute VB_Name = "frmdisplay_database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index$(50), indexcnt%
Dim DbName As String

 
Property Get GetDbname() As String
 GetDbname = DbName
End Property

Private Sub put_grid(RS As Recordset, my_tabledef As TableDef)
 Dim my_field As Field
 Dim Row%, i%
 i% = 0
 Row% = 0
 RS.MoveFirst
 While Not RS.EOF
  Row% = Row% + 1
  GridFields.Row = Row%
  For Each my_field In my_tabledef.Fields
   i% = i% + 1
   GridFields.Col = i%
   If RS(my_field.Name).Value <> "" Then
     GridFields.TEXT = RS(my_field.Name).Value
   End If
  Next
  i% = 0
  RS.MoveNext
 Wend
End Sub

Property Let setDbname(Thedb As String)
 DbName = Thedb
End Property
Private Sub cmdok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 Dim dbpm As database
 Dim my_tabledef As TableDef
 Call centermdiform(Me, mdimain, 0, 0)
 Show
 Screen.MousePointer = vbHourglass
 Set dbpm = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%) ', gConnect)
 For Each my_tabledef In dbpm.TableDefs
  If left$(my_tabledef.Name, 4) <> "MSys" And UCase$(left$(my_tabledef.Name, 7)) <> "YLDDATA" Then
    If my_tabledef.Connect = "" Then
      LstTables.AddItem my_tabledef.Name
    End If
  End If
 Next
 dbpm.Close: Set dbpm = Nothing
 Screen.MousePointer = vbDefault
 If LstTables.ListCount > 0 Then LstTables.ListIndex = 0
End Sub



Private Sub Gridfields_DblClick(ByVal Col As Long, ByVal Row As Long)
 Dim i%, indx$, colcnt%
 If Row = 0 Then Exit Sub
' If Not AccessAllowed(1, False) Then Exit Sub
 If indexcnt% = 0 Then
   MsgBox "You may not edit data in this table", vbOKOnly + vbInformation, Me.Caption
   Exit Sub
 End If
 Load frmedit_data
 frmedit_data.setDbname = DbName
 indx$ = LstTables.TEXT & Chr$(9)
 For i% = 1 To indexcnt%
  For colcnt% = 1 To GridFields.MaxCols
   GridFields.Row = 0
   GridFields.Col = colcnt%
   If GridFields.TEXT = Index$(i%) Then
     GridFields.Row = Row
     indx$ = indx$ & GridFields.TEXT & Chr$(9)
   End If
  Next colcnt%
 Next i%
 frmedit_data.Tag = indx$
 frmedit_data.Show vbModal
End Sub


Private Sub LstTables_Click()
 Dim TABLENUM%, i%, Row%
 Dim RS As Recordset
 Dim dbpm As database
 Dim my_tabledef As TableDef
 Dim my_field As Field
 Dim my_index As Index
 Screen.MousePointer = vbHourglass
 On Error GoTo ehandle
 i% = 0
 indexcnt% = 0
 GridFields.ReDraw = False
 GridFields.MaxRows = 1
 GridFields.MaxCols = 0

 Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%) ', gConnect)
 Set my_tabledef = dbpm.TableDefs(Trim$(LstTables.TEXT))
 Set RS = dbpm.OpenRecordset(LstTables.TEXT, dbOpenTable)
 For Each my_index In my_tabledef.Indexes
  If UCase$(my_index.Name) = "PRIMARYKEY" Then
    For Each my_field In my_index.Fields
     indexcnt% = indexcnt% + 1
     Index$(indexcnt%) = my_field.Name
    Next
  End If
 Next
 For Each my_field In my_tabledef.Fields
  i% = i% + 1
  GridFields.Row = -1
  GridFields.Col = i%
  GridFields.MaxCols = i%
  GridFields.Lock = True
  GridFields.SetText i%, 0, my_field.Name
  If Len(my_field.Name) < my_field.size Then
    GridFields.ColWidth(i%) = my_field.size * 125
   Else
    GridFields.ColWidth(i%) = Len(my_field.Name) * 125
  End If
  If my_field.Type = dbBoolean Then
    GridFields.CellType = 10 'SS_CELL_TYPE_CHECKBOX
  End If
 Next
 If my_tabledef.RecordCount = 0 Then
   RS.Close: Set RS = Nothing
   dbpm.Close: Set dbpm = Nothing
   lblnumrecs.Caption = 0
   GridFields.MaxRows = 0
   Screen.MousePointer = vbDefault: Exit Sub
   Exit Sub
 End If
 RS.MoveLast
 lblnumrecs.Caption = RS.RecordCount
 'If my_tabledef.RecordCount > 0 Then GridFields.MaxRows = my_tabledef.RecordCount
 GridFields.MaxRows = my_tabledef.RecordCount
 GridFields.Row = 0
 Call put_grid(RS, my_tabledef)
 RS.Close: Set RS = Nothing
 dbpm.Close: Set dbpm = Nothing
 GridFields.ReDraw = True
 Screen.MousePointer = vbDefault
Exit Sub
 
ehandle:
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())
Exit Sub

End Sub
