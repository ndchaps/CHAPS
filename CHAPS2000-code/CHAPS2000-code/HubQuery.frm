VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form FrmHubQuery 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query"
   ClientHeight    =   5205
   ClientLeft      =   930
   ClientTop       =   1665
   ClientWidth     =   7800
   Icon            =   "HubQuery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7800
   Tag             =   "First time"
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   330
      Left            =   5745
      TabIndex        =   15
      Top             =   3165
      Width           =   840
   End
   Begin VB.CommandButton cmdQueryDelete 
      Caption         =   "Delete Query"
      Height          =   330
      Left            =   6630
      TabIndex        =   16
      Top             =   3165
      Width           =   1110
   End
   Begin VB.ComboBox cboQueryNameList 
      Height          =   315
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3165
      Width           =   3270
   End
   Begin VB.CommandButton cmdQuerySave 
      Caption         =   "&Save Query"
      Height          =   330
      Left            =   4590
      TabIndex        =   14
      Top             =   3165
      Width           =   1110
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Height          =   390
      Left            =   6630
      TabIndex        =   21
      Top             =   4635
      Width           =   960
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show &Results"
      Height          =   390
      Left            =   4845
      TabIndex        =   20
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build &Query"
      Height          =   390
      Left            =   3795
      TabIndex        =   19
      Top             =   4635
      Width           =   1005
   End
   Begin VB.TextBox txtQuery 
      Height          =   915
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   3540
      Width           =   7710
   End
   Begin TabDlg.SSTab SSTabQuery 
      Height          =   3120
      Left            =   45
      TabIndex        =   34
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   5503
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Select What"
      TabPicture(0)   =   "HubQuery.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Criteria"
      TabPicture(1)   =   "HubQuery.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCriteria"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "HubQuery.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraResults"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCriteria 
         ClipControls    =   0   'False
         Height          =   2700
         Left            =   -74925
         TabIndex        =   22
         Top             =   345
         Width           =   7560
         Begin FPSpread.vaSpread grdCriteria 
            Height          =   1215
            Left            =   90
            OleObjectBlob   =   "HubQuery.frx":0496
            TabIndex        =   28
            Top             =   1395
            Width           =   7365
         End
         Begin VB.TextBox txtDistinctSql 
            Height          =   270
            Left            =   4710
            TabIndex        =   27
            Top             =   1065
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.OptionButton optAny 
            Caption         =   "Any are true"
            Height          =   195
            Left            =   3405
            TabIndex        =   26
            Top             =   1125
            Width           =   1230
         End
         Begin VB.OptionButton optAll 
            Caption         =   "All are true"
            Height          =   195
            Left            =   2130
            TabIndex        =   25
            Top             =   1125
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.Label lblHint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"HubQuery.frx":071B
            ForeColor       =   &H00FF0000&
            Height          =   855
            Index           =   3
            Left            =   90
            TabIndex        =   23
            Top             =   180
            Width           =   7365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Match when:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   525
            TabIndex        =   24
            Top             =   1110
            Width           =   1455
         End
      End
      Begin VB.Frame fraResults 
         Height          =   2700
         Left            =   -74925
         TabIndex        =   29
         Top             =   345
         Width           =   7560
         Begin FPSpread.vaSpread grdDisplay 
            Bindings        =   "HubQuery.frx":08A1
            Height          =   1665
            Left            =   105
            OleObjectBlob   =   "HubQuery.frx":08B5
            TabIndex        =   30
            Top             =   210
            Width           =   7350
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   2580
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   2  'Snapshot
            RecordSource    =   ""
            Top             =   1290
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print Results"
            Height          =   330
            Left            =   6345
            TabIndex        =   32
            ToolTipText     =   "Print results to the current Window's printer"
            Top             =   1935
            Width           =   1110
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "Save Results"
            Height          =   315
            Left            =   6345
            TabIndex        =   33
            ToolTipText     =   "Save results as a Tab Delimited Text File"
            Top             =   2280
            Width           =   1110
         End
         Begin VB.Label lblHint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"HubQuery.frx":0ACF
            ForeColor       =   &H00FF0000&
            Height          =   660
            Index           =   5
            Left            =   90
            TabIndex        =   31
            Top             =   1935
            Width           =   6180
         End
      End
      Begin VB.Frame fraSelect 
         Height          =   2700
         Left            =   75
         TabIndex        =   0
         Top             =   345
         Width           =   7560
         Begin VB.Frame Frame1 
            Caption         =   "Select a Table"
            Height          =   2430
            Left            =   90
            TabIndex        =   1
            Top             =   165
            Width           =   2175
            Begin VB.ComboBox cboTables 
               Height          =   315
               Left            =   105
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   240
               Width           =   1980
            End
            Begin VB.Label lblHint 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"HubQuery.frx":0BDC
               ForeColor       =   &H00FF0000&
               Height          =   1710
               Index           =   2
               Left            =   90
               TabIndex        =   3
               Top             =   630
               Width           =   1995
            End
         End
         Begin VB.Frame fraFields 
            Caption         =   "Data Fields"
            Height          =   2430
            Left            =   2325
            TabIndex        =   4
            Top             =   165
            Width           =   2250
            Begin MhglbxLib.Mh3dList lstFields 
               Height          =   1380
               Left            =   90
               TabIndex        =   5
               Top             =   225
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
               _ExtentY        =   2434
               _StockProps     =   79
               Caption         =   "Mh3dList1"
               ForeColor       =   -2147483640
               BackColor       =   16777215
               TintColor       =   16711935
               Caption         =   "Mh3dList1"
               ColTitleButtons =   0   'False
               BevelStyleInner =   0
               BevelSizeInner  =   0
               BorderType      =   1
               BorderColor     =   0
               Case            =   0
               Col             =   0
               ColCharacter    =   9
               ColScale        =   0
               ColSizing       =   0
               DividerStyle    =   0
               FillColor       =   16777215
               FontStyle       =   0
               LightColor      =   16777215
               MultiSelect     =   0
               PictureHeight   =   0
               PictureWidth    =   0
               AdjustHeight    =   0
               ScrollBars      =   1
               ShadowColor     =   8421504
               WallPaper       =   0
               Sorted          =   0   'False
               TextColor       =   0
               WrapList        =   0   'False
               WrapWidth       =   0
               ColInstr        =   0
               TitleHeight     =   -1
               TitleFontBold   =   0   'False
               TitleFontItalic =   0   'False
               TitleFontName   =   "MS Sans Serif"
               TitleFontSize   =   8.25
               TitleFontStrike =   0   'False
               TitleFontUnder  =   0   'False
               TitleFontStyle  =   0
               TitleBevelStyle =   0
               TitleBevelSize  =   0
               TitleColor      =   0
               FocusColor      =   0
               HighColor       =   16777215
               VirtualList     =   0   'False
               BufferSize      =   100
               SortOrder       =   ""
               SelectedColor   =   8388608
               Transparent     =   0   'False
               TransparentColor=   0
               TitleFillColor  =   12632256
               Platform        =   0
               FireDrawItem    =   0   'False
               DrawItemLeft    =   0
               DrawItemRight   =   0
               DataSourceList  =   ""
               ListDividersH   =   -1  'True
               ListDividersV   =   -1  'True
               TitleDividers   =   -1  'True
               DataField       =   ""
               DataFieldCount  =   0
            End
            Begin VB.Label lblHint 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Double Click on the data fields you wish to display in your query."
               ForeColor       =   &H00FF0000&
               Height          =   675
               Index           =   1
               Left            =   90
               TabIndex        =   6
               Top             =   1665
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Selected Fields"
            Height          =   2430
            Left            =   4635
            TabIndex        =   9
            Top             =   165
            Width           =   2835
            Begin MhglbxLib.Mh3dList lstSelFields 
               Height          =   1380
               Left            =   90
               TabIndex        =   10
               Top             =   225
               Width           =   2640
               _Version        =   65536
               _ExtentX        =   4657
               _ExtentY        =   2434
               _StockProps     =   79
               Caption         =   "Mh3dList1"
               BackColor       =   16777215
               TintColor       =   16711935
               Caption         =   "Mh3dList1"
               ColTitleButtons =   0   'False
               BevelStyleInner =   0
               BevelSizeInner  =   0
               BorderType      =   1
               BorderColor     =   0
               Case            =   0
               Col             =   0
               ColCharacter    =   9
               ColScale        =   0
               ColSizing       =   0
               DividerStyle    =   0
               FillColor       =   16777215
               FontStyle       =   0
               LightColor      =   16777215
               MultiSelect     =   0
               PictureHeight   =   0
               PictureWidth    =   0
               AdjustHeight    =   0
               ScrollBars      =   1
               ShadowColor     =   8421504
               WallPaper       =   0
               Sorted          =   0   'False
               TextColor       =   0
               WrapList        =   0   'False
               WrapWidth       =   0
               ColInstr        =   -1
               TitleHeight     =   -1
               TitleFontBold   =   0   'False
               TitleFontItalic =   0   'False
               TitleFontName   =   "MS Sans Serif"
               TitleFontSize   =   8.25
               TitleFontStrike =   0   'False
               TitleFontUnder  =   0   'False
               TitleFontStyle  =   0
               TitleBevelStyle =   0
               TitleBevelSize  =   0
               TitleColor      =   0
               FocusColor      =   0
               HighColor       =   16777215
               VirtualList     =   0   'False
               BufferSize      =   100
               SortOrder       =   ""
               SelectedColor   =   8388608
               Transparent     =   0   'False
               TransparentColor=   0
               TitleFillColor  =   12632256
               Platform        =   0
               FireDrawItem    =   0   'False
               DrawItemLeft    =   0
               DrawItemRight   =   0
               DataSourceList  =   ""
               ListDividersH   =   -1  'True
               ListDividersV   =   -1  'True
               TitleDividers   =   -1  'True
               DataField       =   ""
               DataFieldCount  =   0
            End
            Begin VB.Label lblHint 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Double Click on the data fields you wish to remove from your query."
               ForeColor       =   &H00FF0000&
               Height          =   675
               Index           =   0
               Left            =   90
               TabIndex        =   11
               Top             =   1665
               Width           =   2640
            End
         End
         Begin VB.Frame Frame3 
            Height          =   630
            Left            =   90
            TabIndex        =   7
            Top             =   1275
            Visible         =   0   'False
            Width           =   2115
            Begin MhglbxLib.Mh3dList lstSelTables 
               Height          =   345
               Left            =   90
               TabIndex        =   8
               Top             =   195
               Width           =   1905
               _Version        =   65536
               _ExtentX        =   3360
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Mh3dList1"
               BackColor       =   14215660
               TintColor       =   16711935
               Caption         =   "Mh3dList1"
               ColTitleButtons =   0   'False
               BevelStyleInner =   0
               BevelSizeInner  =   0
               BorderType      =   1
               BorderColor     =   0
               Case            =   0
               Col             =   0
               ColCharacter    =   9
               ColScale        =   0
               ColSizing       =   0
               DividerStyle    =   0
               FontStyle       =   0
               LightColor      =   16777215
               MultiSelect     =   0
               PictureHeight   =   0
               PictureWidth    =   0
               AdjustHeight    =   0
               ScrollBars      =   1
               ShadowColor     =   8421504
               WallPaper       =   0
               Sorted          =   0   'False
               TextColor       =   0
               WrapList        =   0   'False
               WrapWidth       =   0
               ColInstr        =   -1
               TitleHeight     =   -1
               TitleFontBold   =   0   'False
               TitleFontItalic =   0   'False
               TitleFontName   =   "MS Sans Serif"
               TitleFontSize   =   8.25
               TitleFontStrike =   0   'False
               TitleFontUnder  =   0   'False
               TitleFontStyle  =   0
               TitleBevelStyle =   0
               TitleBevelSize  =   0
               TitleColor      =   0
               FocusColor      =   0
               HighColor       =   16777215
               VirtualList     =   0   'False
               BufferSize      =   100
               SortOrder       =   ""
               SelectedColor   =   8388608
               Transparent     =   0   'False
               TransparentColor=   0
               TitleFillColor  =   12632256
               Platform        =   0
               FireDrawItem    =   0   'False
               DrawItemLeft    =   0
               DrawItemRight   =   0
               DataSourceList  =   ""
               ListDividersH   =   -1  'True
               ListDividersV   =   -1  'True
               TitleDividers   =   -1  'True
               DataField       =   ""
               DataFieldCount  =   0
            End
         End
      End
   End
   Begin VB.Label lblSqlQuery 
      Alignment       =   1  'Right Justify
      Caption         =   "Saved Queries"
      Height          =   180
      Left            =   135
      TabIndex        =   12
      Top             =   3210
      Width           =   1065
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"HubQuery.frx":0C88
      ForeColor       =   &H00FF0000&
      Height          =   660
      Index           =   4
      Left            =   45
      TabIndex        =   18
      Top             =   4500
      Width           =   3660
   End
End
Attribute VB_Name = "FrmHubQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mode$

Public Sub ClearGridSelected(spread1 As Control)
 If Not spread1.IsBlockSelected Then Exit Sub
 Screen.MousePointer = vbHourglass
 spread1.ReDraw = False
 spread1.Row = spread1.SelBlockRow
 spread1.Row2 = spread1.SelBlockRow2
 If spread1.Row = -1 Then spread1.Row = 1
 If spread1.Row2 = -1 Then spread1.Row2 = spread1.MaxRows
 spread1.Col = 1
 spread1.col2 = spread1.MaxCols
 spread1.BlockMode = True
 spread1.Action = SS_ACTION_DELETE_ROW
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 spread1.col2 = 0
 spread1.Row2 = 0
 spread1.ReDraw = True
 spread1.Action = ss_ACTION_DESELECT_BLOCK
 Screen.MousePointer = vbDefault
End Sub

Private Sub FormReset()
txtQuery.TEXT = ""
cboTables.ListIndex = -1
Call DisplayGridClear
Call CriteriaGridClear
lstFields.Clear
lstSelTables.Clear
lstSelFields.Clear
SSTabQuery.TabVisible(0) = True
SSTabQuery.TabVisible(1) = True
SSTabQuery.Tab = 0
cmdBuild.Enabled = False
cmdShow.Enabled = False
cmdQueryDelete.Enabled = False
cboQueryNameList.ListIndex = -1
cboTables.SetFocus
End Sub

Private Function GetFieldType(TableFieldName$) As Integer
'TableFieldName$ format: <table>.<field>
'returns field type per dbText, dbDate, dbInteger etc codes
Dim dbpm As database, tblDef As TableDef
Dim tblName$, fldName$

Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
tblName$ = Left$(TableFieldName$, InStr(TableFieldName$, ".") - 1)
Set tblDef = dbpm.TableDefs(tblName$)
fldName$ = Mid$(TableFieldName$, InStr(TableFieldName$, ".") + 1)
GetFieldType% = tblDef.Fields(fldName$).Type

dbpm.Close: Set dbpm = Nothing
End Function

Private Sub KeywordMargin()
Dim SQL$, i%
Dim Keyword$(5), KeywordAt%

Keyword$(0) = " FROM"
Keyword$(1) = " LEFT JOIN"
Keyword$(2) = " GROUP BY"
Keyword$(3) = " HAVING"
Keyword$(4) = " WHERE"
Keyword$(5) = " ORDER BY"

SQL$ = txtQuery.TEXT
'put SQL keywords at left margin for readability
For i% = 0 To 5
  KeywordAt% = InStr(1, SQL$, Keyword(i%))
  While KeywordAt%
    SQL$ = Mid(SQL$, 1, KeywordAt% - 1) & vbCrLf & Mid(SQL$, KeywordAt% + 1)
    KeywordAt% = InStr(1, SQL$, Keyword(i%))
  Wend
Next i%
txtQuery.TEXT = SQL$
End Sub

Private Sub QueryDelete()
Dim DB As database
Dim rsQuery As Recordset
Dim RESPONSE%

RESPONSE% = MsgBox("Delete Selected Query?", vbInformation + vbYesNo, Me.Caption)
If RESPONSE% = vbYes Then
  Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
  Set rsQuery = DB.OpenRecordset("Query", dbOpenTable)
  rsQuery.Index = "primarykey"
  rsQuery.Seek "=", cboQueryNameList.TEXT
  If Not rsQuery.NoMatch Then
    rsQuery.Delete
  End If
  rsQuery.Close: Set rsQuery = Nothing
  DB.Close: Set DB = Nothing
End If
cboQueryNameList.RemoveItem (cboQueryNameList.ListIndex)
Call FormReset
End Sub

Private Sub QuerySave(Name$, theType$, SQL$)
Dim dbpm As database
Dim rsQuery As Recordset

Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set rsQuery = dbpm.OpenRecordset("Query", dbOpenTable)

rsQuery.Index = "primarykey"
rsQuery.Seek "=", Name$
If Not rsQuery.NoMatch Then
  rsQuery.Edit
Else
  rsQuery.AddNew
End If
rsQuery!Name = Name$
rsQuery!QueryType = theType$
rsQuery!SQL = SQL$
rsQuery.Update
rsQuery.Close: Set rsQuery = Nothing
dbpm.Close: Set dbpm = Nothing
End Sub

Private Sub CriteriaGridClear()
Screen.MousePointer = vbHourglass
With grdCriteria
  .Col = 1: .col2 = 4
  .Row = 1:: .Row2 = 500:
  .BlockMode = True
  .Action = 12 'clear text
  .BlockMode = False
  .Col = 1: .col2 = 1 'clear criteria combo
  .Row = 1: .Row2 = 500
  .BlockMode = True
  .TypeComboBoxList = ""
  .BlockMode = False
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub DisplayGridClear()
Screen.MousePointer = vbHourglass
With grdDisplay
  .Row = 0
  .Col = 1
  .Row2 = 500
  .col2 = 500
  .BlockMode = True
  .Action = 12 'clear text
  .BlockMode = False
  .MaxCols = 2
  .MaxRows = 0
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub CriteriaCBOLoad()
Dim i%, CBO$

For i% = 0 To lstSelFields.ListCount - 1
  CBO$ = CBO$ & lstSelFields.ColList(i%) & Chr$(9)
Next i%

grdCriteria.Col = 1: grdCriteria.col2 = 1
grdCriteria.Row = 1: grdCriteria.Row2 = 500
grdCriteria.BlockMode = True
grdCriteria.TypeComboBoxList = CBO$
grdCriteria.BlockMode = False
End Sub

Property Let setDbname(DB As String)
 Data1.DatabaseName = dbfile$
End Property

Function GetDbname() As String
GetDbname = dbfile$
End Function

Private Function SetVal(TableFieldName$) As String
Dim tblName$, fldName$

tblName$ = Left$(TableFieldName$, InStr(TableFieldName$, ".") - 1)
fldName$ = Mid$(TableFieldName$, InStr(TableFieldName$, ".") + 1)
SetVal = TableFieldName$ & " AS [" & tblName$ & " " & fldName$ & "]"
End Function

Private Function TableExists(TableName$, HmTables, TableArray$()) As Boolean
Dim t As Integer
TableExists = False
For t = 1 To HmTables
  If UCase$(TableName$) = UCase$(TableArray$(t)) Then
    TableExists = True
    Exit For
  End If
Next t
End Function

Private Sub cboQueryNameList_Click()
Dim dbpm As database
Dim rsQueryNames As Recordset
Dim rsCriteria As Recordset
Dim SQL$, table$, SelectClause$, tblfld$, i%

If cboQueryNameList.ListIndex = -1 Then Exit Sub
If mode$ = "QuerySave" Then 'event code fired by updating from query save
  mode$ = ""
  Exit Sub
End If
cmdQuerySave.Enabled = False
cmdQueryDelete.Enabled = True
Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)

Set rsQueryNames = dbpm.OpenRecordset("Query", dbOpenTable)
rsQueryNames.Index = "primarykey"
rsQueryNames.Seek "=", cboQueryNameList.TEXT

If Not rsQueryNames.NoMatch Then
  SQL$ = Field2Str(rsQueryNames!SQL)
  txtQuery.TEXT = SQL$
  
'code below can later support saving/restoring criteria with save query
'  table$ = (Mid(sql$, 8, InStr(8, sql$, ".") - 8)) 'parse table name from sql$
'  Call Set_Combo(cboTables, table$) 'reload saved sql's selected table
'  txtQuery.Text = sql$
'  SelectClause$ = Mid(sql$, 1, InStr(1, sql$, "FROM") - 1)
'
'  'add select fields to lstSelFields
'  While InStr(1, SelectClause$, ".") - Len(table$) > 0
'    tblfld$ = Mid(SelectClause$, InStr(1, SelectClause$, ".") - Len(table$))
'    SelectClause$ = tblfld$
'    tblfld$ = Mid(tblfld$, 1, InStr(1, tblfld$, " ") - 1)
'    lstSelFields.additem tblfld$
'    SelectClause$ = Mid(SelectClause$, Len(tblfld$) + 2)
'  Wend
  
'  Call CriteriaCBOLoad
  
'  'load grdCriteria with criteria stuff from a table with the proper fields
'  With grdCriteria
'    sql$ = "SELECT * FROM <QueryCriteria> WHERE User = '" & CURRENT_USER$ & "' AND QueryName = '" & cboQueryNameList.Text & "'"
'    Set rsCriteria = dbpm.OpenRecordset(sql$, dbOpenSnapshot)
'    i% = 1
'    While Not rsCriteria.EOF
'      .Row = i%
'      .Col = 1
'      .Text = Field2Str(rsCriteria!TableField)
'      .Col = 2
'      .Text = Field2Str(rsCriteria!Operator)
'      .Col = 3
'      .Text = Field2Str(rsCriteria!Value)
'      .Col = 4
'      .Text = IIf(rsCriteria!Sort, 1, 0)
'      .Col = 5
'      .Text = IIf(rsCriteria!Desc, 1, 0)
'      i% = i% + 1
'      rsCriteria.MoveNext
'    Wend
'  End With

  SSTabQuery.TabVisible(0) = False
  SSTabQuery.TabVisible(1) = False
  cmdQuerySave.Enabled = False
  Call DisplayGridClear
  cmdBuild.Enabled = False
  cmdQueryDelete.Enabled = True
End If

rsQueryNames.Close: Set rsQueryNames = Nothing
dbpm.Close: Set dbpm = Nothing
grdDisplay.SetFocus
End Sub

Private Sub cboTables_Click()
Dim dbpm As database
Dim i%, tabl$
Dim FieldName$

If cboTables.TEXT = "" Then Exit Sub

Screen.MousePointer = vbHourglass

lstFields.Clear
lstSelFields.Clear 'hub
lstSelTables.Clear
txtQuery.TEXT = ""
cmdBuild.Enabled = False
cmdShow.Enabled = False
cmdQuerySave.Enabled = False
Call CriteriaGridClear
Call DisplayGridClear

Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
tabl$ = cboTables.TEXT

For i% = 0 To dbpm.TableDefs(tabl$).Fields.count - 1
  lstFields.AddItem dbpm.TableDefs(tabl$).Fields(i%).Name
Next i%

dbpm.Close: Set dbpm = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdbuild_Click()
Dim i%, txt, order$, hmOrder%, where$, hmWhere%
Dim Field, fieldtype%, operator, Value
Dim t, TableName$, HistTableName$, HmTable As Integer, HmHistTable As Integer
Dim FromTables(4) As String, FromHistTables(4) As String
Dim AddTable As Boolean, AddHistTable As Boolean
Dim groupby$, moregroupby$

Screen.MousePointer = vbHourglass
HmTable = 0
HmHistTable = 0

txtQuery.TEXT = "SELECT"

For i% = 0 To lstSelFields.ListCount - 1
  lstSelFields.Col = 0
  txtQuery.TEXT = txtQuery.TEXT & " " & SetVal(lstSelFields.ColList(i%))
  TableName$ = Left$(lstSelFields.ColList(i%), InStr(lstSelFields.ColList(i%), ".") - 1)
  If i% <> lstSelFields.ListCount - 1 Then txtQuery.TEXT = txtQuery.TEXT & ","
Next i%

txtQuery.TEXT = txtQuery.TEXT & " FROM " & cboTables.TEXT
where$ = " WHERE "
order$ = " ORDER BY "
With grdCriteria
  For i% = 1 To .MaxRows
    .GetText 1, i%, Field
    If Field = "" Then Exit For
    .GetText 2, i%, operator  'build WHERE clausing
    If Trim(operator) <> "" Then
      .GetText 3, i%, Value
      If Value = "" Then
        MsgBox "Value required on Criteria Tab for each criteria.", vbInformation + vbOKOnly, "Build Query"
        GoTo exitcode
      End If
      fieldtype% = GetFieldType(CStr(Field))
      If fieldtype% = dbText Or fieldtype% = dbMemo Then Value = "'" & Value & "'" 'text, memo, boolean delimiters
      If fieldtype% = dbDate Then Value = "#" & Value & "#" 'date delimiters
      hmWhere% = hmWhere% + 1
      If hmWhere% = 1 Then
        where$ = where$ & Field & " " & operator & " " & Value
      Else
        If optAll Then
          where$ = where$ & " AND " & Field & " " & operator & " " & Value
        Else
          where$ = where$ & " OR " & Field & " " & operator & " " & Value
        End If
      End If
    End If
    .GetText 4, i%, txt 'get Sort, build ORDER BY clausing
    If Val(txt) = 1 Then
      hmOrder% = hmOrder% + 1
      txt = Field
      If hmOrder% = 1 Then
        order$ = order$ & Trim(txt)
      Else
        order$ = order$ & "," & Trim(txt)
      End If
      .GetText 5, i%, txt
      If Val(txt) = 1 Then order$ = order$ & " DESC"
    End If
  Next i%
End With
If hmWhere% > 0 Then txtQuery.TEXT = txtQuery.TEXT & where$
If hmOrder% > 0 Then txtQuery.TEXT = txtQuery.TEXT & order$
Call KeywordMargin
cmdShow.Enabled = True
cmdBuild.Enabled = False

exitcode:
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQueryDelete_Click()
If cboQueryNameList.TEXT <> "" Then Call QueryDelete
End Sub

Private Sub CmdDone_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
Dim outfilename$, ret%

On Local Error GoTo extpoint
Dim savdir$, SAVDRIVE$

savdir$ = CurDir
If Mid$(savdir$, 2, 1) = ":" Then
 SAVDRIVE$ = Left$(savdir$, 2)
End If

mdimain!CDIPRINTSET.CancelError = True
mdimain!CDIPRINTSET.Flags = cdlOFNOverwritePrompt
mdimain!CDIPRINTSET.ShowSave

outfilename$ = mdimain!CDIPRINTSET.Filename

ret% = grdDisplay.SaveTabFile(outfilename$)

extpoint:
If Mid$(savdir$, 2, 1) = ":" Then
  ChDrive (SAVDRIVE$)
End If
ChDir (savdir$)
End Sub

Private Sub CmdNew_Click()
Call FormReset
End Sub

Private Sub cmdprint_Click()
Dim font1 As String, font2, Buf As String, qry$
Dim flagQRYend As Boolean, prtqry%
Dim CrLfAt%, QueryName$

prtqry% = 150 'max printer line length
'print grid commands
' create a font as Arial, size 10, bold, no italics,
' underline, no strikethrough and save it as font #1
font1 = "/fn""Arial"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs1"
' Create a font as Times, size 20, no bold, italics, no
' underline, strikethrough and save it as font #2
'font2 = "/fn""Arial"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
' Recall font configurations and set the header and footer text
'buf = font1 + font2 + "/f1This is font #1 " + Chr$(13) + "/f2This is font #2"
't% = Val(Text1.Text)

qry$ = Trim(txtQuery.TEXT)
Buf = font1 + "/n/l"
  flagQRYend = False
  
q:
  CrLfAt% = InStr(1, qry$, vbCrLf) 'replace vbCrLf with /n for print
  If CrLfAt% <> 0 And CrLfAt% < prtqry% Then
    Buf = Buf & Mid(qry$, 1, CrLfAt% - 1) + "/n/l"
    qry$ = Mid$(qry$, CrLfAt% + 2)
    GoTo q
  End If
  If Len(qry$) > prtqry% Then
    Buf = Buf + Left$(qry$, prtqry%) + "/n/l"
    qry$ = Mid$(qry$, prtqry% + 1)
  Else
    Buf = Buf + qry$
    flagQRYend = True
  End If
  If Not flagQRYend Then GoTo q
  Buf = Buf + "/n/n"
  QueryName$ = cboQueryNameList.TEXT
  grdDisplay.PrintHeader = "This was your query: " & QueryName$ & Buf & "It produced the following results"
  Buf = font1 + "/n/c/p"
  grdDisplay.PrintFooter = Buf
  grdDisplay.PrintBorder = False
  grdDisplay.PrintColHeaders = True
  grdDisplay.PrintRowHeaders = False
  grdDisplay.Action = 32
End Sub

Private Sub cmdquerysave_Click()
Dim Name$

If txtQuery.TEXT = "" Then
  MsgBox "Built Query required.", vbOKOnly + vbInformation, Me.Caption
  Exit Sub
End If
Load frminput_box
With frminput_box
  .txtinput.TEXT = cboQueryNameList.TEXT '"" if cboQueryNameList.Text is empty or selected query name
  .txtinput.SelStart = 0
  .txtinput.SelLength = Len(.txtinput.TEXT)
  .Caption = "Query Save"
  !lblinput.Caption = "Query Name"
  !txtinput.MaxLength = 80
  .Show vbModal
  If .Tag = "Cancel" Then Exit Sub 'user cancelled or ""
  Name$ = .Tag
End With

Call QuerySave(Name$, "R", txtQuery.TEXT)
mode$ = "QuerySave" 'reset to "" in cboQueryNameList_Click code
cboQueryNameList.AddItem Name$
cboQueryNameList.ListIndex = cboQueryNameList.newindex
cmdQuerySave.Enabled = False
End Sub

Private Sub Cmdshow_Click()
Dim i%, AdjdWidth%

If txtQuery.TEXT = "" Then Call cmdbuild_Click
On Local Error GoTo ehandle
Screen.MousePointer = vbHourglass
grdDisplay.MaxRows = 0

Data1.DatabaseName = dbfile$
Data1.RecordSource = txtQuery.TEXT
Data1.Refresh

For i% = 3 To grdDisplay.MaxCols
  AdjdWidth% = grdDisplay.MaxTextColWidth(i%)
  If AdjdWidth% > grdDisplay.ColWidth(i%) Then grdDisplay.ColWidth(i%) = AdjdWidth%
  If AdjdWidth% > 4000 Then grdDisplay.ColWidth(i%) = 4000
Next i%

If Not Data1.Recordset.BOF And Not Data1.Recordset.EOF Then
  Data1.Recordset.MoveLast
  grdDisplay.MaxRows = Data1.Recordset.RecordCount
  Data1.Recordset.MoveFirst
End If
cmdShow.Enabled = False
grdDisplay.SetFocus

exitcode:
  SSTabQuery.Tab = 2
  Screen.MousePointer = vbDefault
  Exit Sub

ehandle:
  TEXT$(1) = "DataBase: " & dbfile$
  TEXT$(2) = ""
  TEXT$(3) = ""
  TEXT$(4) = ""
  TEXT$(5) = ""
  GMODNAME$ = Me.Name & " cmdShow_Click"
  GERRNUM$ = Str$(Err.Number)
  GERRSOURCE$ = Err.Source
  Call POP_ERROR(TEXT$())
  GoTo exitcode
End Sub

Private Sub FormInit()
Dim ret

cboTables.ListIndex = -1 'hub
cmdBuild.Enabled = False
cmdShow.Enabled = False
cmdQueryDelete.Enabled = False
SSTabQuery.Tab = 0
'Initialize the criteria grid headings and cell types
With grdCriteria
  .SetText 1, 0, "Field Name"
  .SetText 2, 0, "Operator"
  .SetText 3, 0, "Value"
  .SetText 4, 0, "Sort"
  .SetText 5, 0, "Desc Order"
  .Row = -1
  .ColWidth(0) = 300
  .ColWidth(1) = 2500
  .ColWidth(2) = 1400
  .ColWidth(3) = 1200
  .ColWidth(4) = 500
  .ColWidth(5) = 1000
  .Col = 1
  .CellType = 8 'SS_CELL_TYPE_COMBOBOX

  .Col = 1
  .Row = -1
  .Col = 2
  .CellType = 8 'SS_CELL_TYPE_COMBOBOX
  .Col = 3
  .CellType = 1 'SS_CELL_TYPE_EDIT
  .TypeHAlign = 0 'SS_CELL_H_ALIGN_LEFT
  .TypeEditMultiLine = False
  .TypeEditLen = 50
  .Col = 4
  .TypeCheckCenter = True
  .CellType = 10 'SS_CELL_TYPE_CHECKBOX
  .Col = 5
  .TypeCheckCenter = True
  .CellType = 10 'SS_CELL_TYPE_CHECKBOX
End With
Call OperatorCBOLoad
End Sub

Private Sub OperatorCBOLoad()
Dim CBO$, nxt$(12), i%
nxt$(0) = "="
nxt$(1) = ">"
nxt$(2) = "<"
nxt$(3) = ">="
nxt$(4) = "<="
nxt$(5) = "<>"
nxt$(6) = "LIKE"

For i% = 0 To 6
      CBO$ = CBO$ & nxt$(i%) & Chr$(9)
Next i%
With grdCriteria
  .Col = 2: .col2 = 2
  .Row = 1: .Row2 = 500
  .BlockMode = True
  .TypeComboBoxList = CBO$
  .BlockMode = False
End With
End Sub

Private Sub Form_Activate()
If Me.Tag = "" Then Exit Sub
FrmHubQuery.AutoRedraw = True
DoEvents
'Call CreateTableAttachment(dbfile$, RepDir$, "RptFields" & current_USER$, "RptFields")
Call CriteriaCBOLoad
Me.Tag = ""
cboTables.SetFocus
End Sub

Private Sub Form_Load()
Dim dbpm As database
Dim rsQueryNames As Recordset
Dim i%

Call centermdiform(Me, mdimain, 0, 0)
Call FormInit
SSTabQuery.Tab = 0
Frame1.Refresh
SSTabQuery.Tab = 0
Data1.ReadOnly = False
Data1.DatabaseName = dbfile$
Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)

For i% = 0 To dbpm.TableDefs.count - 1
  If Left$(dbpm.TableDefs(i%).Name, 4) <> "MSys" Then
     cboTables.AddItem dbpm.TableDefs(i%).Name
  End If
Next i%

Set rsQueryNames = dbpm.OpenRecordset("Query", dbOpenSnapshot)
While Not rsQueryNames.EOF
  If rsQueryNames!QueryType = "R" Then cboQueryNameList.AddItem rsQueryNames!Name
  rsQueryNames.MoveNext
Wend

cmdQuerySave.Enabled = False
rsQueryNames.Close: Set rsQueryNames = Nothing
dbpm.Close: Set dbpm = Nothing

With grdCriteria
  .SelectBlockOptions = SS_SELBLOCKOPT_ROWS 'allow row(s) to be selected for deletes
  .Row = -1
  .UserResizeRow = SS_USER_RESIZE_OFF 'disable user row height resizing
  .Col = 0
  .UserResizeCol = SS_USER_RESIZE_OFF 'disable user col 0 width resizing
End With

With grdDisplay
  .RowHeight(0) = 440
  .UserResize = SS_USER_RESIZE_COL
  .SelectBlockOptions = SS_SELBLOCKOPT_ROWS 'allow user to select rows for highlighting
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call DeleteTableAttachment(dbfile$, "rptfields " & current_USER$)
Set FrmHubQuery = Nothing
End Sub

Private Sub grdCriteria_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
cmdBuild.Enabled = True
txtQuery.TEXT = ""
End Sub

Private Sub grdCriteria_Change(ByVal Col As Long, ByVal Row As Long)
fraCriteria.Refresh
cmdBuild.Enabled = True
txtQuery.TEXT = ""
End Sub

Private Sub grdCriteria_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim SQL$
Dim txt, FieldName$, operator$
Dim my_tabledef As TableDef
Dim my_field As Field, temp$, fname$
Dim TableName$

If Row = 0 Then Exit Sub
If Col <> 3 Then Exit Sub
grdCriteria.GetText 1, Row, txt
If Trim(txt) <> "" Then
  FieldName$ = txt
  TableName$ = Left$(FieldName$, InStr(FieldName$, ".") - 1)
  grdCriteria.GetText 2, Row, txt
  operator$ = txt
  SQL$ = "SELECT DISTINCT " & FieldName$ & " as dv FROM " & cboTables.TEXT
  txtDistinctSql.TEXT = SQL$
  frmListDistinct.Show vbModal

  If frmListDistinct.Tag = "Cancel" Then
    Unload frmListDistinct
    Exit Sub
  End If
  temp$ = frmListDistinct.Tag
  Unload frmListDistinct
  With grdCriteria
    .Col = 3
    .Row = Row
    If .TEXT <> temp$ Then
      txtQuery.TEXT = ""
      cmdBuild.Enabled = True
    End If
    .TEXT = temp$
    .Col = 0
    .Row = 0
    .Action = SS_ACTION_ACTIVE_CELL
    .Action = SS_ACTION_GOTO_CELL
    .Action = SS_POSITION_UPPER_LEFT
  End With
End If
End Sub

Private Sub grdCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  Call ClearGridSelected(Me!grdCriteria)
End If
If grdCriteria.Col = 3 Then cmdBuild.Enabled = True
End Sub

Private Sub grdDisplay_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal datatype As Integer)
Dim i%
Dim CalcColWidth%

With grdDisplay
  .Row = -1
  .Col = Col
  Select Case datatype
    Case dbText
      .CellType = SS_CELL_TYPE_STATIC_TEXT
    Case dbLong
      grdDisplay.CellType = SS_CELL_TYPE_INTEGER
    Case dbInteger
      .CellType = SS_CELL_TYPE_INTEGER
    Case dbDouble
      .CellType = SS_CELL_TYPE_FLOAT
    Case dbDate
      .CellType = SS_CELL_TYPE_DATE
    Case dbBoolean
      .CellType = SS_CELL_TYPE_CHECKBOX
      .TypeCheckCenter = True
    Case dbMemo
      .CellType = SS_CELL_TYPE_STATIC_TEXT
  End Select
  .DAutoSizeCols = 0
  CalcColWidth% = 138 * Len(DataField)
  If CalcColWidth% > .ColWidth(Col) Then .ColWidth(Col) = CalcColWidth%
  If .ColWidth(Col) > 1600 Then
    .ColWidth(Col) = grdDisplay.ColWidth(Col) / 2.4
  End If
End With
End Sub

Private Sub lstFIELDS_DblClick()
Dim i%, Found As Boolean, CBO$

Found = False
lstSelTables.Col = 0
lstSelFields.Col = 0
lstFields.Col = 0

For i% = 0 To lstSelTables.ListCount - 1
  If lstSelTables.ColList(i%) = cboTables.TEXT Then
    Found = True
  End If
Next i%
If Not Found Then
  lstSelTables.AddItem cboTables.TEXT
End If

Found = False
CBO$ = ""
For i% = 0 To lstSelFields.ListCount - 1
  If lstSelFields.ColList(i%) = cboTables.TEXT & "." & lstFields.ColText Then
    Found = True
  End If
Next i%

If Not Found Then
  lstSelFields.AddItem cboTables.TEXT & "." & lstFields.ColText
  lstSelFields.ListIndex = lstSelFields.NewItem
  txtQuery.TEXT = ""
End If
Call CriteriaCBOLoad 'reload grdCriteria col 1 (fieldname) cbo
cmdBuild.Enabled = True
End Sub

Private Sub lstSelFields_DblClick()
Dim fldName$, i%, txt

If lstSelFields.ListIndex >= 0 Then GoSub DelItem
Exit Sub

DelItem:
  With grdCriteria
    'delete field from criteria if in list
    For i% = 1 To .MaxRows
      .GetText 1, i%, txt
      If txt = "" Then Exit For
      If txt = lstSelFields.TEXT Then
        .Row = i%
        .Action = SS_ACTION_DELETE_ROW
      End If
    Next i%
  End With
  
  'delete selected field from lstSelFields listbox
  lstSelFields.ListIndex = lstSelFields.ListIndex - 1
  lstSelFields.RemoveItem lstSelFields.ListIndex + 1
  txtQuery.TEXT = ""
  Call CriteriaCBOLoad
  cmdShow.Enabled = False
  If lstSelFields.ListCount = 0 Then
    cmdBuild.Enabled = False
  Else
    cmdBuild.Enabled = True
  End If
  Return
End Sub

Private Sub optAll_Click()
txtQuery.TEXT = ""
End Sub

Private Sub optAny_Click()
txtQuery.TEXT = ""
End Sub

Private Sub txtQUERY_Change()
cmdQuerySave.Enabled = True
If txtQuery.TEXT <> "" Then
  cmdBuild.Enabled = True
  cmdShow.Enabled = True
End If
End Sub



