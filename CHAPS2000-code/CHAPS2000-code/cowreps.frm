VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form cowreps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cow Reports"
   ClientHeight    =   4470
   ClientLeft      =   3060
   ClientTop       =   2730
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   4680
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   255
      Width           =   3675
      _Version        =   65536
      _ExtentX        =   6482
      _ExtentY        =   2778
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
   Begin VB.ComboBox cboyear 
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   2655
      Width           =   1275
   End
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Herd"
      Height          =   345
      Left            =   2760
      TabIndex        =   33
      Top             =   1875
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdCowSelect 
      Caption         =   "S&elect"
      Height          =   345
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   855
      Left            =   3360
      TabIndex        =   3
      Top             =   3030
      Width           =   1215
      Begin VB.OptionButton optprint 
         Caption         =   "Print"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optpreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3675
      TabIndex        =   2
      Top             =   4065
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   2235
      TabIndex        =   0
      Top             =   4065
      Width           =   1000
   End
   Begin VB.Frame fraCowList 
      Caption         =   "Order By"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3030
      Width           =   3135
      Begin VB.CheckBox ChkDT 
         Caption         =   "Detail"
         Height          =   240
         Left            =   2235
         TabIndex        =   14
         Top             =   390
         Width           =   870
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Cow ID"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Cow Age"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   495
         Width           =   1230
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "MPPA"
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Sire ID"
         Height          =   255
         Index           =   3
         Left            =   1380
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame FraPerf 
      Height          =   1035
      Left            =   120
      TabIndex        =   25
      Top             =   3030
      Width           =   3180
      Begin VB.Frame Frame3 
         Caption         =   "Include"
         Height          =   825
         Left            =   1935
         TabIndex        =   29
         Top             =   135
         Visible         =   0   'False
         Width           =   1155
         Begin VB.OptionButton OptPerf 
            Caption         =   "Pedigree"
            Height          =   210
            Index           =   4
            Left            =   45
            TabIndex        =   32
            Top             =   585
            Width           =   975
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Active"
            Height          =   210
            Index           =   3
            Left            =   45
            TabIndex        =   31
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Cull"
            Height          =   225
            Index           =   2
            Left            =   45
            TabIndex        =   30
            Top             =   375
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sort By"
         Height          =   825
         Left            =   90
         TabIndex        =   26
         Top             =   120
         Width           =   1140
         Begin VB.OptionButton OptPerf 
            Caption         =   "Sire ID"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   510
            Width           =   1050
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Cow ID"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   225
            Width           =   975
         End
      End
   End
   Begin VB.Frame FraRef 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3135
      Width           =   3135
      Begin VB.OptionButton OptInclude 
         Caption         =   "Pedigree"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   18
         Top             =   195
         Width           =   1140
      End
      Begin VB.OptionButton OptInclude 
         Caption         =   "Culled"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   17
         Top             =   540
         Width           =   1140
      End
      Begin VB.OptionButton OptInclude 
         Caption         =   "Active"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   1140
      End
   End
   Begin VB.Frame FraBrd 
      Caption         =   "Order By"
      Height          =   960
      Left            =   120
      TabIndex        =   19
      Top             =   3015
      Width           =   3135
      Begin VB.OptionButton optOrder 
         Caption         =   "Breeding Cond Score"
         Height          =   390
         Index           =   8
         Left            =   1380
         TabIndex        =   24
         Top             =   450
         Width           =   1410
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Est Calf Interval"
         Height          =   255
         Index           =   7
         Left            =   1380
         TabIndex        =   23
         Top             =   195
         Width           =   1410
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Cow Age"
         Height          =   255
         Index           =   5
         Left            =   105
         TabIndex        =   21
         Top             =   435
         Width           =   1230
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Cow ID"
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   1095
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Est Fetal Age"
         Height          =   255
         Index           =   6
         Left            =   105
         TabIndex        =   22
         Top             =   660
         Width           =   1305
      End
   End
   Begin VB.Label LBLBTOD 
      Alignment       =   1  'Right Justify
      Caption         =   "Bull Turn Out Date"
      Height          =   225
      Left            =   600
      TabIndex        =   35
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Label lblCows 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many Cows"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2340
      Width           =   1545
   End
   Begin VB.Label lblMultiCows 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All"
      Height          =   255
      Left            =   1755
      TabIndex        =   6
      Top             =   2340
      Width           =   900
   End
End
Attribute VB_Name = "cowreps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 13

Dim t As Integer
Dim reports(hmreps%) As String
Dim mCowList As New FrmSelect_Multi_Cows

Private Sub Create_Cow_NotesRPT()
Dim where$, orderby$, SQL$, DB As DAO.database
If OptPerf(3).Value Then where = " where cowprof.active = 'A' "
If OptPerf(2).Value Then where = " where cowprof.active = 'C' "
If OptPerf(4).Value Then where = " where cowprof.active = 'P' "
where = where & " and cowprof.herdid = '" & herdid & "' "
If OptPerf(0).Value Then orderby = " order by cowprof.cowid "
If OptPerf(1).Value Then orderby = " order by cowprof.birthdate "
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from cow_notes"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
SQL = "insert into cow_notes in '" & repfile & "' SELECT DISTINCTROW cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.notes, cowprof.active "
If Val(lblMultiCows.Caption) > 0 Then
   Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
   SQL = SQL & " FROM RPTCows INNER JOIN cowprof ON (RPTCows.CowID = cowprof.cowID) AND (RPTCows.HerdID = cowprof.HerdID) "
Else
   SQL = SQL & " From cowprof " & where & orderby
End If
DB.Execute SQL
Call DeleteTableAttachment(dbfile, "RPTCows")
DB.Close: Set DB = Nothing
End Sub

Private Sub Create_Cow_RefRPT()
Dim pDB As DAO.database, pRS As DAO.Recordset, SQL$, where$
Set pDB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
pDB.Execute "delete * from CowRef"
Set pDB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
SQL = "insert into cowref in '" & repfile & "' SELECT DISTINCTROW cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire, cowprof.dam, cowprof.calfid, cowprof.regnum, cowprof.regname, cowprof.elecID, cowprof.enteredherd, cowprof.source, cowprof.active, Max([CowAge]) AS Age, first(cowprof.notes ) as notes "
If Val(cowreps.lblMultiCows) > 0 Then
      Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
      SQL = SQL & " FROM RPTCows INNER JOIN (cowprof LEFT JOIN calfbirth ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) ON (RPTCows.CowID = cowprof.cowID) AND (RPTCows.HerdID = cowprof.HerdID) "
Else
      SQL = SQL & " FROM cowprof LEFT JOIN calfbirth ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID) "
End If
If OptInclude(0).Value Then where = " where cowprof.active = 'A' "
If OptInclude(1).Value Then where = " where cowprof.active = 'C' "
If OptInclude(2).Value Then where = " where cowprof.active = 'P' "
SQL = SQL & where & " and  cowprof.herdid = '" & herdid & "' GROUP BY cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire, cowprof.dam, cowprof.calfid, cowprof.regnum, cowprof.regname, cowprof.elecID, cowprof.enteredherd, cowprof.source, cowprof.active " & sortcows
pDB.Execute SQL, dbFailOnError

Call DeleteTableAttachment(dbfile, "RPTCows")
'Set pDB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
'Set pRS = pDB.OpenRecordset("select * from cowref where age <> null", dbOpenDynaset)
'Do Until pRS.EOF
'   pRS.Edit
'   pRS!dam = FindDam(2, Field2Num(pRS!age))
'   pRS.Update
'   pRS.MoveNext
'Loop
'pRS.Close: Set pRS = Nothing
pDB.Close: Set pDB = Nothing
End Sub

Private Function Create_CowBreed_RPT() As Boolean
Dim SQL$, DB As DAO.database, order$, RS As DAO.Recordset, TurnDate As Date
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from cowbreed"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select currentdate from turnoutdate where herdid = '" & herdid & "'", dbOpenSnapshot)
If cboyear.TEXT <> "" And cboyear.TEXT <> "Not Set" Then TurnDate = Left(cboyear.TEXT, 10) Else GoTo Warning
If RS.EOF Then
Warning:
   'TurnDate = Field2Date(RS!thetext)
   MsgBox "Please Set Bull Turn Out Date For Herd " & herdid, vbOKOnly
   RS.Close: Set RS = Nothing
   DB.Close: Set DB = Nothing
   Create_CowBreed_RPT = False
   Exit Function
End If
SQL = "insert into cowbreed in '" & repfile & "' SELECT DISTINCTROW cowbrd.CowID, cowbrd.breeddate1, cowbrd.breeddate2, cowbrd.breeddate3, cowbrd.breedbull1, cowbrd.breedbull2, cowbrd.breedbull3, cowbrd.stat, [cowbrd].[conceptdate]-[cowbrd].[age] AS AstDate "
If Val(lblMultiCows.Caption) > 0 Then
   Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
   SQL = SQL & " FROM (cowprof INNER JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID)) INNER JOIN RPTCows ON (cowprof.cowID = RPTCows.CowID) AND (cowprof.HerdID = RPTCows.HerdID) "
Else
   SQL = SQL & " FROM cowprof INNER JOIN cowbrd ON cowprof.cowID = cowbrd.CowID AND cowprof.HerdID = cowbrd.HerdID "
End If
SQL = SQL & " where cowprof.active = 'A' and cowbrd.calfdate = #" & TurnDate & "# and cowprof.herdid = '" & herdid & "'"
If optOrder(0).Value Then order = " order by cowbrd.cowid"
If optOrder(1).Value Then order = " order by cowbrd.stat"
SQL = SQL & order
DB.Execute SQL
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Create_CowBreed_RPT = True
report.Setformulas("ExpDate") = "'Exposed Date: " & TurnDate & "'"
Call DeleteTableAttachment(dbfile, "RPTCows")
End Function

Private Sub create_Culled_refLIST_report()
  Dim SQL$
  Dim DB As database
  Dim DBREP As database
  Dim Index%
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  
  DB.Execute "create table rptCull (CullCode TEXT)"
  
  frmPopUpCull.LstCullCodes.Col = 0
  Do Until Index = frmPopUpCull.LstCullCodes.ListCount
      If frmPopUpCull.LstCullCodes.Tagged(Index) Then DB.Execute "insert into rptcull (cullcode) values ('" & frmPopUpCull.LstCullCodes.ColList(Index) & "')"
     Index = Index + 1
  Loop
  
  DBREP.Execute ("delete * from culldref")
  DBREP.Close
  SQL$ = "insert into culldref in '" & repfile$ & "' SELECT cowPROF.HerdID, cowPROF.cowID, cowPROF.source, cowPROF.dateculled, cowPROF.reasonculled, cowPROF.commentsculled  "
  If Val(lblMultiCows.Caption) > 0 Then
    SQL = SQL & " FROM cowPROF INNER JOIN RPTCows ON (cowPROF.cowID = RPTCows.CowID) AND (cowPROF.HerdID = RPTCows.HerdID) "
  Else
    SQL = SQL & " FROM cowPROF INNER JOIN rptCull ON cowPROF.reasonculled = rptCull.CullCode "
  End If
  SQL = SQL & " WHERE ACTIVE = 'C' and cowprof.herdid = '" & herdid & "'"
  If IsDate(frmPopUpCull.txtStartDate.TEXT) Then SQL$ = SQL$ & " and cowprof.dateculled >= #" & frmPopUpCull.txtStartDate.TEXT & "#"
  If IsDate(frmPopUpCull.txtEndDate.TEXT) Then SQL$ = SQL$ & " and cowprof.dateculled <= #" & frmPopUpCull.txtEndDate.TEXT & "#"
  Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
  DB.Execute (SQL$)
  Call DeleteTableAttachment(dbfile, "RPTCows")
  DB.Execute "drop table rptCull"
  DB.Close
End Sub

Private Sub Create_Pedigree_RPT()
Dim DB As DAO.database, SQL$, where$, orderby$
On Local Error GoTo ErrHandler
Dim repdb As database
Dim tbData As Recordset
Dim tbavg As Recordset
Dim theavg As Double
Dim theagvcnt As Integer


100 Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
110 DB.Execute "delete * from pedigree"
120 DB.Execute "delete * from pedigree_avgs"
130 Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
    Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
140 Call CreateTableAttachment(dbfile$, repfile$, "Pedigree", "Pedigree")
Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
Screen.MousePointer = vbHourglass
GoSub Build_Cow_Data
GoSub Update_2nd_Gen
GoSub Update_3rd_Gen
GoSub Build_Progeny_Avgs
GoSub Build_Sort_Order
TEXT(1) = ""
Call DeleteTableAttachment(dbfile, "Pedigree")
Call DeleteTableAttachment(dbfile, "RPTCows")
150 DB.Close: Set DB = Nothing
    repdb.Close: Set repdb = Nothing
Exit Sub

ErrHandler:
TEXT(2) = Erl
GMODNAME$ = Me.Name & " - Create_Pedigree_RPT"
Resume
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())

Build_Cow_Data:
   TEXT(1) = "Build_Cow_Data"
200    SQL = "insert into pedigree in '" & repfile & "' "
210    'SQL = SQL & " SELECT DISTINCTROW cowprof.cowID AS CowID_1, cowprof.mpda AS MPPA, cowprof.sire AS Sire_ID_2, cowprof.dam AS Cow_ID_2, cowprof.calfid, calfbirth.birthdate AS Birth_Date, calfbirth.birthwt AS Birth_Weight, calfwean.wt205 AS Adj_205, calfwean.ratio AS Adj_205_R, calfcarcass.score AS Frame_Score, calfrep.w365 AS Adj_365, 0 AS Adj_365_R, calfcarcass.ygrade AS Yield_Grade, calfcarcass.qgrade AS Quality_Grade, calfcarcass.ywt AS HCW, cowepd.epdbirthwt, cowepd.epdweanwt, cowepd.epdyearwt, cowepd.epdmatww, cowepd.epdmatmilk, cowepd.accbirthwt, cowepd.accweanwt, cowepd.accyearwt, cowepd.accmatww, cowepd.accmatmilk, cowepd.misc1, cowepd.misc2, cowepd.misc3, cowepd.misc4, cowepd.misc5, cowepd.misc6, cowepd.misc8, cowepd.misc9, cowepd.misc10, cowepd.acc1, cowepd.acc2, cowepd.acc3, cowepd.acc4, cowepd.acc5, cowepd.acc6, cowepd.acc7, cowepd.acc8, cowepd.acc9, cowepd.acc10 "
220    SQL = SQL & " SELECT DISTINCTROW cowprof.cowID AS CowID_1, cowprof.mpda AS MPPA, cowprof.sire AS Sire_ID_2, cowprof.dam AS Cow_ID_2, cowprof.calfid, calfbirth.birthdate AS Birth_Date, calfbirth.birthwt AS Birth_Weight, calfwean.wt205 AS Adj_205, calfwean.ratio AS Adj_205_R, calfcarcass.score AS Frame_Score, calfrep.w365 AS Adj_365, 0 AS Adj_365_R, calfcarcass.ygrade AS Yield_Grade, calfcarcass.qgrade AS Quality_Grade, calfcarcass.ywt AS HCW, cowepd.epdbirthwt, cowepd.epdweanwt, cowepd.epdyearwt, cowepd.epdmatww, cowepd.epdmatmilk, cowepd.accbirthwt, cowepd.accweanwt, cowepd.accyearwt, cowepd.accmatww, cowepd.accmatmilk, cowepd.misc1, cowepd.misc2, cowepd.misc3, cowepd.misc4, cowepd.misc5, cowepd.misc6, cowepd.misc8, cowepd.misc9, cowepd.misc10, cowepd.acc1, cowepd.acc2, cowepd.acc3, cowepd.acc4, cowepd.acc5, cowepd.acc6, cowepd.acc7, cowepd.acc8, cowepd.acc9, cowepd.acc10, cowprof.HerdID "
         If Val(lblMultiCows.Caption) > 0 Then
            SQL = SQL & " FROM RPTCows INNER JOIN (((((cowprof LEFT JOIN cowepd ON (cowprof.cowID = cowepd.cowID) AND (cowprof.HerdID = cowepd.HerdID)) LEFT JOIN calfbirth ON (cowprof.calfid = calfbirth.CalfID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) ON (RPTCows.CowID = cowprof.cowID) AND (RPTCows.HerdID = cowprof.HerdID)"
         Else
            SQL = SQL & " FROM ((((cowprof LEFT JOIN cowepd ON (cowprof.cowID = cowepd.cowID) AND (cowprof.HerdID = cowepd.HerdID)) LEFT JOIN calfbirth ON (cowprof.calfid = calfbirth.CalfID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)  "
         End If
         SQL = SQL & " where cowprof.herdid = '" & herdid & "' "
         DB.Execute SQL, dbFailOnError
Return

Build_Progeny_Avgs:
   TEXT(1) = "Build_Progeny_Avgs"
   'progeny avgs by cow id
'300    SQL = "insert into pedigree_avgs in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, Year(Max([calfbirth].[birthdate]))-Year(Min([calfbirth].[birthdate])) AS Years_Service, Count(calfbirth.CalfID) AS Progeny, Sum(iif(calfbirth.birthwt > 0, calfbirth.birthwt, 0)) / Sum(iif(calfbirth.birthwt > 0, 1, 0)) AS Avg_BW, Sum(iif(calfwean_1.wt205 > 0, calfwean_1.wt205, 0)) / Sum(iif(calfwean_1.wt205 > 0, 1, 0)) AS Avg_Adj_205, Sum(iif(calfwean_1.ratio > 0 ,calfwean_1.ratio, 0)) / Sum(iif(calfwean_1.ratio > 0, 1, 0)) AS Avg_Adj_205_R, Sum(iif(calfwean_1.score > 0, calfwean_1.score, 0)) / Sum(iif(calfwean_1.score > 0, 1, 0)) AS Avg_FS, Sum(iif(calfrep.w365 > 0, calfrep.w365, 0)) / Sum(iif(calfrep.w365 > 0, 1, 0)) AS Avg_Adj_365, Sum(iif(calfcarcass_1.ywt > 0, calfcarcass_1.ywt, 0)) / Sum(iif(calfcarcass_1.ywt > 0, 1, 0)) AS Avg_HCW, Sum(iif(calfcarcass_1.ygrade > 0, calfcarcass_1.ygrade, 0)) / Sum(iif(calfcarcass_1.ygrade>0, 1, 0))  AS Avg_YGrade "
'SQL = "insert into pedigree_avgs in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, Year(Max([calfbirth].[birthdate]))-Year(Min([calfbirth].[birthdate])) + 1 AS Years_Service, Count(calfbirth.CalfID) AS Progeny, Sum(iif(calfbirth.birthwt > 0, calfbirth.birthwt, 0)) / Sum(iif(calfbirth.birthwt > 0, 1, 0)) AS Avg_BW,"
'SQL = SQL & " Sum(iif(calfwean_1.wt205 > 0, switch(calfbirth.sex = '0', calfwean_1.wt205 * 1.0, calfbirth.sex = '1', calfwean_1.wt205 * .95, calfbirth.sex = '2', calfwean_1.wt205 * 1.05, calfbirth.sex = '3', calfwean_1.wt205), 0)) / Sum(iif(calfwean_1.wt205 > 0, 1, 0)) AS Avg_Adj_205, Sum(iif(calfwean_1.ratio > 0 ,calfwean_1.ratio, 0)) / Sum(iif(calfwean_1.ratio > 0, 1, 0)) AS Avg_Adj_205_R, Sum(iif(calfwean_1.score > 0, calfwean_1.score, 0)) / Sum(iif(calfwean_1.score > 0, 1, 0)) AS Avg_FS, Sum(iif(calfrep.w365 > 0, calfrep.w365, 0)) / Sum(iif(calfrep.w365 > 0, 1, 0)) AS Avg_Adj_365, Sum(iif(calfcarcass_1.ywt > 0, calfcarcass_1.ywt, 0)) / Sum(iif(calfcarcass_1.ywt > 0, 1, 0)) AS Avg_HCW, Sum(iif(calfcarcass_1.ygrade > 0, calfcarcass_1.ygrade, 0)) / Sum(iif(calfcarcass_1.ygrade>0, 1, 0))  AS Avg_YGrade"
SQL = "insert into pedigree_avgs in '" & repfile & "' SELECT DISTINCTROW calfbirth.CowID, Year(Max([calfbirth].[birthdate]))-Year(Min([calfbirth].[birthdate]))+1 AS Years_Service, Count(calfbirth.CalfID) AS Progeny, Sum(IIf(calfbirth.birthwt>0,calfbirth.birthwt,0))/Sum(IIf(calfbirth.birthwt>0,1,0)) AS Avg_BW, Sum(IIf(calfwean_1.wt205>0,Switch(calfbirth.sex='0',calfwean_1.wt205*1,calfbirth.sex='1',calfwean_1.wt205*0.95,calfbirth.sex='2',calfwean_1.wt205*1.05,calfbirth.sex='3',calfwean_1.wt205),0))/Sum(IIf(calfwean_1.wt205>0,1,0)) AS Avg_Adj_205, Sum(IIf(calfwean_1.ratio>0,calfwean_1.ratio,0))/Sum(IIf(calfwean_1.ratio>0,1,0)) AS Avg_Adj_205_R, Sum(IIf(calfwean_1.score>0,calfwean_1.score,0))/Sum(IIf(calfwean_1.score>0,1,0)) AS Avg_FS, Sum(IIf(calfrep.w365>0,calfrep.w365,0))/Sum(IIf(calfrep.w365>0,1,0)) AS Avg_Adj_365, Sum(IIf(calfcarcass_1.ywt>0,calfcarcass_1.ywt,0))/Sum(IIf(calfcarcass_1.ywt>0,1,0)) AS Avg_HCW,"
'SQL = SQL & " Sum(IIf(calfcarcass_1.ygrade>0,calfcarcass_1.ygrade,0))/Sum(IIf(calfcarcass_1.ygrade>0,1,0)) AS Avg_YGrade, Sum(Switch(calfcarcass.qgrade='Prime+',10.5,calfcarcass.qgrade='Prime ',9.5,calfcarcass.qgrade='Prime-',8.5,calfcarcass.qgrade='Choice+',7,calfcarcass.qgrade='Choice',6.5,calfcarcass.qgrade='Choice-',5.5,calfcarcass.qgrade='Select+',4.75,calfcarcass.qgrade='Select-',4.25,calfcarcass.qgrade='Standard+',3.5,calfcarcass.qgrade='Standard-',2.5,IsNull(calfcarcass.qgrade),0)) / Sum(iif(isnull(calfcarcass.qgrade) = false, 1, 0)) AS Avg_QScore, calfbirth.herdid "
SQL = SQL & " Sum(IIf(calfcarcass_1.ygrade>0,calfcarcass_1.ygrade,0))/Sum(IIf(calfcarcass_1.ygrade>0,1,0)) AS Avg_YGrade, 0 AS Avg_QScore, calfbirth.herdid "
310   If Val(lblMultiCows.Caption) > 0 Then
            SQL = SQL & " FROM ((cowprof INNER JOIN (((calfbirth INNER JOIN calfcarcass AS calfcarcass_1 ON (calfbirth.CalfID = calfcarcass_1.CalfID) AND (calfbirth.HerdID = calfcarcass_1.HerdID)) INNER JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) INNER JOIN calfwean AS calfwean_1 ON (calfbirth.CalfID = calfwean_1.CalfID) AND (calfbirth.HerdID = calfwean_1.HerdID)) ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) INNER JOIN RPTCows ON (cowprof.cowID = RPTCows.CowID) AND (cowprof.HerdID = RPTCows.HerdID) "
         Else
            SQL = SQL & " FROM (cowprof INNER JOIN (((calfbirth INNER JOIN calfcarcass AS calfcarcass_1 ON (calfbirth.CalfID = calfcarcass_1.CalfID) AND (calfbirth.HerdID = calfcarcass_1.HerdID)) INNER JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) INNER JOIN calfwean AS calfwean_1 ON (calfbirth.CalfID = calfwean_1.CalfID) AND (calfbirth.HerdID = calfwean_1.HerdID)) ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID) "
         End If
320      SQL = SQL & where & " GROUP BY calfbirth.CowID, calfbirth.herdid"
330      DB.Execute SQL
 
 
' " build AVG_Qscore "
 
 Set tbavg = repdb.OpenRecordset("pedigree_avgs", dbOpenTable)
 While Not tbavg.EOF
   SQL = "SELECT DISTINCTROW calfbirth.CowID, calfbirth.HerdID, calfbirth.CalfID, calfcarcass.qgrade FROM calfbirth LEFT JOIN calfcarcass ON calfbirth.CalfID = calfcarcass.CalfID AND calfbirth.HerdID = calfcarcass.HerdID    WHERE (((calfbirth.cowID)='" & tbavg!CowID & "') AND ((calfbirth.HerdID)='" & tbavg!herdid & "'))"
   Set tbData = DB.OpenRecordset(SQL, dbOpenDynaset)
   theavg = 0
   theagvcnt = 0

   While Not tbData.EOF
    Select Case tbData!qgrade
    Case "Prime+"
      theavg = theavg + 10.5
      theagvcnt = theagvcnt + 1

    Case "Prime"
      theavg = theavg + 9.5
      theagvcnt = theagvcnt + 1

    Case "Prime-"
      theavg = theavg + 8.5
      theagvcnt = theagvcnt + 1
    
    Case "Choice+"
      theavg = theavg + 7.5
      theagvcnt = theagvcnt + 1
    
    Case "CAB"
      theavg = theavg + 6.75
      theagvcnt = theagvcnt + 1
    
    Case "STS"
      theavg = theavg + 6.75
      theagvcnt = theagvcnt + 1
    
    Case "Choice"
      theavg = theavg + 6.5
      theagvcnt = theagvcnt + 1
    
    Case "Choice-"
      theavg = theavg + 5.5
      theagvcnt = theagvcnt + 1
    
    Case "AAA"
      theavg = theavg + 5.5
      theagvcnt = theagvcnt + 1
    
    Case "Select+"
      theavg = theavg + 4.75
      theagvcnt = theagvcnt + 1
    
    Case "Select"
      theavg = theavg + 4.5
      theagvcnt = theagvcnt + 1
    
    Case "AA"
      theavg = theavg + 4.5
      theagvcnt = theagvcnt + 1
    
    Case "Select-"
      theavg = theavg + 4.25
      theagvcnt = theagvcnt + 1
    
    Case "Standard+"
      theavg = theavg + 3.5
      theagvcnt = theagvcnt + 1
    
    Case "A"
      theavg = theavg + 3.5
      theagvcnt = theagvcnt + 1
    
    Case "Standard"
      theavg = theavg + 3
      theagvcnt = theagvcnt + 1
    
    Case "Standard-"
      theavg = theavg + 2.5
      theagvcnt = theagvcnt + 1
    
    Case "B1"
      theavg = theavg + 1
      theagvcnt = theagvcnt + 1
    End Select
    tbData.MoveNext
   Wend
   If theagvcnt <> 0 Then
    tbavg.Edit
    Dim calc As Double
    Dim RESPONSE As Double
    calc = funround2(2, theavg / theagvcnt)
    If calc >= 10.5 And calc <= 10.99 Then RESPONSE = 10.5
    If calc >= 9.5 And calc <= 10.49 Then RESPONSE = 9.5
    If calc >= 8.5 And calc <= 9.49 Then RESPONSE = 8.5
    If calc >= 7.5 And calc <= 8.49 Then RESPONSE = 7.5
    If calc >= 6.75 And calc <= 7.49 Then RESPONSE = 6.75
    If calc >= 6.5 And calc <= 6.74 Then RESPONSE = 6.5
    If calc >= 5.5 And calc <= 6.49 Then RESPONSE = 5.5
    If calc >= 4.75 And calc <= 5.49 Then RESPONSE = 4.75
    If calc >= 4.5 And calc <= 4.74 Then RESPONSE = 4.5
    If calc >= 4.25 And calc <= 4.49 Then RESPONSE = 4.25
    If calc >= 3.5 And calc <= 4.24 Then RESPONSE = 3.5
    If calc >= 3# And calc <= 3.49 Then RESPONSE = 3
    If calc >= 2.5 And calc <= 2.99 Then RESPONSE = 2.5
    If calc >= 1# And calc <= 2.49 Then RESPONSE = 1
     
     tbavg!Avg_QScore = RESPONSE
     tbavg.Update
   End If
   tbavg.MoveNext
 Wend
 tbavg.Close: Set tbavg = Nothing
 tbData.Close: Set tbData = Nothing

 
Return

Update_2nd_Gen:
   TEXT(1) = "Update_2nd_Gen"
   DB.Execute "UPDATE (Pedigree LEFT JOIN sireprof ON (Pedigree.herdid = sireprof.HerdID) AND (Pedigree.Sire_ID_2 = sireprof.SireID)) LEFT JOIN cowprof ON (Pedigree.herdid = cowprof.HerdID) AND (Pedigree.Cow_ID_2 = cowprof.cowID) SET Pedigree.Sire_ID_3 = [sireprof].[sire], Pedigree.Cow_ID_3 = [sireprof].[dam], Pedigree.Sire_ID_4 = [cowprof].[sire], Pedigree.Cow_ID_4 = [cowprof].[dam], Pedigree.C2_regnum = [cowprof].[regnum], Pedigree.C2_regname = [cowprof].[regname], Pedigree.S2_regnum = [sireprof].[regnum], Pedigree.S2_regname = [sireprof].[regname], Pedigree.C1_regnum = [cowprof].[regnum], Pedigree.C1_regname = [cowprof].[regname]      "
   DB.Execute "UPDATE (((Pedigree LEFT JOIN sireprof ON (Pedigree.herdid = sireprof.HerdID) AND (Pedigree.Sire_ID_3 = sireprof.SireID)) LEFT JOIN cowprof ON (Pedigree.herdid = cowprof.HerdID) AND (Pedigree.Cow_ID_3 = cowprof.cowID)) LEFT JOIN sireprof AS sireprof_1 ON (Pedigree.herdid = sireprof_1.HerdID) AND (Pedigree.Sire_ID_4 = sireprof_1.SireID)) LEFT JOIN cowprof AS cowprof_1 ON (Pedigree.herdid = cowprof_1.HerdID) AND (Pedigree.Cow_ID_4 = cowprof_1.cowID) SET Pedigree.S3_regnum = [sireprof].[regnum], Pedigree.S3_regname = [sireprof].[regname], Pedigree.C3_regnum = [cowprof].[regnum], Pedigree.C3_regname = [cowprof].[regname], Pedigree.S4_regnum = [sireprof_1].[regnum], Pedigree.S4_regname = [sireprof_1].[regname], Pedigree.C4_regnum = [cowprof_1].[regnum], Pedigree.C4_regname = [cowprof_1].[regname]"
   DB.Execute "update pedigree, cowprof set pedigree.c1_regnum = cowprof.regnum, pedigree.c1_regname = cowprof.regname where cowprof.cowid = pedigree.cowid_1 and cowprof.herdid = pedigree.herdid "
Return

Update_3rd_Gen:
   TEXT(1) = "Update_3rd_Gen"
  
Return

Build_Sort_Order:
   TEXT(1) = "Build_Sort_Order"
600    Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
610    If OptPerf(0).Value Then orderby = " cowid_1 "
620    If OptPerf(1).Value Then orderby = " sire_id_2 "
630    DB.Execute "select * into tmpPedigree from pedigree order by " & orderby
640    DB.Execute "delete * from pedigree"
650    DB.Execute "insert into pedigree select * from tmpPedigree"
660    DB.Execute "drop table tmpPedigree"
670    DB.Execute "select * into tmpavgs from pedigree_avgs"
680    DB.Execute "delete * from pedigree_avgs"
690    DB.Execute "insert into pedigree_avgs select * from tmpavgs"
700    DB.Execute "drop table tmpavgs"
Return

End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
 selherd_List.Show vbModal
 If selherd_List.Tag = "CANCEL" Then Exit Sub
 herdid$ = selherd_List.Tag
 Unload selherd_List
 Call load_year
 Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCowSelect_Click()
mCowList.Show vbModal
If mCowList!lstCows.SelectedCount > 0 Then
   lblMultiCows.Caption = Trim$(Str$(mCowList!lstCows.SelectedCount))
Else
   lblMultiCows.Caption = "All"
End If
End Sub

Private Sub CMDOk_Click()
Dim order$, TITLE$, Herds$, title1$, title2$, Title3$, Title4$, XAvg#, XCows&
 Screen.MousePointer = vbHourglass
 report.Initialize ' init the class
 If optprint Then report.SetDestination = 1
 If Val(lblMultiCows.Caption) > 0 Then
   Call Build_Cow_RPT_List(mCowList.lstCows)
 End If
 Select Case lstreports.ListIndex
  Case 0
   Call Create_Cow_RefRPT
   report.SetReportFileName = dbdir$ & "\" & "cowREF.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = "Cow Reference List"
   report.Setcommonformulas(title1, title2, Title3) = Title4
  Case 1
   If frmPopUpCull.Tag = "True" Then
   Call create_Culled_refLIST_report
   report.SetReportFileName = dbdir$ & "\" & "culldREF.rpt"
   report.setDbname = repfile$
   'report.SetReportCaption = reports$(2)
   'report.Setcommonformulas("", "", "") = ""
  Else
   MsgBox "Please Select Cull Codes To Include In the Report", vbOKOnly + vbCritical, Me.Caption
   Screen.MousePointer = vbDefault
   frmPopUpCull.Show vbModal
   Exit Sub
  End If
   
   Case 2, 3
      Call Create_Cow_List(XAvg, XCows)
      Call CowOrder
      report.setDbname = repfile$
      report.SetReportCaption = "Lifetime Progeny Report"
      title1 = "Weaning Performance"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      If lstreports.ListIndex = 2 Then
         report.SetReportFileName = dbdir$ & "\cwlst.rpt"
         report.Setformulas("IntAvg") = "'Average calving interval for " & XCows & " cows is " & funround2(1, XAvg) & " days'"
      Else
         report.SetReportFileName = dbdir$ & "\cwlstdt.rpt"
         report.Setformulas("IntAvg") = "'Average calving interval for " & XCows & " cows is " & funround2(1, XAvg) & " days'"
      End If
   Case 4
      If optOrder(0).Value Then order = "LPR_Header.cowID"
      If optOrder(1).Value Then order = "LPR_Header.CowAge"
      If optOrder(2).Value Then order = "LPR_Header.MPPA desc, lpr_header.cowid"
      If optOrder(3).Value Then order = "LPR_Header.sire"
      Call Create_Lifetime_Progeny_RPTS(lstreports.ListIndex, order)
      report.setDbname = repfile$
      report.SetReportCaption = "Lifetime Progeny Report"
      title1 = "Background Performance"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\LPR_Back.rpt"
      report.Setformulas("Misc1") = "'" & IIf(calfhead(10) = "", "Misc1", calfhead(10)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(11) = "", "Misc2", calfhead(11)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(11) = "", "Misc3", calfhead(12)) & "'"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 5
      If optOrder(0).Value Then order = "LPR_Header.cowID"
      If optOrder(1).Value Then order = "LPR_Header.CowAge"
      If optOrder(2).Value Then order = "LPR_Header.MPPA desc, lpr_header.cowid"
      If optOrder(3).Value Then order = "LPR_Header.sire"
      Call Create_Lifetime_Progeny_RPTS(lstreports.ListIndex, order)
      report.setDbname = repfile$
      report.SetReportCaption = "Lifetime Progeny Report"
      title1 = "Replacement Performance"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\LPR_Repl.rpt"
      report.Setformulas("Misc1") = "'" & IIf(calfhead(10) = "", "Misc1", calfhead(10)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(11) = "", "Misc2", calfhead(11)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(11) = "", "Misc3", calfhead(12)) & "'"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 6
      If optOrder(0).Value Then order = "LPR_Header.cowID"
      If optOrder(1).Value Then order = "LPR_Header.CowAge"
      If optOrder(2).Value Then order = "LPR_Header.MPPA desc, lpr_header.cowid"
      If optOrder(3).Value Then order = "LPR_Header.sire"
      Call Create_Lifetime_Progeny_RPTS(lstreports.ListIndex, order)
      report.setDbname = repfile$
      report.SetReportCaption = "Lifetime Progeny Report"
      title1 = "Feedlot Performance"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\LPR_Feed.rpt"
      report.Setformulas("Misc1") = "'" & IIf(calfhead(13) = "", "Misc1", calfhead(13)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(14) = "", "Misc2", calfhead(14)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(15) = "", "Misc3", calfhead(15)) & "'"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
  Case 7
      If optOrder(0).Value Then order = "LPR_Header.cowID"
      If optOrder(1).Value Then order = "LPR_Header.CowAge"
      If optOrder(2).Value Then order = "LPR_Header.MPPA desc, lpr_header.cowid"
      If optOrder(3).Value Then order = "LPR_Header.sire"
      Call Create_Lifetime_Progeny_RPTS(lstreports.ListIndex, order)
      report.setDbname = repfile$
      report.SetReportCaption = "Lifetime Progeny Report"
      title1 = "Carcass Performance"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\LPR_Carc.rpt"
      report.Setformulas("Misc1") = "'" & IIf(calfhead(16) = "", "Misc1", calfhead(16)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(17) = "", "Misc2", calfhead(17)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(18) = "", "Misc3", calfhead(18)) & "'"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 8
      If Create_CowBreed_RPT = False Then Screen.MousePointer = 0: Exit Sub
      report.setDbname = repfile$
      report.SetReportCaption = reports(9)
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\CwBreed.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 9
       If Create_Cow_ConceptionRPT = False Then Screen.MousePointer = 0: Exit Sub
      report.setDbname = repfile$
      report.SetReportCaption = reports(10)
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\CwCncpt.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 10
      Call Create_Pedigree_RPT
      report.setDbname = repfile$
      report.SetReportCaption = reports(11)
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\Pedigree.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      If lblMultiCows.Caption <> "All" Then
         If mCowList.OptType(0).Value Then report.Setformulas("Status") = "'Active'"
         If mCowList.OptType(1).Value Then report.Setformulas("Status") = "'Culled'"
         If mCowList.OptType(2).Value Then report.Setformulas("Status") = "'Pedigree'"
      End If
      report.Setformulas("epd1") = "'" & IIf(epdhead1 = "", "Epd1", epdhead1) & "'"
      report.Setformulas("epd2") = "'" & IIf(epdhead2 = "", "Epd2", epdhead2) & "'"
      report.Setformulas("epd3") = "'" & IIf(epdhead3 = "", "Epd3", epdhead3) & "'"
      report.Setformulas("epd4") = "'" & IIf(epdhead4 = "", "Epd4", epdhead4) & "'"
      report.Setformulas("epd5") = "'" & IIf(epdhead5 = "", "Epd5", epdhead5) & "'"
      report.Setformulas("epd6") = "'" & IIf(epdhead6 = "", "Epd6", epdhead6) & "'"
      report.Setformulas("epd7") = "'" & IIf(epdhead7 = "", "Epd7", epdhead7) & "'"
      report.Setformulas("epd8") = "'" & IIf(epdhead8 = "", "Epd8", epdhead8) & "'"
      report.Setformulas("epd9") = "'" & IIf(epdhead9 = "", "Epd9", epdhead9) & "'"
      report.Setformulas("epd10") = "'" & IIf(epdhead10 = "", "Epd10", epdhead10) & "'"
   Case 11
      Call Create_Pedigree_RPT
      report.setDbname = repfile$
      report.SetReportCaption = reports(12)
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\CowPerf.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 12
      Call Create_Cow_NotesRPT
      report.setDbname = repfile$
      report.SetReportCaption = reports(13)
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.SetReportFileName = dbdir$ & "\cownotes.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
End Select
  report.PrintReport
  Screen.MousePointer = vbDefault
End Sub
 
Private Function Create_Cow_ConceptionRPT() As Boolean
Dim SQL$, DB As DAO.database, order$, RS As DAO.Recordset, TurnDate As Date, RS2 As DAO.Recordset, repdb As DAO.database
Dim indx%, orderby$
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from cwcncpt"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select currentdate from turnoutdate where herdid = '" & herdid & "'", dbOpenSnapshot)
If cboyear.TEXT <> "" And cboyear.TEXT <> "Not Set" Then TurnDate = Left(cboyear.TEXT, 10) Else GoTo Warning
If RS.EOF Then
Warning:
   MsgBox "Please Set Bull Turn Out Date For Herd " & herdid, vbOKOnly
   RS.Close: Set RS = Nothing
   DB.Close: Set DB = Nothing
   Create_Cow_ConceptionRPT = False
   Exit Function
End If
SQL = "insert into cwcncpt in '" & repfile & "'  SELECT DISTINCTROW cowprof.herdid, cowprof.cowID, cowprof.breed, Max(calfbirth.CowAge) AS Age, Max([calfbirth].[birthdate]) AS LastCalving, cowprof.mpda "
If Val(lblMultiCows.Caption) > 0 Then
   Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
   SQL = SQL & " FROM (cowprof LEFT JOIN calfbirth ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) INNER JOIN RPTCows ON (cowprof.cowID = RPTCows.CowID) AND (cowprof.HerdID = RPTCows.HerdID) "
Else
   SQL = SQL & " FROM cowprof LEFT JOIN calfbirth ON cowprof.cowID = calfbirth.CowID AND cowprof.HerdID = calfbirth.HerdID "
End If
SQL = SQL & " WHERE cowprof.active='A' and cowprof.herdid = '" & herdid & "' GROUP BY cowprof.herdid, cowprof.cowID, cowprof.breed, cowprof.mpda"

DB.Execute SQL
Call DeleteTableAttachment(dbfile, "RPTCows")
Screen.MousePointer = vbHourglass
Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
Set RS = repdb.OpenRecordset("select * from cwcncpt", dbOpenDynaset)
Do Until RS.EOF
   SQL = "select stat, breedcond, datediff('d', cowbrd.conceptdate, now) + cowbrd.age as EFA, (cowbrd.conceptdate - cowbrd.age + 285) " & IIf(Field2Date(RS!lastcalving) <> "--/--/----", "- #" & Field2Date(RS!lastcalving) & "#", "") & " as ECI from cowbrd where herdid = '" & Field2Str(RS!herdid) & "' and cowid = '" & Field2Str(RS!CowID) & "' and calfdate = #" & TurnDate & "# "
   Set RS2 = DB.OpenRecordset(SQL, dbOpenDynaset)
   If Not RS2.EOF Then
      RS.Edit
      RS!stat = Field2Str(RS2!stat)
      RS!condscore = Field2Num(RS2!breedcond)
      RS!efa = Field2Num(RS2!efa)
      RS!eci = Field2Num(RS2!eci)
      RS.Update
   End If
   RS2.Close: Set RS2 = Nothing
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
'build age cross tab
repdb.Execute "delete * from cwcncpt_age"
'this sql calculates the avg count of cow id's, date of last calving, condition score, estimated fetal age(AFE), # open and # pregnant for each year grouped by age
SQL = "insert into cwcncpt_age SELECT DISTINCTROW CwCncpt.Age, Count(CwCncpt.Age) AS CountOfAge, (SELECT DISTINCTROW " & _
   "Format(Avg([lastcalving]),'mm/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=2) AS LastCalving_2, (SELECT DISTINCTROW Form" & _
   "at(Avg([lastcalving]),'mm/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=3) AS LastCalving_3, (SELECT DISTINCTROW Format(Av" & _
   "g([lastcalving]),'mm/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=4) AS LastCalving_4, (SELECT DISTINCTROW Format(Avg([last" & _
   "calving]),'mm/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=5) AS LastCalving_5, (SELECT DISTINCTROW Format(Avg([lastcalvin" & _
   "g]),'mm/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=6) AS LastCalving_6, (SELECT DISTINCTROW Format(Avg([lastcalving]),'m" & _
   "m/dd/yyyy') From CwCncpt HAVING CwCncpt.Age=7) AS LastCalving_7, (SELECT DISTINCTROW Format(Avg([lastcalving]),'mm/dd/y" & _
   "yyy') From CwCncpt HAVING CwCncpt.Age=8) AS LastCalving_8, (SELECT DISTINCTROW Format(Avg([lastcalving]),'mm/dd/yyyy') " & _
   "From CwCncpt HAVING CwCncpt.Age=9) AS LastCalving_9, (SELECT DISTINCTROW Format(Avg([lastcalving]),'mm/dd/yyyy') From C" & _
   "wCncpt HAVING CwCncpt.Age=10) AS LastCalving_10, (SELECT DISTINCTROW Format(Avg([lastcalving]),'mm/dd/yyyy') From CwCn" & _
   "cpt HAVING CwCncpt.Age=11) AS LastCalving_11, (SELECT DISTINCTROW Format(Avg([lastcalving]),'mm/dd/yyyy') From CwCncpt H" & _
   "AVING CwCncpt.Age>=12) AS LastCalving_12, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt." & _
   "Age=2) AS AvgCond_2, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=3) AS AvgCond_" & _
   "3, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=4) AS AvgCond_4, (SELECT DISTINC" & _
   "TROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=5) AS AvgCond_5, (SELECT DISTINCTROW Avg(CwCncpt." & _
   "CondScore) From CwCncpt HAVING CwCncpt.Age=6) AS AvgCond_6, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From Cw" & _
   "Cncpt HAVING CwCncpt.Age=7) AS AvgCond_7, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCnc" & _
   "pt.Age=8) AS AvgCond_8, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=9) AS AvgCond" & _
   "_9, (SELECT DISTINCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=10) AS AvgCond_10, (SELECT DISTI" & _
   "NCTROW Avg(CwCncpt.CondScore) From CwCncpt HAVING CwCncpt.Age=11) AS AvgCond_11, (SELECT DISTINCTROW Avg(CwCn" & _
   "cpt.CondScore) From CwCncpt HAVING CwCncpt.Age>=12) AS AvgCond_12, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCn" & _
   "cpt HAVING CwCncpt.Age=2) AS Avg_AEF_2, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=3) " & _
   "AS Avg_AEF_3, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=4) AS Avg_AEF_4, (SELECT DIS" & _
   "TINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=5) AS Avg_AEF_5, (SELECT DISTINCTROW Avg(CwCncpt.EFA) " & _
   "From CwCncpt HAVING CwCncpt.Age=6) AS Avg_AEF_6, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt."
SQL = SQL & "Age=7) AS Avg_AEF_7, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=8) AS Avg_AEF" & _
   "_8, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=9) AS Avg_AEF_9, (SELECT DISTINCTROW Avg" & _
   "(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age=10) AS Avg_AEF_10, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt " & _
   "HAVING CwCncpt.Age=11) AS Avg_AEF_11, (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt HAVING CwCncpt.Age>=12) AS" & _
   " Avg_AEF_12, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING " & _
   " CwCncpt.Age=2) AS Open_2, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age" & _
   " HAVING CwCncpt.Age=3) AS Open_3, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY " & _
   "CwCncpt.Age HAVING CwCncpt.Age=4) AS Open_4, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' " & _
   "GROUP BY CwCncpt.Age HAVING CwCncpt.Age=5) AS Open_5, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCn" & _
   "cpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=6) AS Open_6, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHE" & _
   "RE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=7) AS Open_7, (SELECT DISTINCTROW Count(CwCncpt.Stat) From Cw" & _
   "Cncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=8) AS Open_8, (SELECT DISTINCTROW Count(CwCncpt.Stat) " & _
   "From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=9) AS Open_9, (SELECT DISTINCTROW Count(CwC" & _
   "ncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=10) AS Open_10, (SELECT DISTINCTROW " & _
   "Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=11) AS Open_11, (SELECT DIS" & _
   "TINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.Age HAVING CwCncpt.Age>=12) AS Open_12, (S" & _
   "ELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=2) AS O" & _
   "pen_P_2, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.A" & _
   "ge=3) AS Open_P_3, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING " & _
   "CwCncpt.Age=4) AS Open_P_4, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age " & _
   "HAVING CwCncpt.Age=5) AS Open_P_5, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY " & _
   "CwCncpt.Age HAVING CwCncpt.Age=6) AS Open_P_6, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' " & _
   "GROUP BY CwCncpt.Age HAVING CwCncpt.Age=7) AS Open_P_7, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwC" & _
   "ncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=8) AS Open_P_8, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt W" & _
   "HERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=9) AS Open_P_9, (SELECT DISTINCTROW Count(CwCncpt.Stat) From "
SQL = SQL & " CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=10) AS Open_P_10, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=11) AS Open_P_11, (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.Age HAVING CwCncpt.Age=12) AS Open_P_12 From CwCncpt GROUP BY CwCncpt.Age "

repdb.Execute SQL
'loop through cwcncpt_age and add missing years
Set RS2 = repdb.OpenRecordset("select * from cwcncpt_age", dbOpenDynaset)
For indx = 2 To 12
   RS2.FindFirst "age = " & indx
   If RS2.NoMatch Then
      RS2.AddNew
      RS2!age = indx
      RS2.Update
   End If
Next
RS2.Close: Set RS2 = Nothing
repdb.Execute "delete * from cwcncpt_cond"
'this sql generates avg's and counts for cowid, avg of last cow calving date, avg of age, estimated fetal age (EFA), avg of MPPA, counts of open/pregnant cows grouped by breeding condition
SQL = "insert into cwcncpt_cond SELECT DISTINCTROW CwCncpt.CondScore, Count(CwCncpt.CowID) as Count_CowID, (SELECT DISTINCTROW Count(CwCncpt.cowID) FROM CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS Cond_1, "
'avg last calving date
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS Last_1, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=2) AS Last_2, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=3) AS Last_3, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=4) AS Last_4, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=5) AS Last_5, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=6) AS Last_6, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=7) AS Last_7, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=8) AS Last_8, "
SQL = SQL & " (SELECT DISTINCTROW Format(Avg([LastCalving]),'mm/dd/yyyy') From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=9) AS Last_9,"
'avg cow age
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS Age_1, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=2) AS Age_2, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=3) AS Age_3, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=4) AS Age_4, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=5) AS Age_5, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=6) AS Age_6, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=7) AS Age_7, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=8) AS Age_8, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.Age) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=9) AS Age_9,"
'avg efa
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS EFA_1, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=2) AS EFA_2, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=3) AS EFA_3, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=4) AS EFA_4, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=5) AS EFA_5, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=6) AS EFA_6, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=7) AS EFA_7, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=8) AS EFA_8, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.EFA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=9) AS EFA_9,"
'avg mppa(mpda in chaps table)
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS MPPA_1, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=2) AS MPPA_2, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=3) AS MPPA_3, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=4) AS MPPA_4, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=5) AS MPPA_5, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=6) AS MPPA_6, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=7) AS MPPA_7, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=8) AS MPPA_8, "
SQL = SQL & " (SELECT DISTINCTROW Avg(CwCncpt.MPDA) From CwCncpt GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=9) AS MPPA_9, "
'open status count
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.CondScore HAVING CwCncpt.CondScore=1) AS Open_1, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=2) AS Open_2, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=3) AS Open_3, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=4) AS Open_4, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=5) AS Open_5, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=6) AS Open_6, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=7) AS Open_7, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=8) AS Open_8, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='O' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=9) AS Open_9, "
'pregnant status count
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=1) AS Open_P_1, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=2) AS Open_P_2, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=3) AS Open_P_3, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=4) AS Open_P_4, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=5) AS Open_P_5, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=6) AS Open_P_6, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=7) AS Open_P_7, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=8) AS Open_P_8, "
SQL = SQL & " (SELECT DISTINCTROW Count(CwCncpt.Stat) From CwCncpt WHERE CwCncpt.Stat ='P' GROUP BY CwCncpt.condscore HAVING CwCncpt.condscore=9) AS Open_P_9 From CwCncpt GROUP BY CwCncpt.CondScore"

repdb.Execute SQL
'loop through cwcncpt_cond and add missing years
Set RS2 = repdb.OpenRecordset("select * from cwcncpt_cond", dbOpenDynaset)
For indx = 1 To 9
   RS2.FindFirst "condscore = " & indx
   If RS2.NoMatch Then
      RS2.AddNew
      RS2!condscore = indx
      RS2.Update
   End If
Next
RS2.Close: Set RS2 = Nothing

'correct order by
If optOrder(4).Value Then orderby = " cowid "
If optOrder(5).Value Then orderby = " age "
If optOrder(6).Value Then orderby = " efa "
If optOrder(7).Value Then orderby = " eci "
If optOrder(8).Value Then orderby = " condscore "

repdb.Execute "select * into tmp_cwcncpt from cwcncpt order by " & orderby
repdb.Execute "delete * from cwcncpt"
repdb.Execute "insert into cwcncpt select * from tmp_cwcncpt"
repdb.Execute "drop table tmp_cwcncpt"

repdb.Close: Set repdb = Nothing
Create_Cow_ConceptionRPT = True
report.Setformulas("ExpDate") = "'Exposed Date: " & TurnDate & "'"
Screen.MousePointer = vbDefault
End Function
 
Private Sub cmdselectvend_Click()
 'FrmSelect_Multi_Herds.Show vbModal
 'If FrmSelect_Multi_Herds!lstherd.SelectedCount > 0 Then
 '  lblhow_many_herd.Caption = Trim$(Str$(FrmSelect_Multi_Herds!lstherd.SelectedCount))
 ' Else
 '  lblhow_many_herd.Caption = "All"
 'End If
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "Cow Reference List"
 reports$(2) = "Culled Reference List"
 reports$(3) = "Lifetime Progeny Report -- Short"
 reports$(4) = "Lifetime Progeny Report -- Weaning"
 reports$(5) = "Lifetime Progeny Report -- Background"
 reports$(6) = "Lifetime Progeny Report -- Replacement"
 reports$(7) = "Lifetime Progeny Report -- Feedlot"
 reports$(8) = "Lifetime Progeny Report -- Carcass"
 reports(9) = "Cow Breeding Report"
 reports(10) = "Cow Breeding and Conception Report"
 reports(11) = "Cow Performance Pedigree Report"
 reports(12) = "Cow Performance Report"
 reports(13) = "Cow Notes Report"
 For t = 1 To hmreps%
     lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optpreview.Value = True
 OptInclude(0).Value = True
 optOrder(0).Value = True
 optOrder(4).Value = True
 OptPerf(0).Value = True
 OptPerf(3).Value = True
Load mCowList 'Select multi cow form
Load FrmSelect_Multi_Herds
Call load_year
'lblhow_many_herd.Caption = "1"
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload mCowList 'Select multi cow form
Load FrmSelect_Multi_Herds
Set FrmSelect_Multi_Herds = Nothing
Set mCowList = Nothing
End Sub

Private Sub lstreports_Click()
   Call Display_Criteria
End Sub

Private Sub Display_Criteria()
'optOrder(0).Value = True
'mCowList.OptType(1).Enabled = True
'mCowList.OptType(2).Enabled = True
LBLBTOD.Visible = False
cboyear.Visible = False
FraRef.Visible = False
FraBrd.Visible = False
FraPerf.Visible = False
fraCowList.Visible = False
OptPerf(1).Caption = "Sire ID"
Select Case lstreports.ListIndex
Case 0
   FraRef.Visible = True
   fraCowList.Visible = False
   'lblCows.Enabled = False
   'lblMultiCows.Enabled = False
   'cmdCowSelect.Enabled = False
   'optOrder(0).Enabled = False
   'optOrder(1).Enabled = False
   'optOrder(2).Enabled = False
   'optOrder(3).Enabled = False
   'chkDT.Enabled = False
   optOrder(1).Caption = "Cow Age"
   optOrder(2).Visible = False
   optOrder(3).Visible = False
   ChkDT.Visible = False
Case 1
   FraRef.Visible = False
   fraCowList.Visible = False
   frmPopUpCull.Show vbModal
   'LblCows.Enabled = False
   'lblMultiCows.Enabled = False
   'cmdCowSelect.Enabled = False
   'optOrder(0).Enabled = False
   'optOrder(1).Enabled = False
   'optOrder(2).Enabled = False
   'optOrder(3).Enabled = False
   'chkDT.Enabled = False
   optOrder(1).Caption = "Cow Age"
   optOrder(2).Visible = True
   optOrder(3).Visible = True
   ChkDT.Visible = False
Case 2, 3, 4, 5, 6, 7
   'mCowList.OptType(1).Enabled = False
   'mCowList.OptType(2).Enabled = False
   FraRef.Visible = False
   lblCows.Enabled = True
   lblMultiCows.Enabled = True
   fraCowList.Visible = True
   cmdCowSelect.Enabled = True
   optOrder(0).Enabled = True
   optOrder(1).Enabled = True
   optOrder(2).Enabled = True
   optOrder(3).Enabled = True
   ChkDT.Visible = False
   optOrder(1).Caption = "Cow Age"
   optOrder(2).Visible = True
   optOrder(3).Visible = True
   'ChkDT.Visible = False
 Case 8
   FraRef.Visible = False
   cboyear.Visible = True
   LBLBTOD.Visible = True
   lblCows.Enabled = True
   lblMultiCows.Enabled = True
   fraCowList.Visible = True
   cmdCowSelect.Enabled = True
   optOrder(0).Enabled = True
   
   optOrder(1).Caption = "Preg Status"
   optOrder(2).Visible = False
   optOrder(3).Visible = False
   ChkDT.Visible = False
  Case 9
   LBLBTOD.Visible = True
   cboyear.Visible = True
   FraBrd.Visible = True
  Case 10
    FraPerf.Visible = True
  Case 11
    FraPerf.Visible = True
  Case 12
    FraPerf.Visible = True
    OptPerf(1).Caption = "Birth Date"
End Select
Exit Sub
'Disable:
'   FraRef.Visible = False
'   fraCowList.Visible = False
'
End Sub

Private Sub load_year()
'Dim indx As Integer, INDX2 As Integer, OLDDATE$(), CurDate$
'Screen.MousePointer = vbHourglass
'cboyear.Clear
'CurDate = ReturnBullTurnOutDate(herdid$, OLDDATE())
'If CurDate = "" Then Exit Sub
'Do Until indx = 5
'    If OLDDATE(indx) = "--/--/----" Then
'      cboyear.AddItem "Not Set"
'    Else
'      cboyear.AddItem OLDDATE(indx), indx
'    End If
'    indx = indx + 1
'Loop
'cboyear.AddItem CurDate & "*", 0
Dim DB As DAO.database, RS As DAO.Recordset
Dim SQL As String
Screen.MousePointer = vbHourglass
cboyear.Clear
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
SQL = "select DISTINCT calfdate from cowbrd where herdid = '" & herdid$ & "' order by calfdate desc"
Set RS = DB.OpenRecordset(SQL, dbOpenSnapshot)
While Not RS.EOF
   cboyear.AddItem RS!calfdate
   RS.MoveNext
Wend
If cboyear.ListCount > 0 Then cboyear.ListIndex = 0
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
Exit Sub
End Sub

