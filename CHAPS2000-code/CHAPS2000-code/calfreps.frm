VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form calfreps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calf Reports"
   ClientHeight    =   4950
   ClientLeft      =   3480
   ClientTop       =   825
   ClientWidth     =   6510
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   6510
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   2175
      Left            =   105
      TabIndex        =   1
      Top             =   240
      Width           =   3150
      _Version        =   65536
      _ExtentX        =   5556
      _ExtentY        =   3836
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
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Herd"
      Height          =   345
      Left            =   5025
      TabIndex        =   55
      Top             =   240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   810
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
      Left            =   5040
      TabIndex        =   2
      Top             =   1980
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   3600
      TabIndex        =   0
      Top             =   1980
      Width           =   1000
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort Criteria"
      Height          =   2415
      Left            =   3360
      TabIndex        =   9
      Top             =   2520
      Width           =   3135
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton OptSort 
            Caption         =   "Calf Age"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Act Wean Wt"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Sire ID"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Adj 205 Wt"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Calf ID"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Cow Age"
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   22
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Birth Wt"
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Frame Score"
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   20
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contemporary Groups"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2895
         Begin VB.OptionButton optCont 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optCont 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Frame FraPostWean 
      Height          =   2415
      Left            =   105
      TabIndex        =   34
      Top             =   2520
      Width           =   3135
      Begin VB.Frame FraHarvestDates 
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   150
         TabIndex        =   56
         Top             =   1215
         Width           =   2880
         Begin Threed.SSCommand SSCommand5 
            Height          =   270
            Left            =   2550
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   476
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "calfreps.frx":0000
         End
         Begin MSMask.MaskEdBox txtPWSH 
            Height          =   315
            Left            =   1560
            TabIndex        =   58
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin Threed.SSCommand SSCommand6 
            Height          =   270
            Left            =   2565
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   660
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   476
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "calfreps.frx":04BE
         End
         Begin MSMask.MaskEdBox txtPWEH 
            Height          =   315
            Left            =   1560
            TabIndex        =   60
            Top             =   630
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "End Harvest Date"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   630
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Start Harvest Date"
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   285
            Width           =   1455
         End
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   2670
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   270
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":097C
      End
      Begin MSMask.MaskEdBox txtPWSBD 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   270
         Left            =   2685
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   615
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":0E3A
      End
      Begin MSMask.MaskEdBox txtPWEB 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "End Birthdate"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Birthdate"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame fraDate 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3135
      Begin VB.ComboBox cboyear 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CheckBox chkOverwrite 
         Caption         =   "Overwrite or Create 205 Adj Ratios"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   1920
         Width           =   2895
      End
      Begin Threed.SSCommand cmdCal 
         Height          =   270
         Left            =   2685
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   390
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":12F8
      End
      Begin MSMask.MaskEdBox txtStartDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand cmdCal1 
         Height          =   270
         Left            =   2685
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":17B6
      End
      Begin MSMask.MaskEdBox txtEndDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   270
         Left            =   2685
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1110
         Visible         =   0   'False
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":1C74
      End
      Begin MSMask.MaskEdBox txtTurnDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   270
         Left            =   2685
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1470
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "calfreps.frx":2132
      End
      Begin MSMask.MaskEdBox txtWeighDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   1440
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label2 
         Caption         =   "Weigh Date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Bull Turn Out Date"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEnd 
         Caption         =   "End Birthdate"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblStart 
         Caption         =   "Start Birthdate"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame FraPW 
      Caption         =   "Post Weaning Reports"
      Height          =   2415
      Left            =   3360
      TabIndex        =   41
      Top             =   2535
      Width           =   3135
      Begin VB.Frame FraPostWeanSex 
         Caption         =   "Sex Groups"
         Height          =   975
         Left            =   60
         TabIndex        =   42
         Top             =   180
         Width           =   3015
         Begin VB.CheckBox chkSex 
            Caption         =   "Misc"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkSex 
            Caption         =   "Bulls"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkSex 
            Caption         =   "Heifers"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkSex 
            Caption         =   "Steers"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame FraPostWeanSort 
         Caption         =   "Sort By"
         Height          =   1215
         Left            =   60
         TabIndex        =   47
         Top             =   1140
         Width           =   3015
         Begin VB.OptionButton optOrder 
            Caption         =   "Fr Score"
            Height          =   195
            Index           =   6
            Left            =   1680
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Fr Score"
            Height          =   195
            Index           =   5
            Left            =   1680
            TabIndex        =   53
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "WDA"
            Height          =   195
            Index           =   4
            Left            =   1680
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "ADG"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   1275
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Birth Wt"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   1515
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "ID"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "365 Wt"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "calfreps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 13
Dim title1$, title2$, Title3$, Title4$
Dim t As Integer
Dim reports(hmreps%) As String

Private Sub Create_Calf_NotesRPT()
Dim SQL$, DB As DAO.database
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from calf_notes"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
SQL = "insert into calf_notes in '" & repfile & "' SELECT DISTINCTROW calfbirth.HerdID, calfbirth.CalfID, calfbirth.notes as calfbirth_notes, calfback.notes as calfback_notes, calffeed.notes as calffeed_notes, calfcarcass.notes as calfcarcass_notes, calfrep.notes as calfrep_notes, calfwean.notes as calfwean_notes, calfbirth.birthdate FROM ((((calfbirth LEFT JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID)) LEFT JOIN calfback ON (calfbirth.CalfID = calfback.CalfID) AND (calfbirth.HerdID = calfback.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID) where calfbirth.herdid = '" & herdid & "' "
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then SQL = SQL & " and calfbirth.birthdate >= #" & txtPWSBD.TEXT & "# and calfbirth.birthdate <= #" & txtPWEB.TEXT & "# "
SQL = SQL & " order by calfbirth.calfid"
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub load_year()
'Dim OLDDATE$(), CurDate$, indx As Integer, INDX2 As Integer
'Screen.MousePointer = vbHourglass
'cboyear.Clear
'CurDate = ReturnBullTurnOutDate(herdid$, OLDDATE())
'If CurDate = "" Then Exit Sub
'Do Until indx = 5
'    If OLDDATE(indx) <> "" Then cboyear.AddItem OLDDATE(indx)
'    indx = indx + 1
'Loop
'cboyear.AddItem CurDate & "*", 0
'If cboyear.ListCount > 0 Then cboyear.ListIndex = 0




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



Me.Show
fraDate.Refresh
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub CreateCR_Carc()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calfcarcass.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "#"
End If
SQL = "insert into CR_Carc in '" & repfile & "' SELECT DISTINCTROW calfcarcass.HerdID, calfcarcass.CalfID, calfcarcass.carcassdate, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.yribeye, calfcarcass.qgrade, calfcarcass.qcolor, calfcarcass.qtexture, calfcarcass.qmaturity, calfcarcass.misc1, calfcarcass.misc2, calfcarcass.misc3, calfcarcass.QSCORE FROM calfbirth RIGHT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID) " & where & " ORDER BY calfcarcass.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_carc"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub CreateCR_Feed()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calffeed.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "#"
End If
SQL = "insert into CR_Feed in '" & repfile & "' SELECT DISTINCTROW calffeed.HerdID, calffeed.CalfID, calffeed.int1date, calffeed.int1wt, calffeed.int2date, calffeed.int2wt, calffeed.findate, calffeed.finwt, calffeed.misc1, calffeed.misc2, calffeed.misc3 FROM calfbirth RIGHT JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID) " & where & " ORDER BY calffeed.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_feed"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub CreateCR_Repl()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calfrep.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "#"
End If
SQL = "insert into CR_Repl in '" & repfile & "' SELECT DISTINCTROW calfrep.HerdID, calfrep.CalfID, calfrep.recdate, calfrep.recwt, calfrep.rechip, calfrep.intdate, calfrep.intwt, calfrep.inthip, calfrep.daydate, calfrep.daywt, calfrep.dayhip, calfrep.misc1, calfrep.misc2, calfrep.misc3 FROM calfbirth RIGHT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID) " & where & " ORDER BY calfrep.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_repl"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub CreateCR_Back()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calfback.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "#"
End If
SQL = "insert into CR_Back in '" & repfile & "' SELECT DISTINCTROW calfback.HerdID, calfback.CalfID, calfback.recdate, calfback.recweight, calfback.recheight, calfback.intdate, calfback.intweight, calfback.intheight, calfback.findate, calfback.finweight, calfback.finheight, calfback.misc1, calfback.misc2, calfback.misc3 FROM calfback LEFT JOIN calfbirth ON (calfback.HerdID = calfbirth.HerdID) AND (calfback.CalfID = calfbirth.CalfID) " & where & " ORDER BY calfback.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_back"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub CreateCR_Wean()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calfwean.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "#"
End If
SQL = "insert into CR_Wean in '" & repfile & "' SELECT DISTINCTROW calfwean.HerdID, calfwean.CalfID, calfwean.actweight, calfwean.managecode, calfwean.chipheight, calfwean.cdatemeas, calfwean.group, calfwean.grade, calfwean.misc1, calfwean.misc2, calfwean.misc3, calfwean.misc4, calfwean.misc5, calfwean.misc6 FROM calfbirth RIGHT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)  " & where & " ORDER BY calfwean.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_wean"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub CreateCR_Birth()
Dim SQL$, DB As DAO.database, StartDate As Date, enddate As Date, where$
where = " WHERE calfbirth.herdid = '" & herdid & "'"
If txtPWSBD <> "--/--/----" And txtPWEB <> "--/--/----" Then
   StartDate = CDate(txtPWSBD)
   enddate = CDate(txtPWEB)
   where = where & " and calfbirth.birthdate >= #" & StartDate & "# and calfbirth.birthdate <= #" & enddate & "# "
End If
SQL = "insert into CR_Birth in '" & repfile & "' SELECT DISTINCTROW calfbirth.HerdID, calfbirth.CalfID, calfbirth.sireID, calfbirth.CowID, calfbirth.CowAge, calfbirth.breed, calfbirth.sex, calfbirth.birthdate, calfbirth.birthwt, calfbirth.calvingease, calfbirth.registration, calfbirth.regname, calfbirth.elecid From calfbirth " & where & " ORDER BY calfbirth.CalfID"
Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
DB.Execute "delete * from cr_birth"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub

Private Sub create_refLIST_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
 ' On Error Resume Next
  DBREP.Execute ("delete * from calfref")
  DBREP.Close
  SQL$ = "insert into calfref in '" & repfile$ & "' select * "
  SQL$ = SQL$ & " from calfBIRTH where calfbirth.herdid = '" & herdid & "'"
  'hmfields% = 1
  'Col(1) = 0
  'fieldvar$(1) = "calfBIRTH.herdid"
  'If lblhow_many_herd <> "All" Then
  '   Call create_sql_selection(FrmSelect_Multi_Herds!lstherd, Col(), fieldvar$(), hmfields%, formula1$)
     'SQL$ = SQL$ & "where " & formula1$
  '   If formula1 <> "" Then SQL = SQL & " where " & formula1
  'End If
  'MsgBox sql$
  DB.Execute (SQL$)
  DB.Close
End Sub

Private Function validform() As Boolean
Dim iResponse As Integer, tbMisc As Recordset
validform = True
If herdid$ = "" Then
      MsgBox "Please Select A Herd", vbOKOnly + vbExclamation, Me.Caption
      Screen.MousePointer = vbDefault
      validform = False
      Exit Function
End If
If cboyear.ListCount = 0 Then
      MsgBox "Please Set This Herd's Default Exposed Date", vbOKOnly + vbExclamation, Me.Caption
      
      Screen.MousePointer = vbDefault
      validform = False
      Exit Function
End If
If txtWeighDate.TEXT <> "--/--/----" Then
      If CDate(txtWeighDate.TEXT) > Date + 365 Or CDate(txtWeighDate.TEXT) < Date - 365 Then
         iResponse = MsgBox("WARNING:" & vbCrLf & "THE INTEGRITY OF YOUR DATA IS IN JEOPARDY!" & vbCrLf & "Executing this option on past production records will recalculate the 205 Day Adjusted Weight Ratios and the MPPA values that are currently being stored.  Overall Herd Performance Measures of Production and Reproduction will also be re-calculated.  Do not run this option unless you fully understand the ramifications of reclassifying contemporary groups.", vbOKCancel + vbCritical, Me.Caption)
         Screen.MousePointer = vbDefault
         If iResponse = vbCancel Then validform = False: Exit Function:
      End If
End If
If txtStartDate.TEXT = "--/--/----" Or txtEndDate.TEXT = "--/--/----" Then Exit Function
If DateDiff("d", CDate(txtStartDate.TEXT), CDate(txtEndDate.TEXT)) > 365 Then
         If chkOverwrite.Value <> vbChecked Then Exit Function
         iResponse = MsgBox("WARNING:" & vbCrLf & "THE INTEGRITY OF YOUR DATA IS IN JEOPARDY!" & vbCrLf & "Executing this option on past production records will recalculate the 205 Day Adjusted Weight Ratios and the MPPA values that are currently being stored.  Overall Herd Performance Measures of Production and Reproduction will also be re-calculated.  Do not run this option unless you fully understand the ramifications of reclassifying contemporary groups.", vbOKCancel + vbCritical, Me.Caption)
         Screen.MousePointer = vbDefault
         If iResponse = vbCancel Then validform = False: Exit Function:
End If
End Function


Private Sub CMDCancel_Click()
 Unload Me
End Sub


Private Sub cmdchange_Click()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
herdid$ = selherd_List.Tag
chkOverwrite.Enabled = True
Call load_year
Unload selherd_List
Screen.MousePointer = vbDefault
End Sub

Private Sub CMDOk_Click()
 Dim rowcount As Double, StartDate$, enddate$, DB As database, RS As Recordset
 Dim SortOrder As Integer, age As Double, avgwt As Double, allcalf As Double, pos%, order$
 report.Initialize ' init the class
 Screen.MousePointer = vbHourglass
 If optprint Then report.SetDestination = 1
 Select Case lstreports.ListIndex
  Case 0
   Call create_refLIST_report
      report.SetReportFileName = dbdir$ & "calfREF.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(1)
   Case 1
   Call CreateCR_Birth
      report.SetReportFileName = dbdir$ & "CR_Birth.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(2)
      report.Setorientation = False
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 2
   Call CreateCR_Wean
      report.SetReportFileName = dbdir$ & "CR_Wean.rpt"
      report.Setformulas("Misc1") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(1)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(2) = "", "Misc2", calfhead(2)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(3) = "", "Misc3", calfhead(3)) & "'"
      report.Setformulas("Misc4") = "'" & IIf(calfhead(4) = "", "Misc4", calfhead(4)) & "'"
      report.Setformulas("Misc5") = "'" & IIf(calfhead(5) = "", "Misc5", calfhead(5)) & "'"
      report.Setformulas("Misc6") = "'" & IIf(calfhead(6) = "", "Misc6", calfhead(6)) & "'"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(3)
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 3
   Call CreateCR_Back
      report.SetReportFileName = dbdir$ & "CR_Back.rpt"
      report.setDbname = repfile$
      report.Setformulas("Misc1") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(7)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(2) = "", "Misc2", calfhead(8)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(3) = "", "Misc3", calfhead(9)) & "'"
      report.SetReportCaption = reports$(4)
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 4
   Call CreateCR_Repl
      report.SetReportFileName = dbdir$ & "CR_Repl.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(5)
      report.Setformulas("Misc1") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(10)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(2) = "", "Misc2", calfhead(11)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(3) = "", "Misc3", calfhead(12)) & "'"
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 5
   Call CreateCR_Feed
      report.SetReportFileName = dbdir$ & "CR_Feed.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(6)
      report.Setformulas("Misc1") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(13)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(2) = "", "Misc2", calfhead(14)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(3) = "", "Misc3", calfhead(15)) & "'"
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 6
   Call CreateCR_Carc
      report.SetReportFileName = dbdir$ & "CR_Carc.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(7)
      report.Setformulas("Misc1") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(16)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(2) = "", "Misc2", calfhead(17)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(3) = "", "Misc3", calfhead(18)) & "'"
      report.Setcommonformulas(title1, title2, Title3) = Title4
   Case 7
   If Not validform Then Exit Sub
   If optCont(0).Value Then group = True Else group = False
      If OptSort(0).Value = True Then SortOrder = 0
      If OptSort(1).Value = True Then SortOrder = 1
      If OptSort(2).Value = True Then SortOrder = 2
      If OptSort(3).Value = True Then SortOrder = 3
      If OptSort(4).Value = True Then SortOrder = 4
      If OptSort(5).Value = True Then SortOrder = 5
      If OptSort(6).Value = True Then SortOrder = 6
      If OptSort(7).Value = True Then SortOrder = 7
   If IsDate(Left(cboyear.TEXT, 10)) Then TurnDate = Left(cboyear.TEXT, 10)
   'FrmSelect_Multi_Herds.lstherd.Col = 1
   'pos% = InStr(1, FrmSelect_Multi_Herds.lstherd.TEXT, vbTab)
   'herdid$ = Trim$(Left$(FrmSelect_Multi_Herds.lstherd.TEXT, pos% - 1))
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
   Set RS = DB.OpenRecordset("select * from herd where herdid = '" & herdid$ & "'", dbOpenSnapshot)
   If RS.RecordCount > 0 Then
   report.Setformulas("Desc") = "'" & RS!herddesc & "'"
   report.Setformulas("name") = "'" & RS!herdName & "'"
   report.Setformulas("address") = "'" & RS!address & "'"
   report.Setformulas("citystatezip2") = "'" & RS!city & " " & RS!state & " " & RS!zip & "'"
   End If
   RS.Close: Set RS = Nothing
   DB.Close: Set DB = Nothing
      title2$ = "Weigh Date: " & txtWeighDate.TEXT
      title1$ = "Herd ID: " & herdid$
      Title3$ = "Birthdate Range: " & calfreps.txtStartDate.TEXT & " to " & calfreps.txtEndDate.TEXT
      report.Setformulas("title1") = "'" & title1$ & "'"
      report.Setformulas("title2") = "'" & title2$ & "'"
      report.Setformulas("title3") = "'" & Title3$ & "'"
      If group = True Then
         report.SetReportFileName = dbdir$ & "calfrepg.rpt"
      Else
         report.Setorientation = True
         report.SetReportFileName = dbdir$ & "calfrep.rpt"
      End If
      report.setDbname = repfile$
      report.Setformulas("Report_Title") = "'" & reports$(8) & "'"
      'report.SetReportCaption = reports(8)
      report.Setformulas("MiscLabel") = "'" & IIf(calfhead(1) = "", "Misc1", calfhead(1)) & "'"
      Call CreateCalfReps(SortOrder, IIf(IsDate(txtStartDate.TEXT), txtStartDate.TEXT, #1/1/1900#), IIf(IsDate(txtStartDate.TEXT), txtEndDate.TEXT, #12/30/2099#), IIf(chkOverwrite.Value, True, False))
      If EndSub Then Screen.MousePointer = vbDefault: Exit Sub
   Case 8
      Call BuildPW_Back
      report.SetReportCaption = reports(9)
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      report.Setformulas("BirthDateRange") = "'Birthdate Range :" & txtPWSBD & " To " & txtPWEB & "'"
      report.Setformulas("HarvestDateRange") = "'Harvest Date Range: " & txtPWSH & " To " & txtPWEH & "'"
      report.setDbname = repfile$
      report.SetReportFileName = dbdir$ & "PW_Back.rpt"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.Setformulas("Misc1") = "'" & IIf(calfhead(7) = "", "Misc1", calfhead(7)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(8) = "", "Misc2", calfhead(8)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(9) = "", "Misc3", calfhead(9)) & "'"
      If optOrder(0).Value Then report.Setformulas("SortField") = "{PW_Back.CalfID}"
      If optOrder(1).Value Then report.Setformulas("SortField") = "{PW_Back.Days_On_Feed}"
      If optOrder(2).Value Then report.Setformulas("SortField") = "{PW_Back.Interim_ADG}"
      If optOrder(3).Value Then report.Setformulas("SortField") = "{PW_Back.Final_ADG}"
      If optOrder(4).Value Then report.Setformulas("SortField") = "{PW_Back.intweight}"
      If optOrder(5).Value Then report.Setformulas("SortField") = "{PW_Back.finweight}"
      'If optOrder(6).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.Dressing_Percent}"
   Case 9
      Call Create365Reports(order$, txtPWSBD, txtPWEB, txtPWEH, txtPWSH)
      Call LoadSexRange
      If optOrder(0).Value Then report.Setformulas("SortField") = "{Yearlings.Wgt365}"
      If optOrder(1).Value Then report.Setformulas("SortField") = "{Yearlings.YearID}"
      If optOrder(2).Value Then report.Setformulas("SortField") = "{Yearlings.ActWeight}"
      If optOrder(3).Value Then report.Setformulas("SortField") = "{Yearlings.WDAOff}"
      If optOrder(4).Value Then report.Setformulas("SortField") = "{Yearlings.FrScore}"
      If optOrder(5).Value Then report.Setformulas("SortField") = "{Yearlings.ADGOn}"
      'report.SetReportCaption = reports(10)
      report.Setformulas("Report_Caption_Title") = "'" & reports(10) & "'"
      report.SetReportFileName = dbdir$ & "\" & "yrreps.rpt"
      title2$ = "Birthdate Range:" & IIf(IsDate(txtPWSBD), CStr(txtPWSBD), "--/--/----") & " -- " & IIf(IsDate(txtPWEB), CStr(txtPWEB), "--/--/----")
      Title3$ = "Weigh Date Range: " & txtPWSH & " -- " & txtPWEH
      report.Setcommonformulas(title1$, title2$, Title3$) = Title4$
      report.setDbname = repfile$
      report.Setformulas("Misc1") = "'" & IIf(calfhead(10) = "", "Misc1", calfhead(10)) & "'"
      report.Setformulas("HerdID") = "'Herd ID: " & herdid & "'"
   Case 10
      Call BuildPW_Feed
      report.SetReportCaption = reports(11)
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      report.Setformulas("BirthDateRange") = "'Birthdate Range :" & txtPWSBD & " To " & txtPWEB & "'"
      report.Setformulas("HarvestDateRange") = "'Harvest Date Range: " & txtPWSH & " To " & txtPWEH & "'"
      report.setDbname = repfile$
      report.SetReportFileName = dbdir$ & "PW_Feed.rpt"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.Setformulas("Misc1") = "'" & IIf(calfhead(13) = "", "Misc1", calfhead(13)) & "'"
      report.Setformulas("Misc2") = "'" & IIf(calfhead(14) = "", "Misc2", calfhead(14)) & "'"
      report.Setformulas("Misc3") = "'" & IIf(calfhead(15) = "", "Misc3", calfhead(15)) & "'"
      If optOrder(0).Value Then report.Setformulas("SortField") = "{PW_Feed.CalfID}"
      If optOrder(1).Value Then report.Setformulas("SortField") = "{PW_Feed.Days_On_Feed}"
      If optOrder(2).Value Then report.Setformulas("SortField") = "{PW_Feed.Interim_ADG}"
      If optOrder(3).Value Then report.Setformulas("SortField") = "{PW_Feed.Final_ADG}"
      If optOrder(4).Value Then report.Setformulas("SortField") = "{PW_Feed.int2wt}"
      If optOrder(5).Value Then report.Setformulas("SortField") = "{PW_Feed.finwt}"
      'If optOrder(6).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.Dressing_Percent}"
   Case 11
      Call BuildPW_Carc
      'if optsort(0).Value then report.Setformulas("SortField") =
      report.SetReportCaption = reports(12)
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      report.Setformulas("BirthDateRange") = "'Birthdate Range :" & txtPWSBD & " To " & txtPWEB & "'"
      report.Setformulas("HarvestDateRange") = "'Harvest Date Range: " & txtPWSH & " To " & txtPWEH & "'"
      report.setDbname = repfile$
      report.SetReportFileName = dbdir$ & "PW_Carc.rpt"
      report.Setcommonformulas(title1, title2, Title3) = Title4
      report.Setformulas("Misc1") = "'" & IIf(calfhead(16) = "", "Misc1", calfhead(16)) & "'"
      If optOrder(0).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.CalfID}"
      If optOrder(1).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.ygrade}"
      If optOrder(2).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.qgrade}"
      If optOrder(3).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.ywt}"
      If optOrder(4).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.yfat}"
      If optOrder(5).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.yribeye}"
      If optOrder(6).Value Then report.Setformulas("SortField") = "{PostWean_CarcassDt.Dressing_Percent}"
   Case 12
      Call Create_Calf_NotesRPT
      report.SetReportCaption = reports(13) & " Report"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      report.setDbname = repfile$
      report.SetReportFileName = dbdir$ & "CalfNote.rpt"
      report.Setcommonformulas(title1, title2, Title3) = Title4
End Select
 
 Screen.MousePointer = vbDefault
'report.Setcommonformulas(title1, title2, Title3) = Title4
 report.PrintReport
  'Call DeleteTableAttachment(repfile$, "herd")
Set report = Nothing
End Sub
 
Private Sub BuildPW_Back()
Dim DB As database, RS As Recordset, SQL$, where$
If txtPWSBD <> "--/--/----" Then where = where & " and calfbirth.birthdate >= #" & txtPWSBD & "# "
If txtPWEB <> "--/--/----" Then where = where & " and calfbirth.birthdate <= #" & txtPWEB & "# "
If txtPWSH <> "--/--/----" Then where = where & " and calfback.findate >= #" & txtPWSH & "# "
If txtPWEH <> "--/--/----" Then where = where & " and calfback.findate <= #" & txtPWEH & "# "
'If chkSex(0).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '0' "
'If chkSex(1).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '1' "
'If chkSex(2).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '2' "
'If chkSex(3).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '3' "
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from PW_Back"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
If chkSex(0).Value = vbChecked Then DB.Execute "insert into pw_back in '" & repfile & "' SELECT DISTINCTROW calfback.CalfID, calfback.recdate, calfback.findate, [calfback]![findate]-[calfback]![recdate] AS Days_On_Feed, calfback.recweight, calfback.intweight, calfback.finweight, ([calfback]![intweight]-[calfback]![recweight])/([calfback]![intdate]-[calfback]![recdate]) AS Interim_ADG, ([calfback]![finweight]-[calfback]![recweight])/([calfback]![findate]-[calfback]![recdate]) AS Final_ADG, calfbirth.birthdate, calfbirth.sireid, calfbirth.sex FROM calfback INNER JOIN calfbirth ON (calfback.CalfID = calfbirth.CalfID) AND (calfback.HerdID = calfbirth.HerdID) Where calfbirth.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '0'"
If chkSex(1).Value = vbChecked Then DB.Execute "insert into pw_back in '" & repfile & "' SELECT DISTINCTROW calfback.CalfID, calfback.recdate, calfback.findate, [calfback]![findate]-[calfback]![recdate] AS Days_On_Feed, calfback.recweight, calfback.intweight, calfback.finweight, ([calfback]![intweight]-[calfback]![recweight])/([calfback]![intdate]-[calfback]![recdate]) AS Interim_ADG, ([calfback]![finweight]-[calfback]![recweight])/([calfback]![findate]-[calfback]![recdate]) AS Final_ADG, calfbirth.birthdate, calfbirth.sireid, calfbirth.sex FROM calfback INNER JOIN calfbirth ON (calfback.CalfID = calfbirth.CalfID) AND (calfback.HerdID = calfbirth.HerdID) Where calfbirth.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '1'"
If chkSex(2).Value = vbChecked Then DB.Execute "insert into pw_back in '" & repfile & "' SELECT DISTINCTROW calfback.CalfID, calfback.recdate, calfback.findate, [calfback]![findate]-[calfback]![recdate] AS Days_On_Feed, calfback.recweight, calfback.intweight, calfback.finweight, ([calfback]![intweight]-[calfback]![recweight])/([calfback]![intdate]-[calfback]![recdate]) AS Interim_ADG, ([calfback]![finweight]-[calfback]![recweight])/([calfback]![findate]-[calfback]![recdate]) AS Final_ADG, calfbirth.birthdate, calfbirth.sireid, calfbirth.sex FROM calfback INNER JOIN calfbirth ON (calfback.CalfID = calfbirth.CalfID) AND (calfback.HerdID = calfbirth.HerdID) Where calfbirth.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '2'"
If chkSex(3).Value = vbChecked Then DB.Execute "insert into pw_back in '" & repfile & "' SELECT DISTINCTROW calfback.CalfID, calfback.recdate, calfback.findate, [calfback]![findate]-[calfback]![recdate] AS Days_On_Feed, calfback.recweight, calfback.intweight, calfback.finweight, ([calfback]![intweight]-[calfback]![recweight])/([calfback]![intdate]-[calfback]![recdate]) AS Interim_ADG, ([calfback]![finweight]-[calfback]![recweight])/([calfback]![findate]-[calfback]![recdate]) AS Final_ADG, calfbirth.birthdate, calfbirth.sireid, calfbirth.sex FROM calfback INNER JOIN calfbirth ON (calfback.CalfID = calfbirth.CalfID) AND (calfback.HerdID = calfbirth.HerdID) Where calfbirth.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '3'"
'SQL = "insert into pw_back in '" & repfile & "' SELECT DISTINCTROW calfback.CalfID, calfback.recdate, calfback.findate, [calfback]![findate]-[calfback]![recdate] AS Days_On_Feed, calfback.recweight, calfback.intweight, calfback.finweight, ([calfback]![intweight]-[calfback]![recweight])/([calfback]![intdate]-[calfback]![recdate]) AS Interim_ADG, ([calfback]![finweight]-[calfback]![recweight])/([calfback]![findate]-[calfback]![recdate]) AS Final_ADG, calfbirth.birthdate, calfbirth.sireid, calfbirth.sex FROM calfback INNER JOIN calfbirth ON (calfback.CalfID = calfbirth.CalfID) AND (calfback.HerdID = calfbirth.HerdID) Where calfbirth.herdid = '" & herdid & "' " & Where & " and calfbirth.sex = '0'"
'DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub
 
Private Sub LoadSexRange()
Dim repdb As database, DB As database
'Set db = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set repdb = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
repdb.Execute ("delete * from yearinclude"), dbFailOnError
If chkSex(0).Value <> vbChecked Then repdb.Execute "delete * from yearlings where sex = '0'"
If chkSex(1).Value <> vbChecked Then repdb.Execute "delete * from yearlings where sex = '1'"
If chkSex(2).Value <> vbChecked Then repdb.Execute "delete * from yearlings where sex = '2'"
If chkSex(3).Value <> vbChecked Then repdb.Execute "delete * from yearlings where sex = '3'"

repdb.Close: Set repdb = Nothing
'db.Close: Set db = Nothing
End Sub

Private Sub BuildPW_Feed()
Dim DB As database, RS As Recordset, SQL$, where$
If txtPWSBD <> "--/--/----" Then where = where & " and calfbirth.birthdate >= #" & txtPWSBD & "# "
If txtPWEB <> "--/--/----" Then where = where & " and calfbirth.birthdate <= #" & txtPWEB & "# "
If txtPWSH <> "--/--/----" Then where = where & " and calffeed.findate >= #" & txtPWSH & "# "
If txtPWEH <> "--/--/----" Then where = where & " and calffeed.findate <= #" & txtPWEH & "# "
'If chkSex(0).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '0' "
'If chkSex(1).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '1' "
'If chkSex(2).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '2' "
'If chkSex(3).Value = vbChecked Then Where = Where & IIf(InStr(1, Where, "calfbirth.sex") > 0, " or ", " and ") & " calfbirth.sex = '3' "
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from PW_feed"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
If chkSex(0).Value = vbChecked Then DB.Execute "insert into pw_feed in '" & repfile & "' SELECT DISTINCTROW calfbirth.birthdate, calffeed.CalfID, calffeed.int1date, calffeed.findate,datediff('d',[calffeed]![int1date],[calffeed]![findate]) AS Days_On_Feed, calffeed.int1wt, calffeed.int2wt, calffeed.finwt, ([calffeed]![int2wt]-[calffeed]![int1wt])/([calffeed]![int2date]-[calffeed]![int1date]) AS Interim_ADG, ([calffeed]![finwt]-[calffeed]![int1wt])/([calffeed]![findate]-[calffeed]![int1date]) AS Final_ADG, calffeed.misc1, calffeed.misc2, calffeed.misc3, calfbirth.sex, calfbirth.sireID FROM (calfbirth INNER JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID)) INNER JOIN calfback ON (calfbirth.CalfID = calfback.CalfID) AND (calfbirth.HerdID = calfback.HerdID) Where calffeed.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '0'"
If chkSex(1).Value = vbChecked Then
   SQL$ = "insert into pw_feed in '" & repfile & "' SELECT DISTINCTROW calfbirth.birthdate, calffeed.CalfID, calffeed.int1date, calffeed.findate, datediff('d',[calffeed]![int1date],[calffeed]![findate]) AS Days_On_Feed, calffeed.int1wt, calffeed.int2wt, calffeed.finwt, ([calffeed]![int2wt]-[calffeed]![int1wt])/([calffeed]![int2date]-[calffeed]![int1date]) AS Interim_ADG, ([calffeed]![finwt]-[calffeed]![int1wt])/([calffeed]![findate]-[calffeed]![int1date]) AS Final_ADG, calffeed.misc1, calffeed.misc2, calffeed.misc3, calfbirth.sex, calfbirth.sireID FROM (calfbirth INNER JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID)) INNER JOIN calfback ON (calfbirth.CalfID = calfback.CalfID) AND (calfbirth.HerdID = calfback.HerdID) Where calffeed.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '1'"
   DB.Execute SQL$
End If
If chkSex(2).Value = vbChecked Then DB.Execute "insert into pw_feed in '" & repfile & "' SELECT DISTINCTROW calfbirth.birthdate, calffeed.CalfID, calffeed.int1date, calffeed.findate, datediff('d',[calffeed]![int1date],[calffeed]![findate]) AS Days_On_Feed, calffeed.int1wt, calffeed.int2wt, calffeed.finwt, ([calffeed]![int2wt]-[calffeed]![int1wt])/([calffeed]![int2date]-[calffeed]![int1date]) AS Interim_ADG, ([calffeed]![finwt]-[calffeed]![int1wt])/([calffeed]![findate]-[calffeed]![int1date]) AS Final_ADG, calffeed.misc1, calffeed.misc2, calffeed.misc3, calfbirth.sex, calfbirth.sireID FROM (calfbirth INNER JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID)) INNER JOIN calfback ON (calfbirth.CalfID = calfback.CalfID) AND (calfbirth.HerdID = calfback.HerdID)  Where calffeed.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '2'"
If chkSex(3).Value = vbChecked Then DB.Execute "insert into pw_feed in '" & repfile & "' SELECT DISTINCTROW calfbirth.birthdate, calffeed.CalfID, calffeed.int1date, calffeed.findate, datediff('d',[calffeed]![int1date],[calffeed]![findate]) AS Days_On_Feed, calffeed.int1wt, calffeed.int2wt, calffeed.finwt, ([calffeed]![int2wt]-[calffeed]![int1wt])/([calffeed]![int2date]-[calffeed]![int1date]) AS Interim_ADG, ([calffeed]![finwt]-[calffeed]![int1wt])/([calffeed]![findate]-[calffeed]![int1date]) AS Final_ADG, calffeed.misc1, calffeed.misc2, calffeed.misc3, calfbirth.sex, calfbirth.sireID FROM (calfbirth INNER JOIN calffeed ON (calfbirth.CalfID = calffeed.CalfID) AND (calfbirth.HerdID = calffeed.HerdID)) INNER JOIN calfback ON (calfbirth.CalfID = calfback.CalfID) AND (calfbirth.HerdID = calfback.HerdID) Where calffeed.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '3'"

DB.Close: Set DB = Nothing

End Sub
 
Private Sub BuildPW_Carc()
On Local Error GoTo ErrHandler
Dim DB As database, RS As Recordset, SQL$, where$
Dim repdb As database
If txtPWSBD <> "--/--/----" Then where = where & " and calfbirth.birthdate >= #" & txtPWSBD & "# "
If txtPWEB <> "--/--/----" Then where = where & " and calfbirth.birthdate <= #" & txtPWEB & "# "
If txtPWSH <> "--/--/----" Then where = where & " and calfcarcass.carcassdate >= #" & txtPWSH & "# "
If txtPWEH <> "--/--/----" Then where = where & " and calfcarcass.carcassdate <= #" & txtPWEH & "# "
Set repdb = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
repdb.Execute "delete * from PW_Carc"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
If chkSex(0).Value = vbChecked Then DB.Execute "insert into pw_carc in '" & repfile & "' SELECT DISTINCTROW calfcarcass.yribeye, calfcarcass.CalfID, calfcarcass.carcassdate, [calfcarcass]![carcassdate]-[calfbirth]![birthdate] AS Age_At_Harvest, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.qgrade, calfcarcass.qscore, iif(calffeed.finwt > 0, (calfcarcass!ywt / calffeed!finwt) * 100, 0) AS Dressing_Percent, calfcarcass.misc1, calfbirth.sireID, calfbirth.CowID, calfbirth.sex FROM (calffeed LEFT JOIN calfcarcass ON (calffeed.CalfID = calfcarcass.CalfID) AND (calffeed.HerdID = calfcarcass.HerdID)) LEFT JOIN calfbirth ON (calffeed.CalfID = calfbirth.CalfID) AND (calffeed.HerdID = calfbirth.HerdID) Where calfcarcass.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '0'"
If chkSex(1).Value = vbChecked Then DB.Execute "insert into pw_carc in '" & repfile & "' SELECT DISTINCTROW calfcarcass.yribeye, calfcarcass.CalfID, calfcarcass.carcassdate, [calfcarcass]![carcassdate]-[calfbirth]![birthdate] AS Age_At_Harvest, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.qgrade, calfcarcass.qscore, iif(calffeed.finwt > 0, (calfcarcass!ywt / calffeed!finwt) * 100, 0) AS Dressing_Percent, calfcarcass.misc1, calfbirth.sireID, calfbirth.CowID, calfbirth.sex FROM (calffeed LEFT JOIN calfcarcass ON (calffeed.CalfID = calfcarcass.CalfID) AND (calffeed.HerdID = calfcarcass.HerdID)) LEFT JOIN calfbirth ON (calffeed.CalfID = calfbirth.CalfID) AND (calffeed.HerdID = calfbirth.HerdID) Where calfcarcass.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '1'"
If chkSex(2).Value = vbChecked Then DB.Execute "insert into pw_carc in '" & repfile & "' SELECT DISTINCTROW calfcarcass.yribeye, calfcarcass.CalfID, calfcarcass.carcassdate, [calfcarcass]![carcassdate]-[calfbirth]![birthdate] AS Age_At_Harvest, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.qgrade, calfcarcass.qscore, iif(calffeed.finwt > 0, (calfcarcass!ywt / calffeed!finwt) * 100, 0) AS Dressing_Percent, calfcarcass.misc1, calfbirth.sireID, calfbirth.CowID, calfbirth.sex FROM (calffeed LEFT JOIN calfcarcass ON (calffeed.CalfID = calfcarcass.CalfID) AND (calffeed.HerdID = calfcarcass.HerdID)) LEFT JOIN calfbirth ON (calffeed.CalfID = calfbirth.CalfID) AND (calffeed.HerdID = calfbirth.HerdID) Where calfcarcass.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '2'"
If chkSex(3).Value = vbChecked Then DB.Execute "insert into pw_carc in '" & repfile & "' SELECT DISTINCTROW calfcarcass.yribeye, calfcarcass.CalfID, calfcarcass.carcassdate, [calfcarcass]![carcassdate]-[calfbirth]![birthdate] AS Age_At_Harvest, calfcarcass.ygrade, calfcarcass.ywt, calfcarcass.yfat, calfcarcass.ykidney, calfcarcass.qgrade, calfcarcass.qscore, iif(calffeed.finwt > 0, (calfcarcass!ywt / calffeed!finwt) * 100, 0) AS Dressing_Percent, calfcarcass.misc1, calfbirth.sireID, calfbirth.CowID, calfbirth.sex FROM (calffeed LEFT JOIN calfcarcass ON (calffeed.CalfID = calfcarcass.CalfID) AND (calffeed.HerdID = calfcarcass.HerdID)) LEFT JOIN calfbirth ON (calffeed.CalfID = calfbirth.CalfID) AND (calffeed.HerdID = calfbirth.HerdID) Where calfcarcass.herdid = '" & herdid & "' " & where & " and calfbirth.sex = '3'"
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
GoSub Build_Carc_Dist_Qual
GoSub Build_Carc_Dist_HCW
GoSub Build_Carc_Dist_YLD
GoSub Build_Carc_Dist_Over



   
   SQL = "update pw_carc SET qgrade = 10.5 where qgrade = 'Prime+'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 9.5 where qgrade = 'Prime'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 8.5 where qgrade = 'Prime-'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 7.5 where qgrade = 'Choice+'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 6.75 where qgrade = 'CAB'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 6.75 where qgrade = 'STS'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 6.5 where qgrade = 'Choice'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 5.5 where qgrade = 'Choice-'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 5.5 where qgrade = 'AAA'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 4.75 where qgrade = 'Select+'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 4.5 where qgrade = 'Select'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 4.5 where qgrade = 'AA'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 4.25 where qgrade = 'Select-'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 3.5 where qgrade = 'Standard+'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 3.5 where qgrade = 'A'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = '3.0' where qgrade = 'Standard'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 2.5 where qgrade = 'Standard-'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = '1.0' where qgrade = 'B1'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'HRI'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'NoRoll'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'B2'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'B3'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'B4'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'D1'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'D2'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'D3'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'D4'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'C'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'Dark'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'Stag'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'Comm'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = 'Other'"
   repdb.Execute SQL
   
   SQL = "update pw_carc SET qgrade = 0 where qgrade = ''"
   repdb.Execute SQL
   
   'SQL = "update lpr_carc SET qgrade = 0 where qgrade = null "
   'repdb.Execute SQL
repdb.Close: Set repdb = Nothing
DB.Close: Set DB = Nothing
Exit Sub
ErrHandler:
TEXT(2) = Erl
GMODNAME$ = Me.Name & " - BuildPW_Carc "
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())

Build_Carc_Dist_Qual:
TEXT(1) = "Build_Carc_Dist_Qual"
DB.Execute "delete * from Carc_Dist_Qual"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '10.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Prime+' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '9.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Prime' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '8.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Prime-' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '7.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Choice+' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '6.75' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='CAB' or PW_Carc.qgrade='STS' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '6.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Choice' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '5.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Choice-' or PW_Carc.qgrade='AAA' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '4.75' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Select+' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '4.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Select' or PW_Carc.qgrade='AA' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '4.25' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Select-' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '3.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Standard+' or PW_Carc.qgrade='A' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '3.0' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Standard' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '2.5' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Standard-' GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, '1.0' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='B1' GROUP BY PW_Carc.sex"

'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'HR1' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='HR1' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'NoRoll' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='NoRoll' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'B2' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='B2' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'B3' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='B3' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'B4' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='B4' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'D1' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='D1' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'D2' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='D2' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'D3' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='D3' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'D4' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='D4' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'C' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='C' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'Dark' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Dark' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'Stag' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Stag' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'Comm' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Comm' GROUP BY PW_Carc.sex"
'DB.Execute "insert into Carc_Dist_Qual SELECT DISTINCTROW PW_Carc.sex, 'Other' as QualityGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc WHERE PW_Carc.qgrade='Other' GROUP BY PW_Carc.sex"

Return

Build_Carc_Dist_HCW:
TEXT(1) = "Build_Carc_Dist_HCW"
DB.Execute "delete * from Carc_Dist_HCW"
DB.Execute "insert into Carc_Dist_HCW SELECT DISTINCTROW PW_Carc.sex, '<= 550' as HCW, Count(PW_Carc.ywt) AS [Number] From PW_Carc Where PW_Carc.ywt <= 550 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_HCW SELECT DISTINCTROW PW_Carc.sex, '551 - 650' as HCW, Count(PW_Carc.ywt) AS [Number] From PW_Carc Where PW_Carc.ywt > 550 and PW_Carc.ywt <= 650 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_HCW SELECT DISTINCTROW PW_Carc.sex, '651 - 750' as HCW, Count(PW_Carc.ywt) AS [Number] From PW_Carc Where PW_Carc.ywt > 650 and PW_Carc.ywt <= 750 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_HCW SELECT DISTINCTROW PW_Carc.sex, '751- 850' as HCW, Count(PW_Carc.ywt) AS [Number] From PW_Carc Where PW_Carc.ywt > 750 and PW_Carc.ywt <= 850 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_HCW SELECT DISTINCTROW PW_Carc.sex, '> 850' as HCW, Count(PW_Carc.ywt) AS [Number] From PW_Carc Where PW_Carc.ywt > 850 GROUP BY PW_Carc.sex"
Return

Build_Carc_Dist_YLD:
TEXT(1) = "Build_Carc_Dist_YLD"
DB.Execute "delete * from Carc_Dist_YLD"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '1' as YLDGrade, Count(PW_Carc.ygrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 1 and PW_Carc.ygrade <2 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '2a' as YLDGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 2.00 and PW_Carc.ygrade <2.5 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '2b' as YLDGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 2.50 and PW_Carc.ygrade <3 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '3a' as YLDGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 3.00 and PW_Carc.ygrade <3.5 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '3b' as YLDGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 3.50 and PW_Carc.ygrade <4 GROUP BY PW_Carc.sex"
DB.Execute "insert into Carc_Dist_YLD SELECT DISTINCTROW PW_Carc.sex, '4+' as YLDGrade, Count(PW_Carc.qgrade) AS [Number] From PW_Carc Where PW_Carc.ygrade >= 4 GROUP BY PW_Carc.sex"
Return

Build_Carc_Dist_Over:
TEXT(1) = "Build_Carc_Dist_Over"
DB.Execute "delete * from Carc_Dist_Over"
DB.Execute "INSERT INTO Carc_Dist_Over ( TypeName, Type, [Number] ) SELECT DISTINCTROW Carc_Dist_HCW.HCW AS TypeName, 'HCW' AS Type, Sum(Carc_Dist_HCW.Number) AS [Number] From Carc_Dist_HCW GROUP BY Carc_Dist_HCW.HCW, 'HCW'"
DB.Execute "INSERT INTO Carc_Dist_Over ( Type, TypeName, [Number] ) SELECT DISTINCTROW 'QGrade' AS Type, Carc_Dist_Qual.QualityGrade AS TypeName, Sum(Carc_Dist_Qual.Number) AS CountOfNumber From Carc_Dist_Qual group By 'QGrade', Carc_Dist_Qual.QualityGrade"
DB.Execute "INSERT INTO Carc_Dist_Over ( Type, TypeName, [Number] ) SELECT DISTINCTROW 'YLD' AS Type, Carc_Dist_YLD.YLDGrade, Sum(Carc_Dist_YLD.Number) From Carc_Dist_YLD group By 'YLD', Carc_Dist_YLD.YLDGrade"
Return
End Sub

Private Function SPAData() As Boolean
Dim mDB As database, mRS As Recordset
SPAData = False
Set mDB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set mRS = mDB.OpenRecordset("select * from prefspa", dbOpenSnapshot)
If mRS.RecordCount > 0 Then SPAData = True
mRS.Close: Set mRS = Nothing
mDB.Close: Set mDB = Nothing
End Function
 
Private Function CSFData() As Boolean
Dim mDB As database, mRS As Recordset
CSFData = False
Set mDB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set mRS = mDB.OpenRecordset("select * from prefcsf", dbOpenSnapshot)
If mRS.RecordCount > 0 Then CSFData = True
mRS.Close: Set mRS = Nothing
mDB.Close: Set mDB = Nothing
End Function

 
Private Sub cmdselectvend_Click()
'Dim RS As Recordset
'FrmSelect_Multi_Herds.Show vbModal
' If FrmSelect_Multi_Herds!lstherd.SelectedCount > 0 Then
'   lblhow_many_herd.Caption = Trim$(Str$(FrmSelect_Multi_Herds!lstherd.SelectedCount))
'   FrmSelect_Multi_Herds!lstherd.Col = 0
'   'herdid$ = FrmSelect_Multi_Herds!lstherd.ColText
'   chkOverwrite.Enabled = True
'   Call load_year
'Else
'   chkOverwrite.Enabled = False
'   lblhow_many_herd.Caption = "All"
' End If
End Sub

Private Sub Form_Load()
Dim RS As DAO.Recordset, DB As DAO.database
Call centermdiform(Me, mdimain, 0, 0)
OptSort(0).Value = True
 Load FrmSelect_Multi_Herds
 
 reports$(1) = "Calf Reference List"
 reports(2) = "Calf Reference List -- Birth"
 reports(3) = "Calf Reference List -- Weaning"
 reports(4) = "Calf Reference List -- Backgrounding"
 reports(5) = "Calf Reference List -- Replacement"
 reports(6) = "Calf Reference List -- Feedlot"
 reports(7) = "Calf Reference List -- Carcass"
 reports$(8) = "Herd Analysis"
 reports$(9) = "Post Weaning -- Background"
 reports$(10) = "Post Weaning -- Yearling"
 reports$(11) = "Post Weaning -- Feedlot"
 reports$(12) = "Post Weaning -- Carcass"
 reports(13) = "Calf Notes"
 For t = 1 To hmreps%
  lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optpreview.Value = True
optCont(1).Value = True
Call DisableFrames
Call load_year
optOrder(0).Value = True
chkOverwrite.Enabled = True
txtPWSBD.TEXT = Format(Now, "mm/dd/yyyy")
txtPWEB.TEXT = Format(Now, "mm/dd/yyyy")
txtPWSH.TEXT = Format(Now, "mm/dd/yyyy")
txtPWEH.TEXT = Format(Now, "mm/dd/yyyy")
txtStartDate.TEXT = Format(Now, "mm/dd/yyyy")
txtEndDate.TEXT = Format(Now, "mm/dd/yyyy")
txtWeighDate.TEXT = Format(Now, "mm/dd/yyyy")
Screen.MousePointer = vbDefault
End Sub

Private Sub DisableFrames()
 
End Sub

Private Sub EnableFrames()
 'chkOverwrite.Enabled = True
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set calfreps = Nothing
Unload FrmSelect_Multi_Herds
Set FrmSelect_Multi_Herds = Nothing
End Sub

Private Sub lstreports_Click()
'Label4.Caption = "Start Carcass Date"
'Label3.Caption = "End Carcass Date"
fraDate.Visible = False
FraPostWean.Visible = False
FraPW.Visible = False
FraSort.Visible = False
FraPW.Visible = False
FraHarvestDates.Visible = False
Label3.Caption = "Start Harvest Date"
Label4.Caption = "End Harvest Date"
Select Case lstreports.ListIndex
   Case 0
      '
   Case Is < 7
      FraPostWean.Visible = True
      FraHarvestDates.Visible = False
   Case 7
      fraDate.Visible = True
      FraSort.Visible = True
   Case Is > 7
      FraSort.Visible = False
      FraPW.Visible = True
      FraPostWean.Visible = True
      FraHarvestDates.Visible = True
      Label4.Caption = "Start Weigh Date"
      Label3.Caption = "End Weigh Date"
      Select Case lstreports.ListIndex
         Case 8
            optOrder(0).Caption = "ID"
            optOrder(1).Caption = "DOF"
            optOrder(2).Caption = "Interim ADG"
            optOrder(3).Caption = "Final ADG"
            optOrder(4).Caption = "Interim Wt"
            optOrder(5).Caption = "Final Wt"
            optOrder(6).Visible = False
         Case 9
            optOrder(0).Caption = "365 Wt"
            optOrder(1).Caption = "ID"
            optOrder(2).Caption = "Actual Wt"
            optOrder(3).Caption = "WDA"
            optOrder(4).Caption = "Fr Score"
            optOrder(5).Caption = "ADG"
            optOrder(6).Visible = False
         Case 10
            optOrder(0).Caption = "ID"
            optOrder(1).Caption = "DOF"
            optOrder(2).Caption = "Interim ADG"
            optOrder(3).Caption = "Final ADG"
            optOrder(4).Caption = "Interim Wt"
            optOrder(5).Caption = "Final Wt"
            optOrder(6).Visible = False
         Case 11
            Label4.Caption = "Start Carcass Date"
            Label3.Caption = "End Carcass Date"
            optOrder(6).Visible = True
            optOrder(0).Caption = "ID"
            optOrder(1).Caption = "Yld Grade"
            optOrder(2).Caption = "Quality Grade"
            optOrder(3).Caption = "HCW"
            optOrder(4).Caption = "Fat"
            optOrder(5).Caption = "REA"
            optOrder(6).Caption = "Dressing %"
         Case 12
            FraSort.Visible = False
            FraPW.Visible = False
            FraHarvestDates.Visible = False
      End Select
End Select
End Sub
Private Sub cmdCal_Click()
 gcaldate = txtStartDate.TEXT
 Call GetDate(gcaldate)
 txtStartDate.TEXT = gcaldate

End Sub

Private Sub cmdCal1_Click()
 gcaldate = txtEndDate.TEXT
 Call GetDate(gcaldate)
 txtEndDate.TEXT = gcaldate

End Sub

Private Sub optCont_Click(Index As Integer)
If Index = 0 Then group = True Else group = False
End Sub

Private Sub SSCommand1_Click()
gcaldate = txtTurnDate.TEXT
 Call GetDate(gcaldate)
 txtTurnDate.TEXT = gcaldate
End Sub

Private Sub SSCommand2_Click()
gcaldate = txtWeighDate.TEXT
 Call GetDate(gcaldate)
 txtWeighDate.TEXT = gcaldate
End Sub


Private Sub SSCommand3_Click()
 gcaldate = txtPWSBD.TEXT
 Call GetDate(gcaldate)
 txtPWSBD.TEXT = gcaldate
End Sub

Private Sub SSCommand4_Click()
gcaldate = txtPWEB.TEXT
 Call GetDate(gcaldate)
 txtPWEB.TEXT = gcaldate
End Sub


Private Sub SSCommand5_Click()
gcaldate = txtPWSH.TEXT
 Call GetDate(gcaldate)
 txtPWSH.TEXT = gcaldate
End Sub


Private Sub SSCommand6_Click()
gcaldate = txtPWEH.TEXT
 Call GetDate(gcaldate)
 txtPWEH.TEXT = gcaldate
End Sub


Private Sub txtTurnDate_Change()
If txtTurnDate.TEXT = "--/--/----" Then chkOverwrite.Enabled = False Else chkOverwrite.Enabled = True
End Sub


