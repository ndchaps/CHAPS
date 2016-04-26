VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Input_forms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Forms"
   ClientHeight    =   4230
   ClientLeft      =   1065
   ClientTop       =   1395
   ClientWidth     =   7170
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   7170
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3150
      _Version        =   65536
      _ExtentX        =   5556
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
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Herd"
      Height          =   360
      Left            =   1830
      TabIndex        =   12
      Top             =   1935
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   3150
      Left            =   3735
      TabIndex        =   6
      Top             =   225
      Width           =   3270
      Begin VB.Frame Frame3 
         Height          =   885
         Left            =   165
         TabIndex        =   13
         Top             =   1725
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton OptSort 
            Caption         =   "Sex Code"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   17
            Top             =   525
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Calf ID"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Sire ID"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   525
            Width           =   1095
         End
         Begin VB.OptionButton OptSort 
            Caption         =   "Cow ID"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkdata 
         Alignment       =   1  'Right Justify
         Caption         =   "Include Data"
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   1410
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Dtestart 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   705
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin MSMask.MaskEdBox dteend 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1035
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label lblstart 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Birthdate"
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label lblend 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Birthdate"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1065
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   2430
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
      Left            =   1920
      TabIndex        =   2
      Top             =   3585
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   450
      TabIndex        =   0
      Top             =   3585
      Width           =   1000
   End
End
Attribute VB_Name = "Input_forms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 6

Dim t As Integer
Dim reports(hmreps%) As String


Private Sub create_REPLACEMENT_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn, readonlyyn)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, exclusiveyn, readonlyyn)
 ' On Error Resume Next
    DBREP.Execute ("delete * from replacement")
  DBREP.Close
  SQL$ = "insert into replacement in '" & repfile$ & "' "
  SQL$ = SQL$ & " SELECT DISTINCTROW calfbirth.CalfID, calfbirth.sireID, sireprof.breed, calfbirth.HerdID"
  'SQL$ = SQL$ & " FROM calfbirth INNER JOIN sireprof ON (sireprof.SireID = calfbirth.sireID) AND (calfbirth.HerdID = sireprof.HerdID) "
  SQL$ = SQL$ & " FROM (sireprof INNER JOIN calfbirth ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID)) INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)"
  SQL$ = SQL$ & " where calfbirth.birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "# "
  'hmfields% = 1
  'Col(1) = 0
  'f'ieldvar$(1) = "calfbirth.herdid"
  'If lblhow_many_herd <> "All" Then
  '   Call create_sql_selection(FrmSelect_Multi_Herds!lstherd, Col(), fieldvar$(), hmfields%, formula1$)
  '   SQL$ = SQL$ & IIf(formula1 = "", "", " and " & formula1$)
  'End If
  SQL$ = SQL$ & " and managecode <> 'A'  and managecode <> 'B'  and managecode <> 'C'  and managecode <> 'D' "
  SQL$ = SQL$ & " and calfbirth.sex = '2' and sireprof.herdid = '" & herdid & "' order by calfbirth.calfid "
'  MsgBox sql$
  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
  report.Setformulas("rundate") = "'From " & Dtestart.TEXT & " to " & dteend.TEXT & "'"
End Sub


Private Sub create_hefREPLACE_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
    DBREP.Execute ("delete * from hefreplacement")
  DBREP.Close
  SQL$ = "insert into hefreplacement in '" & repfile$ & "' "
  SQL = SQL & " SELECT DISTINCTROW calfbirth.CalfID, calfbirth.sireID, IIf(isnull(CALFBIRTH.breed),' ',calfbirth.breed) AS breed, calfbirth.HerdID, calfbirth.CowID, calfwean.wt205, calfwean.ratio, calfwean.actweight, cowprof.mpda AS MPPA FROM sireprof RIGHT JOIN (cowprof LEFT JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (cowprof.HerdID = calfbirth.HerdID) AND (cowprof.cowID = calfbirth.CowID)) ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID) "
  SQL$ = SQL$ & " where calfbirth.sex = '2' and cowprof.herdid = '" & herdid & "'  and  calfbirth.birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "# "
  SQL$ = SQL$ & " and managecode <> 'A'  and managecode <> 'B'  and managecode <> 'C'  and managecode <> 'D' "
  SQL$ = SQL$ & " order by calfbirth.calfid "
  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
  report.Setformulas("rundate") = "'From " & Dtestart.TEXT & " to " & dteend.TEXT & "'"
End Sub

Private Sub create_Allcalf_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from hefreplacement")
  DBREP.Close
  SQL$ = "insert into hefreplacement in '" & repfile$ & "' "
  SQL = SQL & " SELECT DISTINCTROW calfbirth.CalfID, calfbirth.sireID, IIf(isnull(CALFBIRTH.breed),' ',calfbirth.breed) AS breed, calfbirth.HerdID, calfbirth.CowID, calfwean.wt205, calfwean.ratio, calfwean.actweight, cowprof.mpda AS MPPA FROM sireprof RIGHT JOIN (cowprof LEFT JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (cowprof.HerdID = calfbirth.HerdID) AND (cowprof.cowID = calfbirth.CowID)) ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID) "
  SQL$ = SQL$ & " where cowprof.herdid = '" & herdid & "'  and  calfbirth.birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "#  "
  
  If OptSort(0).Value Then SQL$ = SQL$ & " order by  calfbirth.CalfID "
  If OptSort(1).Value Then SQL$ = SQL$ & " order by  calfbirth.cowID "
  If OptSort(2).Value Then SQL$ = SQL$ & " order by  calfbirth.sireID "
  If OptSort(3).Value Then SQL$ = SQL$ & " order by  calfbirth.sex "

  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
  report.Setformulas("rundate") = "'From " & Dtestart.TEXT & " to " & dteend.TEXT & "'"
End Sub



Private Sub create_cowwork_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database, where$
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from cowwrk")
  DBREP.Close
  SQL$ = "insert into cowwrk in '" & repfile$ & "' "
  'SQL$ = SQL$ & " SELECT DISTINCTROW cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire "
  'SQL$ = SQL$ & " FROM cowprof "
  SQL$ = SQL$ & " SELECT DISTINCTROW cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof." & _
    "sire, cowprof.mpda as mppa, IIf(Max([calfbirth].[CowAge]) Is Null,Year(Now()) - Year([cowprof].[birthdate]),Max([calfbirth].[cowage])) AS CowAge FROM cowprof LEFT JOIN calfbirth ON cowprof.cowID = ca" & _
    "lfbirth.CowID "
  
  SQL = SQL & " where cowprof.active = 'A' and cowprof.herdid = '" & herdid & "' GROUP BY cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire, cowprof.mpda " & sortcows
  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
End Sub


Private Sub create_inputwo_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$, pWhere$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from inputwithout")
  DBREP.Close
  SQL$ = "insert into inputwithout in '" & repfile$ & "' "
  SQL$ = SQL$ & " SELECT DISTINCTROW cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire, IIf(Max([calfbirth].[cowage]) Is Null,Year(Now())-Year([cowprof].[birthdate]),Max([calfbirth].[cowage])) + 1 AS cow_age "
  SQL$ = SQL$ & " FROM sireprof RIGHT JOIN (cowprof LEFT JOIN (calfbirth LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)) ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID) "
  SQL$ = SQL & " where cowprof.herdid = '" & herdid & "' and cowprof.active = 'A' "
  SQL = SQL & " GROUP BY cowprof.HerdID, cowprof.cowID, cowprof.birthdate, cowprof.breed, cowprof.sire, cowprof.active " & sortcows
  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
End Sub

Private Sub create_input_report()
  Dim SQL$, hmfields%, Col(1), fieldvar$(1), formula1$
  Dim DB As database
  Dim DBREP As database
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from calfinput")
  DBREP.Close
  SQL$ = "insert into calfinput in '" & repfile$ & "' "
  SQL$ = SQL$ & " SELECT DISTINCTROW calfbirth.CalfID, calfbirth.sireID, calfbirth.CowID, cowprof.birthdate as cowprof_birthdate, cowprof.breed AS cowprof_breed, sireprof.breed AS sireprof_breed, calfbirth.birthdate as calfbirth_birthdate, calfbirth.birthwt, calfbirth.calvingease, calfbirth.sex, calfwean.actweight, calfwean.dateweighed, calfwean.managecode, calfwean.cframe, calfwean.group, calfwean.grade, cowprof.sire, calfbirth.HerdID, calfbirth.cowage as cow_age "
  SQL$ = SQL$ & " FROM ((calfbirth INNER JOIN cowprof ON (cowprof.cowID = calfbirth.CowID) AND (calfbirth.HerdID = cowprof.HerdID)) INNER JOIN sireprof ON (sireprof.SireID = calfbirth.sireID) AND (calfbirth.HerdID = sireprof.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)"
  'hmfields% = 1
  'Col(1) = 0
  'Fieldvar$(1) = "calfbirth.herdid"
  'If lblhow_many_herd <> "All" Then
  '   Call create_sql_selection(FrmSelect_Multi_Herds!lstherd, Col(), fieldvar$(), hmfields%, formula1$)
  '   SQL$ = SQL$ & " Where " & formula1$
  'End If
  SQL$ = SQL$ & " where cowprof.herdid = '" & herdid & "' and cowprof.active = 'A' "
'  MsgBox SQL$
  If Dtestart.TEXT <> "--/--/----" Then SQL = SQL & " and calfbirth.birthdate >= #" & Dtestart.TEXT & "# "
  If dteend.TEXT <> "--/--/----" Then SQL = SQL & " and calfbirth.birthdate <= #" & dteend.TEXT & "# "
  SQL = SQL & sortcows
  DB.Execute (SQL$)
  DB.Close: Set DB = Nothing
End Sub

Private Sub chkdata_Click()
If lstreports.ListIndex <> 0 Then Exit Sub
If chkdata.Value = vbChecked Then
   Dtestart.Enabled = True
   dteend.Enabled = True
   lblStart.Enabled = True
   lblEnd.Enabled = True
Else
   Dtestart.Enabled = False
   dteend.Enabled = False
   lblStart.Enabled = False
   lblEnd.Enabled = False
End If
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
selherd_List.Show vbModal
 If selherd_List.Tag = "CANCEL" Then Exit Sub
 herdid$ = selherd_List.Tag
Unload selherd_List
End Sub

Private Sub CMDOk_Click()
 Screen.MousePointer = vbHourglass
 report.Initialize ' init the class
 If optprint Then report.SetDestination = 1
 Select Case lstreports.TEXT
  Case reports$(1)
    If chkdata.Value = vbChecked Then
      Call create_input_report
      report.SetReportFileName = dbdir$ & "\" & "calfinp.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(1)
    Else
      Call create_inputwo_report
      report.SetReportFileName = dbdir$ & "\" & "calfinp2.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(1)
    End If
  Case reports$(2)
   Call create_cowwork_report
   report.SetReportFileName = dbdir$ & "\" & "cowwrkst.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = reports$(2)
   report.Setorientation = True
  Case reports$(3)
   Call create_REPLACEMENT_report
   report.SetReportFileName = dbdir$ & "\" & "RPLWKSHT.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = reports$(3)
   report.Setorientation = False
  Case reports$(4)
   Call create_hefREPLACE_report
   report.SetReportFileName = dbdir$ & "\" & "hefRPL.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = reports$(4)
   report.Setorientation = True
  Case reports$(5)
   Call create_Allcalf_report
   report.SetReportFileName = dbdir$ & "\" & "allCalf.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = reports$(5)
   'report.Setorientation = True
  Case reports$(6)
   Call create_Allcalf_report
   report.SetReportFileName = dbdir$ & "\" & "allCalf2.rpt"
   report.setDbname = repfile$
   report.SetReportCaption = reports$(5)
   'report.Setorientation = True
   
 End Select
  report.PrintReport
 Screen.MousePointer = vbDefault
End Sub
 
Private Sub cmdselectvend_Click()
 'FrmSelect_Multi_Herds.Show vbModal
 'If FrmSelect_Multi_Herds!lstherd.SelectedCount > 0 Then
 '  lblhow_many_herd.Caption = Trim$(Str$(FrmSelect_Multi_Herds!lstherd.SelectedCount))
 ' Else
 '  lblhow_many_herd.Caption = "All"
 'End If
End Sub

Private Sub Form_Load()
 Dim tmpfile$, tmp%
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "Blank Input Forms - Calves"
 reports$(2) = "Cow Worksheet"
 reports$(3) = "Replacement Worksheet (Portrait)"
 reports$(4) = "Replacement Worksheet (LandScape)"
 reports$(5) = "All Calf Worksheet"
 reports$(6) = "All Calf Input Form"
 
 For t = 1 To hmreps%
  lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optpreview.Value = True
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "Start date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then Dtestart.TEXT = Left(tmpfile$, tmp%)
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "End date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then dteend.TEXT = Left(tmpfile$, tmp%)
' lblhow_many_herd.Caption = "1"
End Sub





Private Sub Form_Unload(Cancel As Integer)
 Dim tmp%
 tmp% = WritePrivateProfileString("chaps", "Start date", Dtestart.TEXT, "chaps.ini")
 tmp% = WritePrivateProfileString("chaps", "End date", dteend.TEXT, "chaps.ini")

End Sub


Private Sub lstreports_Click()
  Call Display_Criteria
End Sub
Private Sub Display_Criteria()
 Select Case lstreports.TEXT
  Case reports$(1)
      GoSub disable_stuff
      chkdata.Enabled = True
  Case reports$(2)
      GoSub disable_stuff
  Case reports$(3)
      GoSub disable_stuff
      Dtestart.Enabled = True
      dteend.Enabled = True
      lblStart.Enabled = True
      lblEnd.Enabled = True
   Case reports$(4)
      GoSub disable_stuff
      Dtestart.Enabled = True
      dteend.Enabled = True
      lblStart.Enabled = True
      lblEnd.Enabled = True
   Case reports$(5)
      GoSub disable_stuff
      Dtestart.Enabled = True
      dteend.Enabled = True
      lblStart.Enabled = True
      lblEnd.Enabled = True
      Frame3.Visible = True
   Case reports$(6)
      GoSub disable_stuff
      Dtestart.Enabled = True
      dteend.Enabled = True
      lblStart.Enabled = True
      lblEnd.Enabled = True
      Frame3.Visible = False
 End Select
 Exit Sub
 
disable_stuff:
   Dtestart.Enabled = False
   dteend.Enabled = False
   lblStart.Enabled = False
   lblEnd.Enabled = False
   chkdata.Enabled = False
   Frame3.Visible = False
   Return
End Sub


