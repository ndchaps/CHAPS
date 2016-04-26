VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form YearlingsReps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yearling Reports"
   ClientHeight    =   4875
   ClientLeft      =   4050
   ClientTop       =   1470
   ClientWidth     =   6840
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4875
   ScaleWidth      =   6840
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
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
   Begin VB.Frame FraPostWeanSex 
      Height          =   975
      Left            =   3720
      TabIndex        =   29
      Top             =   2280
      Width           =   3015
      Begin VB.CheckBox chkSex 
         Caption         =   "Steers"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Heifers"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Bulls"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Misc"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraPostWean 
      Height          =   2175
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      Begin Threed.SSCommand cmdCal 
         Height          =   270
         Left            =   2685
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "yreps.frx":0000
      End
      Begin MSMask.MaskEdBox txtStartDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   270
         Left            =   2685
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "yreps.frx":04BE
      End
      Begin MSMask.MaskEdBox txtEndDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   270
         Left            =   2685
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1350
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "yreps.frx":097C
      End
      Begin MSMask.MaskEdBox txtStartDayDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   1320
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   270
         Left            =   2685
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1710
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "yreps.frx":0E3A
      End
      Begin MSMask.MaskEdBox txtEndDayDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   1680
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "End Harvest Date"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Harvest Date"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Birthdate"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "End Birthdate"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   1095
      Left            =   1200
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
      Begin VB.OptionButton optpreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optprint 
         Caption         =   "Print"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame FraPostWeanSort 
      Height          =   1095
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
      Begin VB.OptionButton optOrder 
         Caption         =   "365 Wt"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "ID"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Birth Wt"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "ADG"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "WDA"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Fr Score"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdselectvend 
      Caption         =   "&Select"
      Height          =   385
      Left            =   2400
      TabIndex        =   6
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   5760
      TabIndex        =   2
      Top             =   4425
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   4680
      TabIndex        =   0
      Top             =   4425
      Width           =   1000
   End
   Begin VB.Label lblhmvend 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many Herds"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Label lblhow_many_herd 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
End
Attribute VB_Name = "YearlingsReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 10
Dim t As Integer
Dim reports(hmreps%) As String


Private Sub cmdcancel_Click()
 Unload Me
End Sub


Private Sub cmdok_Click()
Dim Order$, TITLE$, Herds$, title1$, title2$, Title3$, Title4$
 
 Screen.MousePointer = 11
 report.Initialize ' init the class
 If optprint Then report.SetDestination = 1
 If chkSex(0).Value = False And chkSex(1).Value = False And chkSex(2).Value = False And chkSex(3).Value = False Then Screen.MousePointer = vbDefault: MsgBox "Please Select A Sex To Run This Report", vbOKOnly: Exit Sub
 Call LoadSexRange
 Select Case lstreports.TEXT
  Case reports$(1)
   If optOrder(0).Value Then Order$ = "Wgt365"
   If optOrder(1).Value Then Order$ = "yearid"
   If optOrder(2).Value Then Order$ = "birthwt"
   If optOrder(4).Value Then Order$ = "WDAOff"
   If optOrder(3).Value Then Order$ = "ADGOn"
   If optOrder(5).Value Then Order$ = "FrScore"
   Call Create365Reports(Order$, txtStartDate.TEXT, txtEndDate.TEXT, txtEndDayDate.TEXT, txtStartDayDate.TEXT)
   If optOrder(0).Value Then title1$ = reports$(1) & " Order By 365 Weight"
   If optOrder(1).Value Then title1$ = reports$(1) & " Order By Yearling ID"
   If optOrder(2).Value Then title1$ = reports$(1) & " Order By Birth Weight"
   If optOrder(4).Value Then title1$ = reports$(1) & " Order By WDA Off Test"
   If optOrder(3).Value Then title1$ = reports$(1) & " Order By ADG On Test"
   If optOrder(5).Value Then title1$ = reports$(1) & " Order By Frame Score"
   report.SetReportFileName = dbdir$ & "\" & "yrreps.rpt"
   title2$ = "Birthdate Range:" & IIf(IsDate(txtStartDate.TEXT), CStr(txtStartDate.TEXT), "--/--/----") & " -- " & IIf(IsDate(txtEndDate.TEXT), CStr(txtEndDate.TEXT), "--/--/----")
   'Title3$ = IIf(IsDate(txtWeighDate.TEXT), "Weigh Date: " & CStr(txtWeighDate.TEXT), "Weigh Date:  --/--/----")
   report.Setcommonformulas(title1$, title2$, Title3$) = Title4$
   report.setDbname = repfile$
   'report.SetReportCaption = reports$(1)
 End Select
  report.PrintReport
   Screen.MousePointer = 0
End Sub
 
Private Sub LoadSexRange()
Dim repdb As database, db As database
Set db = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set repdb = DBEngine(0).OpenDatabase(repfile$, False, False)
repdb.Execute ("delete * from yearinclude"), dbFailOnError
If chkSex(0) Then db.Execute ("insert into yearinclude in '" & repfile$ & "' select herdid as includeherdid, calfid as includecalfid from calfbirth where sex = '0'"), dbFailOnError
If chkSex(1) Then db.Execute ("insert into yearinclude in '" & repfile$ & "' select herdid as includeherdid, calfid as includecalfid from calfbirth where sex = '1'"), dbFailOnError
If chkSex(2) Then db.Execute ("insert into yearinclude in '" & repfile$ & "' select herdid as includeherdid, calfid as includecalfid from calfbirth where sex = '2'"), dbFailOnError
If chkSex(3) Then db.Execute ("insert into yearinclude in '" & repfile$ & "' select herdid as includeherdid, calfid as includecalfid from calfbirth where sex = '3'"), dbFailOnError
repdb.Close: Set repdb = Nothing
db.Close: Set db = Nothing
End Sub

 
Private Sub cmdselectvend_Click()
 FrmSelect_Multi_Herds.Show vbModal
 If FrmSelect_Multi_Herds!lstherd.SelectedCount > 0 Then
   lblhow_many_herd.Caption = Trim$(Str$(FrmSelect_Multi_Herds!lstherd.SelectedCount))
  Else
   lblhow_many_herd.Caption = "All"
 End If
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "Yearling Reports"
 For t = 1 To hmreps%
  lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optOrder(0).Value = True
 optpreview.Value = True
End Sub



Private Sub cmdCal_Click()
 gcaldate = txtStartDate.TEXT
 Call GetDate(gcaldate)
 txtStartDate.TEXT = gcaldate
End Sub



Private Sub SSCommand1_Click()
gcaldate = txtEndDate.TEXT
 Call GetDate(gcaldate)
 txtEndDate.TEXT = gcaldate
End Sub


Private Sub SSCommand2_Click()
'gcaldate = txtWeighDate.TEXT
 Call GetDate(gcaldate)
 'txtWeighDate.TEXT = gcaldate
End Sub


Private Sub SSCommand3_Click()
gcaldate = txtStartDayDate.TEXT
 Call GetDate(gcaldate)
 txtStartDayDate.TEXT = gcaldate
End Sub

Private Sub SSCommand4_Click()
gcaldate = txtEndDayDate.TEXT
 Call GetDate(gcaldate)
 txtEndDayDate.TEXT = gcaldate
End Sub

