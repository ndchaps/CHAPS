VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form herdreps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herd Reports"
   ClientHeight    =   3630
   ClientLeft      =   720
   ClientTop       =   1455
   ClientWidth     =   3435
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   3435
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
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
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   1000
   End
End
Attribute VB_Name = "herdreps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 10

Dim t As Integer
Dim reports(hmreps%) As String


Private Sub create_refLIST_report()
  Dim SQL$
  Dim db As database
  Dim DBREP As database
  Set db = DBEngine(0).OpenDatabase(dbfile$, False, False)
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from herdref")
  DBREP.Close
  SQL$ = "insert into herdref in '" & repfile$ & "' select * from herd"
  db.Execute (SQL$)
  db.Close
  report.Initialize ' init the class
  If optprint Then report.SetDestination = 1
  report.SetReportFileName = dbdir$ & "hErdref.rpt"
  report.setDbname = repfile$
  report.PrintReport
End Sub


Private Sub cmdcancel_Click()
 Unload Me
End Sub


Private Sub cmdok_Click()
 Dim reportcaption$, reportname$ ', FORMULA$
 Select Case lstreports.TEXT
  Case reports$(1)
   Call create_refLIST_report
   reportcaption$ = reports$(1)
   reportname$ = "herdREF.rpt"
 End Select
End Sub
 
Private Sub form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "Herd Reference List"
 For t = 1 To hmreps%
  lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optpreview.Value = True
End Sub





