VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form usdareps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrative Reports"
   ClientHeight    =   6525
   ClientLeft      =   1605
   ClientTop       =   885
   ClientWidth     =   7425
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   7425
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
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
   Begin VB.CommandButton cmdCowSelect 
      Caption         =   "S&elect"
      Height          =   345
      Left            =   5760
      TabIndex        =   33
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Frame FraEID 
      Caption         =   "EID Range"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   2175
      Begin VB.TextBox TxtLast 
         Height          =   285
         Left            =   480
         MaxLength       =   15
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtFirst 
         Height          =   285
         Left            =   480
         MaxLength       =   15
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblLast 
         Alignment       =   1  'Right Justify
         Caption         =   "Last"
         Height          =   255
         Left            =   50
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LblFirst 
         Alignment       =   1  'Right Justify
         Caption         =   "First"
         Height          =   255
         Left            =   50
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FraCCS 
      Caption         =   "Type"
      Height          =   1095
      Left            =   2520
      TabIndex        =   20
      Top             =   3000
      Width           =   3495
      Begin VB.OptionButton OptSires 
         Caption         =   "Sires"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptCows 
         Caption         =   "Cows"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton OptCalves 
         Caption         =   "Calves"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   255
         Left            =   3150
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "usdareps.frx":0000
      End
      Begin MSMask.MaskEdBox TxtEventDate 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label LblAnEvent 
         Caption         =   "An. Event Date"
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame FraEventCode 
      Caption         =   "Event Codes"
      Height          =   2175
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   5535
      Begin MhglbxLib.Mh3dList LstEventCodes 
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   3201
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
         ColScale        =   2
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
         ShadowColor     =   10070188
         WallPaper       =   0
         Sorted          =   0   'False
         TextColor       =   0
         WrapList        =   0   'False
         WrapWidth       =   0
         ColInstr        =   -1
         TitleHeight     =   0
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
         SelectedColor   =   16721960
         Transparent     =   0   'False
         TransparentColor=   0
         TitleFillColor  =   16776960
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
         ColTitle0       =   "AAE Code"
         ColWidth0       =   15
         ColTitle1       =   "AAE Description"
         ColWidth1       =   100
      End
   End
   Begin VB.Frame FraBirthdate 
      Caption         =   "Birthdate Range"
      Height          =   1000
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   2775
      Begin Threed.SSCommand SSCommand1 
         Height          =   255
         Left            =   2310
         TabIndex        =   29
         Top             =   600
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "usdareps.frx":04BE
      End
      Begin VB.Frame FraHarvestDates 
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   150
         TabIndex        =   12
         Top             =   1215
         Width           =   2880
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   2310
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "usdareps.frx":097C
      End
      Begin MSMask.MaskEdBox TxtStBirth 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin MSMask.MaskEdBox TxtEndBirth 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Birthdate"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "End Birthdate"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Herd"
      Height          =   345
      Left            =   3720
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   360
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
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   1000
   End
   Begin VB.Frame fraCowList 
      Caption         =   "Order By"
      Height          =   1000
      Left            =   135
      TabIndex        =   6
      Top             =   1920
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "EID"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   450
         Width           =   810
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Birthdate"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Label lblMultiCows 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All"
      Height          =   255
      Left            =   6000
      TabIndex        =   35
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label lblCows 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many Animals"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   2160
      Width           =   1665
   End
End
Attribute VB_Name = "usdareps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 3

Dim t As Integer
Dim reports(hmreps%) As String
Dim mCowList As New FrmSelect_Multi_Cows
Dim mSireList As New FrmSelect_Multi_Sires
Dim mCalfList As New FrmSelect_Multi_Calves






Sub Build_Cow_RPT_List(mhListBox As Mh3dList)
Dim DB As DAO.database, indx%
Dim pCowID$, pHerdID$
Set DB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from rptcows"
Do Until indx = mhListBox.ListCount
   If mhListBox.Tagged(indx) Then
      mhListBox.Col = 0
      pCowID = mhListBox.ColList(indx)
      mhListBox.Col = 1
      pHerdID = mhListBox.ColList(indx)
      DB.Execute "insert into rptcows (herdid, cowid) values ('" & pHerdID$ & "', '" & pCowID & "')"
   End If
   indx = indx + 1
Loop
DB.Close: Set DB = Nothing
End Sub


Sub Build_Calf_RPT_List(cListBox As Control)
Dim DB As DAO.database, indx%
Dim pCowID$, pHerdID$
Dim row As Integer
Set DB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from rptcalf"
'Do Until indx = mhListBox.ListCount
'   If mhListBox.Tagged(indx) Then
'      mhListBox.Col = 0
'      pCowID = mhListBox.ColList(indx)
'      mhListBox.Col = 1
'      pHerdID = mhListBox.ColList(indx)
'      DB.Execute "insert into rptcalf (herdid, calfid) values ('" & pHerdID$ & "', '" & pCowID & "')"
'   End If
'   indx = indx + 1
'Loop
   With cListBox
      For row = 0 To .SelCount
         .row = .NextSel
         If .row <> -1 Then
            pCowID = .ColList(0, .row)
            pHerdID = .ColList(1, .row)
            DB.Execute "insert into rptcalf (herdid, calfid) values ('" & pHerdID$ & "', '" & pCowID & "')"
         End If
      Next row
   End With
   DB.Close: Set DB = Nothing

End Sub

Sub Build_Sire_RPT_List(mhListBox As Mh3dList)
Dim DB As DAO.database, indx%
Dim pSireID$, pHerdID$
Set DB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from rptsire"
Do Until indx = mhListBox.ListCount
   If mhListBox.Tagged(indx) Then
      mhListBox.Col = 0
      pSireID = mhListBox.ColList(indx)
      mhListBox.Col = 1
      pHerdID = mhListBox.ColList(indx)
      DB.Execute "insert into rptsire (herdid, sireid) values ('" & pHerdID$ & "', '" & pSireID & "')"
   End If
   indx = indx + 1
Loop
DB.Close: Set DB = Nothing

End Sub


Private Sub Create_CalfAid2()
Dim pDB As DAO.database, pRS As DAO.Recordset, SQL$, where$, EIDID$, i%, EventDate As String
Set pDB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
EIDID$ = Trim$(Str(LstEventCodes.ListIndex))
If LstEventCodes.ListIndex = 15 Then EIDID$ = "20"
If LstEventCodes.ListIndex = 16 Then EIDID$ = "21"
If LstEventCodes.ListIndex = 17 Then EIDID$ = "22"

'clear rpt table
pDB.Execute "delete * from calfRef"
Set pDB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)

If OptCows.Value = True Then
  'add cows
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
  SQL = "insert into calfRef in '" & repfile & "' SELECT  cowprof.HerdID, cowprof.cowID as calfid, " & EventDate & " as birthdate, elecid, cowprof.regname as misc3, '" & EIDID$ & "' as misccode1, regnum as registration "
  If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM cowprof  inner join rptcows on cowprof.cowid  = rptcows.cowid "
  Else
    SQL = SQL & " FROM cowprof "
  End If
  SQL = SQL & " where cowprof.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTCows")
End If

If OptSires.Value = True Then
  'add Sires
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTSire", "RPTSire")
  SQL = "insert into calfref in '" & repfile$ & "' select sireprof.HerdID, sireprof.sireID as calfid, " & EventDate & " as birthdate, elecid, sireprof.regname as misc3, '" & EIDID$ & "' as misccode1, regnum as registration "
  If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM sireprof  inner join rptSire on sireprof.sireid  = rptSire.sireid "
  Else
    SQL = SQL & " FROM sireprof "
  End If
  SQL = SQL & " where sireprof.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTSire")
End If

If OptCalves.Value = True Then
  'add calfs
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
  SQL = "insert into calfref in '" & repfile$ & "' select calfbirth.HerdID,  calfbirth.calfid, " & EventDate & " as birthdate, elecid, misc3, '" & EIDID$ & "' as misccode1, registration "
  If lblMultiCows.Caption <> "All" Then
    SQL = SQL & " FROM CalfBirth inner join rptCalf on CalfBirth.Calfid  = rptCalf.Calfid "
   Else
    SQL = SQL & " FROM CalfBirth "
  End If
  SQL = SQL & " where calfbirth.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTCalf")
End If


pDB.Close: Set pDB = Nothing

End Sub


Private Sub Create_CalfAid3()
Dim pDB As DAO.database, pRS As DAO.Recordset, SQL$, where$, EIDID$, i%, EventDate As String
Set pDB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
EIDID$ = Trim$(Str(LstEventCodes.ListIndex))
If LstEventCodes.ListIndex = 15 Then EIDID$ = "20"
If LstEventCodes.ListIndex = 16 Then EIDID$ = "21"
If LstEventCodes.ListIndex = 17 Then EIDID$ = "22"

'clear rpt table
pDB.Execute "delete * from calfRef"
Set pDB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)


If OptCows.Value = True Then
  'add cows
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTCows", "RPTCows")
  SQL = "insert into calfRef in '" & repfile & "' SELECT  cowprof.HerdID, cowprof.cowID as calfid, birthdate, elecid, cowprof.regnum as registration, cowprof.regname as misc3, '" & EIDID$ & "' as misccode1 "
  If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM cowprof  inner join rptcows on cowprof.cowid  = rptcows.cowid "
  Else
    SQL = SQL & " FROM cowprof "
  End If
  SQL = SQL & " where cowprof.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTCows")
End If

If OptSires.Value = True Then
  'add Sires
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTSire", "RPTSire")
  SQL = "insert into calfref in '" & repfile$ & "' select sireprof.HerdID, sireprof.sireID as calfid, birthdate, elecid, sireprof.regnum as registration, sireprof.regname as misc3, '" & EIDID$ & "' as misccode1 "
  If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM sireprof inner join rptsire on sireprof.sireid  = rptsire.sireid "
  Else
    SQL = SQL & " FROM sireprof "
  End If
  SQL = SQL & " where sireprof.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTSire")
End If

If OptCalves.Value = True Then
  'add calfs
  If IsDate(TxtEventDate.TEXT) Then
    EventDate = "#" & TxtEventDate.TEXT & "#"
     Else
    'EventDate = "''"
    EventDate = "Null"
  End If
  Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
  SQL = "insert into calfref in '" & repfile$ & "' select calfbirth.HerdID,  calfbirth.calfid, " & EventDate & " as birthdate, elecid, misc3, '" & EIDID$ & "' as misccode1, registration "
  If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM CalfBirth inner join rptCalf on CalfBirth.Calfid  = rptCalf.Calfid "
  Else
    SQL = SQL & " FROM CalfBirth "
  End If

  SQL = SQL & " where calfbirth.herdid = '" & herdid & "'"
  If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
     SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
  End If
  If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
     SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
  End If
  pDB.Execute SQL, dbFailOnError
  Call DeleteTableAttachment(dbfile, "RPTCalf")
End If

pDB.Close: Set pDB = Nothing

End Sub


Private Sub Create_Cow_RefRPT()
Dim pDB As DAO.database, pRS As DAO.Recordset, SQL$, where$
Set pDB = DBEngine(0).OpenDatabase(repfile$, exclusiveyn%, readonlyyn%)
'clear rpt table
pDB.Execute "delete * from calfRef"
Set pDB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)

'add cows
'SQL = "insert into calfRef in '" & repfile & "' SELECT  cowprof.HerdID, cowprof.cowID as calfid, cowprof.birthdate, elecid "
'SQL = SQL & " FROM cowprof "
'SQL = SQL & " where cowprof.herdid = '" & herdid & "'"
'pDB.Execute SQL, dbFailOnError

'add Sires
'SQL = "insert into calfref in '" & repfile$ & "' select HerdID, sireID as calfid, birthdate, elecid   "
'SQL = SQL & " from sireprof where sireprof.herdid = '" & herdid & "'"
'pDB.Execute SQL, dbFailOnError


'add calfs
Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
SQL = "insert into calfref in '" & repfile$ & "' select calfbirth.HerdID,  calfbirth.calfid, birthdate, elecid, misc3  "
If lblMultiCows.Caption <> "All" Then
   SQL = SQL & " FROM CalfBirth inner join rptCalf on CalfBirth.Calfid  = rptCalf.Calfid "
  Else
   SQL = SQL & " FROM CalfBirth "
End If

SQL = SQL & " where calfbirth.herdid = '" & herdid & "'"
If IsDate(TxtStBirth.TEXT) And IsDate(TxtEndBirth.TEXT) Then
   SQL = SQL & " and birthdate between #" & TxtStBirth.TEXT & "# and #" & TxtEndBirth.TEXT & "#"
End If
If TxtFirst.TEXT <> "" And TxtLast.TEXT <> "" Then
   SQL = SQL & " and elecid between '" & TxtFirst.TEXT & "' and '" & TxtLast.TEXT & "'"
End If
pDB.Execute SQL, dbFailOnError
Call DeleteTableAttachment(dbfile, "RPTCalf")
pDB.Close: Set pDB = Nothing
End Sub


Private Sub cmdcancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
 selherd_List.Show vbModal
 If selherd_List.Tag = "CANCEL" Then Exit Sub
 herdid$ = selherd_List.Tag
 Unload selherd_List
 Screen.MousePointer = vbDefault
End Sub



Private Sub cmdCowSelect_Click()
If OptCows.Value = True Then
  mCowList.Show vbModal
  If mCowList!lstCows.SelectedCount > 0 Then
     lblMultiCows.Caption = Trim$(Str$(mCowList!lstCows.SelectedCount))
  Else
     lblMultiCows.Caption = "All"
  End If
End If
If OptSires.Value = True Then
  mSireList.Show vbModal
  If mSireList!lstSires.SelectedCount > 0 Then
     lblMultiCows.Caption = Trim$(Str$(mSireList!lstSires.SelectedCount))
  Else
     lblMultiCows.Caption = "All"
  End If
End If
If OptCalves.Value = True Then
  mCalfList.Show vbModal
  
   If mCalfList!lstCalves.SelCount > 0 Then
     lblMultiCows.Caption = Trim$(Str$(mCalfList!lstCalves.SelCount))
  Else
     lblMultiCows.Caption = "All"
  End If
End If

End Sub

Private Sub CMDOk_Click()
Dim order$, TITLE$, Herds$, title1$, title2$, Title3$, Title4$, XAvg#, XCows&
Dim tbData As Recordset
Dim DB As database
   
 Screen.MousePointer = vbHourglass
 report.Initialize ' init the class
 If optprint Then report.SetDestination = 1
 If Val(lblMultiCows.Caption) > 0 Then
    If OptCows.Value = True Then Call Build_Cow_RPT_List(mCowList.lstCows)
    If OptSires.Value = True Then Call Build_Sire_RPT_List(mSireList.lstSires)
    If OptCalves.Value = True Then Call Build_Calf_RPT_List(mCalfList.lstCalves)
 End If
 Select Case lstreports.ListIndex
  Case 0
   Call Create_Cow_RefRPT
   report.SetReportFileName = dbdir$ & "" & "calfaid.rpt"
   report.setDbname = repfile$
   If optOrder(0).Value Then report.Setformulas("groupby") = "{calfref.CalfID}"
   If optOrder(1).Value Then report.Setformulas("groupby") = "{calfref.elecid}"
   If optOrder(2).Value Then report.Setformulas("groupby") = "{calfref.birthdate}"
   
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 
   Set tbData = DB.OpenRecordset("herd", dbOpenTable)
   tbData.Index = "primarykey"
   tbData.Seek "=", herdid
   If Not tbData.NoMatch Then report.Setformulas("premise") = "'" & tbData!Name & "'"
   Set tbData = Nothing
   Set DB = Nothing
   
  Case 1
   Call Create_CalfAid2
   report.SetReportFileName = dbdir$ & "" & "calfaid2.rpt"
   report.setDbname = repfile$
   If optOrder(0).Value Then report.Setformulas("groupby") = "{calfref.CalfID}"
   If optOrder(1).Value Then report.Setformulas("groupby") = "{calfref.elecid}"
   If optOrder(2).Value Then report.Setformulas("groupby") = "{calfref.birthdate}"
   
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 
   Set tbData = DB.OpenRecordset("herd", dbOpenTable)
   tbData.Index = "primarykey"
   tbData.Seek "=", herdid
   If Not tbData.NoMatch Then report.Setformulas("premise") = "'" & tbData!Name & "'"
   Set tbData = Nothing
   Set DB = Nothing
  
  
  Case 2
   Call Create_CalfAid3
   report.SetReportFileName = dbdir$ & "" & "calfaid3.rpt"
   report.setDbname = repfile$
   If optOrder(0).Value Then report.Setformulas("groupby") = "{calfref.CalfID}"
   If optOrder(1).Value Then report.Setformulas("groupby") = "{calfref.elecid}"
   If optOrder(2).Value Then report.Setformulas("groupby") = "{calfref.birthdate}"
   
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 
   Set tbData = DB.OpenRecordset("herd", dbOpenTable)
   tbData.Index = "primarykey"
   tbData.Seek "=", herdid
   If Not tbData.NoMatch Then report.Setformulas("premise") = "'" & tbData!Name & "'"
   Set tbData = Nothing
   Set DB = Nothing
  
 End Select
  report.PrintReport
  Screen.MousePointer = vbDefault
End Sub
 


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "CalfAID PVP"
 reports$(2) = "ATD Animal Tracking Database"
 reports$(3) = "C.O.C.O.O"
 For t = 1 To hmreps%
     lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 LstEventCodes.AddItem "0" & vbTab & "AIN Device Distributed"
 LstEventCodes.AddItem "1" & vbTab & "AIN Allocated by USDA/Aphis"
 LstEventCodes.AddItem "2" & vbTab & "Tag Applied"
 LstEventCodes.AddItem "3" & vbTab & "Moved In"
 LstEventCodes.AddItem "4" & vbTab & "Moved Out"
 LstEventCodes.AddItem "5" & vbTab & "Lost Tag"
 LstEventCodes.AddItem "6" & vbTab & "Replaced Tag or Re-tagged"
 LstEventCodes.AddItem "7" & vbTab & "Imported"
 LstEventCodes.AddItem "8" & vbTab & "Exported"
 LstEventCodes.AddItem "9" & vbTab & "Animal at Location"
 LstEventCodes.AddItem "10" & vbTab & "Slaughtered"
 LstEventCodes.AddItem "11" & vbTab & "Died"
 LstEventCodes.AddItem "12" & vbTab & "Tag Retired"
 LstEventCodes.AddItem "13" & vbTab & "Animal Missing"
 LstEventCodes.AddItem "14" & vbTab & "ICVI - Certificate of Veterinary Inspection"
 'LstEventCodes.AddItem "15" & vbTab & "AIN Device Distributed"
 'LstEventCodes.AddItem "16" & vbTab & "AIN Device Distributed"
 'LstEventCodes.AddItem "17" & vbTab & "AIN Device Distributed"
 'LstEventCodes.AddItem "18" & vbTab & "AIN Device Distributed"
 'LstEventCodes.AddItem "19" & vbTab & "AIN Device Distributed"
 LstEventCodes.AddItem "20" & vbTab & "Invalid AIN"
 LstEventCodes.AddItem "21" & vbTab & "AIN Recalled from Manufacturer"
 LstEventCodes.AddItem "22" & vbTab & "AIN Device Returned to Manufacturer"
 LstEventCodes.ListIndex = 0
 Load mCowList 'Select multi cow form
 Load mSireList 'Select multi sire form
 Load mCalfList 'Select multi calf form
 OptCalves.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload mCowList 'Select multi cow form
Unload mSireList 'Select multi sire form
Unload mCalfList 'Select multi calf form
Load FrmSelect_Multi_Herds
Set FrmSelect_Multi_Herds = Nothing
Set mCowList = Nothing
Set mSireList = Nothing
Set mCalfList = Nothing
End Sub

Private Sub lstreports_Click()
   FraCCS.Visible = False
   FraEID.Visible = True
   FraEventCode.Visible = False
   lblMultiCows.Visible = True
   cmdCowSelect.Visible = True
   lblCows.Visible = True
   lblCows.Caption = "How many Calves"
   If lstreports.ListIndex = 1 Or lstreports.ListIndex = 2 Then
      FraCCS.Visible = True
      FraEventCode.Visible = True
   End If
   If lstreports.ListIndex = 2 Then
     LblAnEvent.Visible = False
     TxtEventDate.Visible = False
     SSCommand4.Visible = False
     lblCows.Caption = "How many Sires"
     If OptCows.Value = True Then lblCows.Caption = "How many Cows"
     If OptCalves.Value = True Then lblCows.Caption = "How many Calves"
     If OptSires.Value = True Then lblCows.Caption = "How many Sires"
   End If
   If lstreports.ListIndex = 1 Then
     LblAnEvent.Visible = True
     TxtEventDate.Visible = True
     SSCommand4.Visible = True
     If OptCows.Value = True Then lblCows.Caption = "How many Cows"
     If OptCalves.Value = True Then lblCows.Caption = "How many Calves"
     If OptSires.Value = True Then lblCows.Caption = "How many Sires"
   End If
   Call Display_Criteria
End Sub

Private Sub Display_Criteria()
Select Case lstreports.ListIndex
Case 0
End Select
Exit Sub
End Sub

Private Sub MskEventDate_Change()
gcaldate = TxtStBirth.TEXT
Call GetDate(gcaldate)
TxtStBirth.TEXT = gcaldate

End Sub

Private Sub OptCalves_Click()
  lblCows.Caption = "How many Calves"
  If lstreports.ListIndex = 2 Or lstreports.ListIndex = 0 Then
    LblAnEvent.Visible = False
    TxtEventDate.Visible = False
    SSCommand4.Visible = False
   Else
    LblAnEvent.Visible = True
    TxtEventDate.Visible = True
    SSCommand4.Visible = True
  End If
End Sub

Private Sub OptCows_Click()
  lblCows.Caption = "How many Cows"
  If lstreports.ListIndex = 2 Or lstreports.ListIndex = 0 Then
    LblAnEvent.Visible = False
    TxtEventDate.Visible = False
    SSCommand4.Visible = False
   Else
    LblAnEvent.Visible = True
    TxtEventDate.Visible = True
    SSCommand4.Visible = True
  End If
End Sub


Private Sub OptSires_Click()
  lblCows.Caption = "How many Sires"
  If lstreports.ListIndex = 2 Or lstreports.ListIndex = 0 Then
    LblAnEvent.Visible = False
    TxtEventDate.Visible = False
    SSCommand4.Visible = False
   Else
    LblAnEvent.Visible = True
    TxtEventDate.Visible = True
    SSCommand4.Visible = True
  End If
End Sub


Private Sub SSCommand1_Click()
gcaldate = TxtEndBirth.TEXT
Call GetDate(gcaldate)
TxtEndBirth.TEXT = gcaldate

End Sub

Private Sub SSCommand3_Click()
gcaldate = TxtStBirth.TEXT
Call GetDate(gcaldate)
TxtStBirth.TEXT = gcaldate

End Sub


Private Sub SSCommand4_Click()
gcaldate = TxtEventDate.TEXT
Call GetDate(gcaldate)
TxtEventDate.TEXT = gcaldate
End Sub


