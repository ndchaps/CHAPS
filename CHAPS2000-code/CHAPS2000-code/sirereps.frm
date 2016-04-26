VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form sirereps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sire Reports"
   ClientHeight    =   3165
   ClientLeft      =   720
   ClientTop       =   1455
   ClientWidth     =   6810
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   6810
   Begin MhglbxLib.Mh3dList lstreports 
      Height          =   1575
      Left            =   45
      TabIndex        =   1
      Top             =   30
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
   Begin VB.CommandButton CmdSelectSire 
      Caption         =   "&Select"
      Height          =   375
      Left            =   1890
      TabIndex        =   15
      Top             =   2055
      Width           =   1290
   End
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Herd"
      Height          =   375
      Left            =   1890
      TabIndex        =   14
      Top             =   1650
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Frame FraPerf 
      Height          =   1110
      Left            =   3300
      TabIndex        =   6
      Top             =   435
      Width           =   3435
      Begin VB.Frame FraInclude 
         Caption         =   "Include"
         Height          =   900
         Left            =   2010
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1365
         Begin VB.OptionButton OptPerf 
            Caption         =   "Pedigree"
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   13
            Top             =   630
            Width           =   1125
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Culled"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   12
            Top             =   420
            Width           =   1125
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Active"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   11
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame FraSort 
         Caption         =   "Sort By"
         Height          =   900
         Left            =   45
         TabIndex        =   7
         Top             =   120
         Width           =   1290
         Begin VB.OptionButton OptPerf 
            Caption         =   "Sire ID"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   9
            Top             =   600
            Width           =   1125
         End
         Begin VB.OptionButton OptPerf 
            Caption         =   "Cow ID"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   8
            Top             =   255
            Width           =   1125
         End
      End
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   705
      Left            =   1155
      TabIndex        =   3
      Top             =   2445
      Width           =   1215
      Begin VB.OptionButton optprint 
         Caption         =   "Print"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optpreview 
         Caption         =   "Preview"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   5745
      TabIndex        =   2
      Top             =   2460
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   4305
      TabIndex        =   0
      Top             =   2460
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "How Many Sires"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2070
      Width           =   1185
   End
   Begin VB.Label LBLHowManySires 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All"
      Height          =   255
      Left            =   1395
      TabIndex        =   16
      Top             =   2070
      Width           =   420
   End
End
Attribute VB_Name = "sirereps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hmreps% = 4

Dim t As Integer
Dim reports(hmreps%) As String

Private Sub Create_Pedigree_RPT()
Dim SQL$, DB As DAO.database, where$, orderby$, RS As DAO.Recordset, tbQScor As DAO.Recordset
Dim repdb As DAO.database
Dim tbavg As Recordset
Dim tbData As Recordset
Dim theavg As Double
Dim theagvcnt As Integer
On Local Error GoTo ErrHandler
100 Set DB = DBEngine(0).OpenDatabase(repfile, False, False)
110 DB.Execute "delete * from S_Perf"
120 DB.Execute "delete * from s_perf_avgs"
Set DB = DBEngine(0).OpenDatabase(dbfile, False, False)
Set repdb = DBEngine(0).OpenDatabase(repfile, False, False)

Call CreateTableAttachment(dbfile$, repfile$, "S_Perf", "S_Perf")
Screen.MousePointer = vbHourglass
GoSub Build_Sire_Data
GoSub Update_2nd_Gen
GoSub Update_3rd_Gen
GoSub Build_S_Perf_Avgs
GoSub Build_Order
TEXT$(1) = ""
Call DeleteTableAttachment(dbfile, "S_Perf")
130 DB.Close: Set DB = Nothing
 repdb.Close: Set repdb = Nothing
Exit Sub

ErrHandler:
'TEXT$(1) = ""
TEXT$(2) = Erl
TEXT$(3) = ""
TEXT$(4) = ""
TEXT$(5) = ""
 GMODNAME$ = Me.Name & " - Create_Pedigree_RPT"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

Build_Sire_Data:
   TEXT$(1) = "Build_Sire_Data"
200    SQL = "insert into s_perf in '" & repfile & "' "
210    SQL = SQL & " SELECT DISTINCTROW sireprof.SireID, sireprof.calfid AS Birth_ID, sireprof.sire AS S2_Name, sireprof.dam AS C2_Name, calfbirth.birthdate, calfbirth.birthwt, calfwean.wt205, calfwean.ratio, calfwean.score, calfrep.w365, calfcarcass.ygrade, calfcarcass.qgrade, calfcarcass.ywt, sireepd.epdbirthwt, sireepd.epdweanwt, sireepd.epdyearwt, sireepd.epdmatww, sireepd.epdmatmilk, sireepd.accbirthwt, sireepd.accweanwt, sireepd.accyearwt, sireepd.accmatww, sireepd.accmatmilk, sireepd.misc1, sireepd.misc2, sireepd.misc3, sireepd.misc4, sireepd.misc5, sireepd.misc6, sireepd.misc7, sireepd.misc8, sireepd.misc9, sireepd.misc10, sireepd.acc1, sireepd.acc2, sireepd.acc3, sireepd.acc4, sireepd.acc5, sireepd.acc6, sireepd.acc7, sireepd.acc8, sireepd.acc9, sireepd.acc10, sireprof.notes, sireprof.regnum, sireprof.regname, sireprof.herdid "
220    If Val(LBLHowManySires.Caption) > 0 Then
            SQL = SQL & " FROM RPTSire INNER JOIN (((((sireprof LEFT JOIN sireepd ON (sireprof.SireID = sireepd.SireID) AND (sireprof.HerdID = sireepd.HerdID)) LEFT JOIN calfbirth ON (sireprof.calfid = calfbirth.CalfID) AND (sireprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) ON (RPTSire.SireID = sireprof.SireID) AND (RPTSire.HerdID = sireprof.HerdID) "
         Else
            SQL = SQL & " FROM ((((sireprof LEFT JOIN sireepd ON (sireprof.SireID = sireepd.SireID) AND (sireprof.HerdID = sireepd.HerdID)) LEFT JOIN calfbirth ON (sireprof.calfid = calfbirth.CalfID) AND (sireprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID) "
         End If
       SQL = SQL & " where sireprof.herdid = '" & herdid & "' "
230 DB.Execute SQL, dbFailOnError
Return

Update_2nd_Gen:
   TEXT$(1) = "Update_2nd_Gen"
   DB.Execute "UPDATE (S_Perf LEFT JOIN sireprof ON (S_Perf.herdid = sireprof.HerdID) AND (S_Perf.S2_Name = sireprof.SireID)) LEFT JOIN cowprof ON (S_Perf.herdid = cowprof.HerdID) AND (S_Perf.C2_Name = cowprof.cowID) SET S_Perf.C2_regnum = [cowprof].[regnum], S_Perf.C2_regname = [cowprof].[regname], S_Perf.S2_regnum = [sireprof].[regnum], S_Perf.S2_regname = [sireprof].[regname], S_Perf.S3_Name = [sireprof].[sire], S_Perf.C3_Name = [sireprof].[dam], S_Perf.S4_Name = [cowprof].[sire], S_Perf.C4_Name = [cowprof].[dam]"
   DB.Execute "UPDATE (((S_Perf LEFT JOIN cowprof AS cowprof_1 ON (S_Perf.C3_Name = cowprof_1.cowID) AND (S_Perf.herdid = cowprof_1.HerdID)) LEFT JOIN cowprof ON (S_Perf.C4_Name = cowprof.cowID) AND (S_Perf.herdid = cowprof.HerdID)) LEFT JOIN sireprof AS sireprof_1 ON (S_Perf.S3_Name = sireprof_1.SireID) AND (S_Perf.herdid = sireprof_1.HerdID)) LEFT JOIN sireprof ON (S_Perf.S4_Name = sireprof.SireID) AND (S_Perf.herdid = sireprof.HerdID) SET S_Perf.C3_regnum = [cowprof_1].[regnum], S_Perf.C3_regname = [cowprof_1].[regname], S_Perf.C4_regnum = [cowprof].[regnum], S_Perf.C4_regname = [cowprof].[regname], S_Perf.S3_regnum = [sireprof_1].[regnum], S_Perf.S3_regname = [sireprof_1].[regname], S_Perf.S4_regnum = [sireprof].[regnum], S_Perf.S4_regname = [sireprof].[regname]"
Return

Update_3rd_Gen:
   TEXT$(1) = "Update_3rd_Gen"
   
Return

Build_S_Perf_Avgs:
   TEXT$(1) = "Build_S_Perf_Avgs"
   'SQL = "insert into s_perf_avgs in '" & repfile & "' SELECT DISTINCTROW sireprof.SireID, Count(Calfbirth.calfid) as numprogeny, Sum(iif(calfbirth.birthwt>0,calfbirth.birthwt,0)) / Sum(iif(calfbirth.birthwt>0,1,0)) AS AvgOfbirthwt, Sum(iif(calfwean.wt205>0,switch(calfbirth.sex = '0', calfwean.wt205 * 1, calfbirth.sex = '1', calfwean.wt205 * .95, calfbirth.sex = '2', calfwean.wt205 * 1.05, calfbirth.sex = '3', calfwean.wt205 * 1),0))/Sum(iif(calfwean.wt205>0,1,0)) AS AvgOfwt205, Sum(iif(calfwean.ratio>0,calfwean.ratio,0))/Sum(iif(calfwean.ratio>0,1,0)) AS AvgOfratio, Sum(iif(calfwean.score>0, calfwean.score, 0))/Sum(iif(calfwean.score>0,1,0)) AS AvgOfscore, Sum(iif(calfrep.w365>0,calfrep.w365,0))/Sum(iif(calfrep.w365>0,1,0)) AS AvgOfw365, Sum(iif(calfcarcass.ygrade>0,calfcarcass.ygrade,0))/Sum(iif(calfcarcass.ygrade>0,1,0)) AS AvgOfygrade, Sum(iif(calfcarcass.ywt>0,calfcarcass.ywt,0))/Sum(iif(calfcarcass.ywt>0,1,0)) AS AvgOfywt"
   SQL = "insert into s_perf_avgs in '" & repfile & "' SELECT DISTINCTROW sireprof.SireID, Count(calfbirth.CalfID) AS numprogeny, Sum(IIf(calfbirth.birthwt>0,calfbirth.birthwt,0))/Sum(IIf(calfbirth.birthwt>0,1,0)) AS AvgOfbirthwt, Sum(IIf(calfwean.wt205>0,Switch(calfbirth.sex='0',calfwean.wt205*1,calfbirth.sex='1',calfwean.wt205*0.95,calfbirth.sex='2',calfwean.wt205*1.05,calfbirth.sex='3',calfwean.wt205*1),0))/Sum(IIf(calfwean.wt205>0,1,0)) AS AvgOfwt205, Sum(IIf(calfwean.ratio>0,calfwean.ratio,0))/Sum(IIf"
   'SQL = SQL & " (calfwean.ratio>0,1,0)) AS AvgOfratio, Sum(IIf(calfwean.score>0,calfwean.score,0))/Sum(IIf(calfwean.score>0,1,0)) AS AvgOfscore, Sum(IIf(calfrep.w365>0,calfrep.w365,0))/Sum(IIf(calfrep.w365>0,1,0)) AS AvgOfw365, Sum(IIf(calfcarcass.ygrade>0,calfcarcass.ygrade,0))/Sum(IIf(calfcarcass.ygrade>0,1,0)) AS AvgOfygrade, Sum(IIf(calfcarcass.ywt>0,calfcarcass.ywt,0))/Sum(IIf(calfcarcass.ywt>0,1,0)) AS AvgOfywt, Year(Max([calfbirth].[birthdate]))-Year(Min([calfbirth].[birthdate]))+1 AS Years_Service, Sum(switch(calfcarcass.qgrade = 'Prime+', 10.5, calfcarcass.qgrade = 'Prime ', 9.5, calfcarcass.qgrade = 'Prime-', 8.5, calfcarcass.qgrade = 'Choice+', 7.0, calfcarcass.qgrade = 'Choice', 6.5, calfcarcass.qgrade = 'Choice-', 5.5, calfcarcass.qgrade = 'Select+', 4.75, calfcarcass.qgrade = 'Select-', 4.25, calfcarcass.qgrade = 'Standard+', 3.5, calfcarcass.qgrade= 'Standard-', 2.5, isnull(calfcarcass.qgrade), 0)) / Sum(iif(isnull(calfcarcass.qgrade) = false, 1, 0)) AS Avg_QScore, sireprof.herdid "
   SQL = SQL & " (calfwean.ratio>0,1,0)) AS AvgOfratio, Sum(IIf(calfwean.score>0,calfwean.score,0))/Sum(IIf(calfwean.score>0,1,0)) AS AvgOfscore, Sum(IIf(calfrep.w365>0,calfrep.w365,0))/Sum(IIf(calfrep.w365>0,1,0)) AS AvgOfw365, Sum(IIf(calfcarcass.ygrade>0,calfcarcass.ygrade,0))/Sum(IIf(calfcarcass.ygrade>0,1,0)) AS AvgOfygrade, Sum(IIf(calfcarcass.ywt>0,calfcarcass.ywt,0))/Sum(IIf(calfcarcass.ywt>0,1,0)) AS AvgOfywt, Year(Max([calfbirth].[birthdate]))-Year(Min([calfbirth].[birthdate]))+1 AS Years_Service, "
   SQL = SQL & " 0 AS Avg_QScore, sireprof.herdid "
   SQL = SQL & " FROM (((sireprof INNER JOIN calfbirth ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID)) LEFT JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)) LEFT JOIN calfcarcass ON (calfbirth.HerdID = calfcarcass.HerdID) AND (calfbirth.CalfID = calfcarcass.CalfID)) LEFT JOIN calfrep ON (calfbirth.CalfID = calfrep.CalfID) AND (calfbirth.HerdID = calfrep.HerdID)" & where
   SQL = SQL & " GROUP BY sireprof.SireID, sireprof.herdid"
460    DB.Execute SQL



' " build AVG_Qscore "
 
 Set tbavg = repdb.OpenRecordset("s_perf_avgs", dbOpenTable)
 While Not tbavg.EOF
   SQL = "SELECT DISTINCTROW sireprof.SireID, sireprof.HerdID, calfbirth.CalfID, calfcarcass.qgrade FROM sireprof INNER JOIN (calfbirth LEFT JOIN calfcarcass ON (calfbirth.CalfID = calfcarcass.CalfID) AND (calfbirth.HerdID = calfcarcass.HerdID)) ON (sireprof.SireID = calfbirth.sireID) AND (sireprof.HerdID = calfbirth.HerdID)    WHERE (((sireprof.SireID)='" & tbavg!sireid & "') AND ((sireprof.HerdID)='" & tbavg!herdid & "'))"
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
    
     
     tbavg!Avg_QScore = RESPONSE 'funround2(2, theavg / theagvcnt)
     tbavg.Update
   End If
   tbavg.MoveNext
 Wend
 tbavg.Close: Set tbavg = Nothing
 tbData.Close: Set tbData = Nothing


Return

Build_Order:
   TEXT$(1) = "Build_Order"
   Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
   If OptPerf(1).Value Then orderby = " order by sireid "
   If OptPerf(0).Value Then orderby = " order by c2_name "
500    DB.Execute "select * into tmpS_Perf from s_perf " & orderby
510    DB.Execute "delete * from S_Perf"
520    DB.Execute "insert into s_perf select * from tmps_perf"
530    DB.Execute "drop table tmps_perf"
Return

End Sub

Private Sub create_refLIST_report()
Dim DB As DAO.database, SQL$, where$
'If OptPerf(2).Value Then Where = " where sireprof.active = 'A' "
'If OptPerf(3).Value Then Where = " where sireprof.active = 'C' "
'If OptPerf(4).Value Then Where = " where sireprof.active = 'P' "
'Where = Where & " and sireprof.herdid = '" & herdid & "' "
where = " where sireprof.herdid = '" & herdid & "' "
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from sireref"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
SQL = "insert into sireref in '" & repfile & "' SELECT DISTINCTROW sireprof.HerdID, sireprof.SireID, sireprof.birthdate, sireprof.breed, sireprof.sire, sireprof.calfid, sireprof.regnum, sireprof.regname, sireprof.elecID, sireprof.enteredherd, sireprof.source, sireprof.active, sireprof.notes, SIREPROF.DAM "
If Val(LBLHowManySires.Caption) > 0 Then
   SQL = SQL & " FROM RPTSire INNER JOIN sireprof ON (RPTSire.SireID = sireprof.SireID) AND (RPTSire.HerdID = sireprof.HerdID) "
Else
   SQL = SQL & " From sireprof "
End If
SQL = SQL & where & sortsires
DB.Execute SQL, dbFailOnError
DB.Close: Set DB = Nothing
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
herdid$ = selherd_List.Tag
End Sub

Private Sub BuildRPTSire()
Dim DB As DAO.database, indx%, pHerd$, pSire$
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from rptsire"
Do Until indx = FrmSelect_Multi_Sires.lstSires.ListCount
   If FrmSelect_Multi_Sires.lstSires.Tagged(indx) = True Then
      FrmSelect_Multi_Sires.lstSires.Col = 0: pSire = FrmSelect_Multi_Sires.lstSires.ColList(indx)
      FrmSelect_Multi_Sires.lstSires.Col = 1: pHerd = FrmSelect_Multi_Sires.lstSires.ColList(indx)
      DB.Execute "insert into rptsire (HerdID, SireID) Values ('" & pHerd & "', '" & pSire & "')"
   End If
   indx = indx + 1
Loop
DB.Close: Set DB = Nothing
End Sub

Private Sub CMDOk_Click()
 report.Initialize ' init the class
 If optprint Then report.SetDestination = 1
 If Val(LBLHowManySires.Caption) > 0 Then
   Call CreateTableAttachment(dbfile, repfile, "RPTSire", "RPTSire")
   Call BuildRPTSire
 End If
 Select Case lstreports.ListIndex
    Case 0
      Call create_refLIST_report
      report.SetReportFileName = dbdir$ & "\" & "sireREF.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(1)
      report.Setcommonformulas("", "", "") = ""
   Case 1
      Call Create_Pedigree_RPT
      report.setDbname = repfile$
      report.SetReportCaption = reports(2)
      report.Setcommonformulas("", "", "") = ""
      report.SetReportFileName = dbdir$ & "\S_Pedi.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
      If LBLHowManySires.Caption <> "All" Then
         If FrmSelect_Multi_Sires.OptType(0).Value Then report.Setformulas("Status") = "'Active'"
         If FrmSelect_Multi_Sires.OptType(1).Value Then report.Setformulas("Status") = "'Culled'"
         If FrmSelect_Multi_Sires.OptType(2).Value Then report.Setformulas("Status") = "'Pedigree'"
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
   Case 2
      Call Create_Pedigree_RPT
      report.setDbname = repfile$
      report.SetReportCaption = reports(3)
      report.Setcommonformulas("", "", "") = ""
      report.SetReportFileName = dbdir$ & "\S_Perf.rpt"
      report.Setformulas("Herdid") = "'Herd ID: " & herdid & "'"
   Case 3
      Call Create_Sire_NotesRPT
      report.SetReportFileName = dbdir$ & "\" & "sirenote.rpt"
      report.setDbname = repfile$
      report.SetReportCaption = reports$(4)
      report.Setcommonformulas("", "", "") = ""
      report.Setformulas("herdid") = "'Herd ID: " & herdid & "'"
End Select
Call DeleteTableAttachment(dbfile, "RPTSire")
report.PrintReport
End Sub
 
Private Sub Create_Sire_NotesRPT()
Dim SQL$, DB As DAO.database, where$, orderby$
Set DB = DBEngine(0).OpenDatabase(repfile, exclusiveyn%, readonlyyn%)
DB.Execute "delete * from sire_notes"
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
If OptPerf(0).Value Then orderby = " order by sireprof.sireid "
If OptPerf(1).Value Then orderby = " order by sireprof.birthdate "
'If OptPerf(2).Value Then Where = " where sireprof.active = 'A' "
'If OptPerf(3).Value Then Where = " where sireprof.active = 'C' "
'If OptPerf(4).Value Then Where = " where sireprof.active = 'P' "
where = where & " where sireprof.herdid = '" & herdid & "' "
SQL = "insert into sire_notes in '" & repfile & "' SELECT DISTINCTROW sireprof.SireID, sireprof.birthdate, sireprof.notes "
If Val(LBLHowManySires.Caption) > 0 Then
   SQL = SQL & " FROM RPTSire INNER JOIN sireprof ON (RPTSire.SireID = sireprof.SireID) AND (RPTSire.HerdID = sireprof.HerdID) "
Else
   SQL = SQL & " FROM sireprof " & where & orderby
End If
DB.Execute SQL
DB.Close: Set DB = Nothing
End Sub
 
Private Sub cmdselectvend_Click()
' FrmSelect_Multi_Herds.Show vbModal
' If FrmSelect_Multi_Herds!lstherd.SelectedCount > 0 Then
'   lblhow_many_herd.Caption = Trim$(Str$(FrmSelect_Multi_Herds!lstherd.SelectedCount))
'  Else
'   lblhow_many_herd.Caption = "All"
' End If
'
End Sub

Private Sub CmdSelectSire_Click()
FrmSelect_Multi_Sires.Show vbModal
If Val(FrmSelect_Multi_Sires.lbltagged) = 0 Then LBLHowManySires.Caption = "All" Else LBLHowManySires.Caption = FrmSelect_Multi_Sires.lbltagged.Caption
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 reports$(1) = "Sire Reference List"
 reports$(2) = "Sire Performance Pedigree Report"
 reports$(3) = "Sire Performance Report"
 reports$(4) = "Sire Notes Report"
 For t = 1 To hmreps%
  lstreports.AddItem reports$(t)
 Next t
 lstreports.ListIndex = 0
 optpreview.Value = True
 OptPerf(0).Value = True
 OptPerf(2).Value = True
' lblhow_many_herd.Caption = "1"
Load FrmSelect_Multi_Herds
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload FrmSelect_Multi_Herds
Set FrmSelect_Multi_Herds = Nothing
End Sub

Private Sub lstreports_Click()
FraSort.Visible = True
OptPerf(1).Caption = "Sire ID"
Select Case lstreports.ListIndex
   Case 0
      FraSort.Visible = False
   Case 3
      OptPerf(1).Caption = "Birth Date"
End Select
End Sub

