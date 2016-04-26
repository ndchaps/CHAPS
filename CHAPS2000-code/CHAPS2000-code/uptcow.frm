VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form FrmUpdate_Cow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Notes Field"
   ClientHeight    =   3420
   ClientLeft      =   2595
   ClientTop       =   4830
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   5400
   Begin MhglbxLib.Mh3dList lstcow 
      Height          =   2895
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   5106
      _StockProps     =   79
      Caption         =   "Cow Id"
      BackColor       =   16777215
      TintColor       =   16711935
      Caption         =   "Cow Id"
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
      MultiSelect     =   2
      PictureHeight   =   0
      PictureWidth    =   0
      AdjustHeight    =   0
      ScrollBars      =   1
      ShadowColor     =   8421504
      WallPaper       =   0
      Sorted          =   -1  'True
      TextColor       =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      ColInstr        =   0
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
      SelectedColor   =   8388608
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
      ColTitle0       =   "Cow ID"
      ColWidth0       =   10
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   2475
      TabIndex        =   4
      Top             =   -60
      Width           =   2835
      Begin VB.ComboBox CBOTab 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1470
         Width           =   1215
      End
      Begin VB.ComboBox cboyear 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1470
         Width           =   1275
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "Change Herd"
         Height          =   345
         Left            =   1455
         TabIndex        =   11
         Top             =   945
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Height          =   1050
         Left            =   45
         TabIndex        =   7
         Top             =   165
         Width           =   1350
         Begin VB.OptionButton OptAct 
            Caption         =   "Active"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   180
            Width           =   1230
         End
         Begin VB.OptionButton OptAct 
            Caption         =   "Culled"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   465
            Width           =   1230
         End
         Begin VB.OptionButton OptAct 
            Caption         =   "Pedigree"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   750
            Width           =   1230
         End
      End
      Begin VB.CommandButton CmdUntag 
         Caption         =   "Untag All"
         Height          =   360
         Left            =   1455
         TabIndex        =   6
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton CmdTagAll 
         Caption         =   "Tag All"
         Height          =   360
         Left            =   1455
         TabIndex        =   5
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.TextBox txtUpdate 
      Height          =   990
      Left            =   2490
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1890
      Width           =   2625
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   4110
      TabIndex        =   1
      Top             =   2985
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Update"
      Height          =   385
      Left            =   2985
      TabIndex        =   0
      Top             =   2985
      Width           =   1000
   End
End
Attribute VB_Name = "FrmUpdate_Cow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadCows()
Dim DB As database, RS As Recordset, strSQL$, CurDate$
If cboyear.TEXT = "" Or cboyear.TEXT = "Not Set" Then Exit Sub
Screen.MousePointer = vbHourglass
CurDate = Left(cboyear.TEXT, 10)
 
 If cboyear.TEXT <> "Not Set" And cboyear.Enabled = True Then strSQL = strSQL & " and cowbrd.calfdate = #" & CurDate & "# "
 If OptAct(0).Value Then
   strSQL$ = "SELECT DISTINCTROW calfbirth.HerdID, calfbirth.CowID, Sum(IIf([calfbirth].[birthdate]>=#" & CurDate & "# And [calfbirth].[birthdate]<=#" & CDate(CurDate) + 365 & "#,1,0)) AS Calf_Count FROM cowprof RIGHT JOIN (cowbrd RIGHT JOIN calfbirth ON (cowbrd.HerdID = calfbirth.HerdID) AND (cowbrd.CowID = calfbirth.CowID)) ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID) where calfbirth.herdid = '" & herdid & "' "
   strSQL = strSQL & " and cowprof.active = 'A' "
   strSQL = strSQL & " GROUP BY calfbirth.HerdID, calfbirth.CowID ORDER BY calfbirth.CowID"
End If
 If OptAct(1).Value Then strSQL = "select herdid, cowid from cowprof where herdid = '" & herdid & "' and cowprof.active = 'C' "
 If OptAct(2).Value Then strSQL = "select herdid, cowid from cowprof where herdid = '" & herdid & "' and cowprof.active = 'P' "
 
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset(strSQL$)
 lstcow.Clear
 Do Until tbData.EOF
    'Select Case tbData!calf_count
    '  Case Is = 0
    '     lstcow.TextColor = vbBlack
    '  Case Is >= 1
    '     lstcow.TextColor = vbRed
    '
    '
    'End Select
    lstcow.AddItem Field2Str(tbData!CowID) & Chr$(9) & Field2Str(tbData!herdid)
    tbData.MoveNext
 Loop
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub CBOTab_Click()
If CBOTab.TEXT <> "Breeding" Then cboyear.Enabled = False Else cboyear.Enabled = True
End Sub

Private Sub cboyear_Click()
Call LoadCows
End Sub

Private Sub cmdcancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub

Private Sub cmdchange_Click()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
herdid$ = selherd_List.Tag
Call LoadCows
End Sub

Private Sub CmdEdit_Click()
Dim DB As DAO.database, SQL$, indx%
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
DB.Execute "delete * from rptcalf"
Do Until indx = lstcow.ListCount
   If lstcow.Tagged(indx) Then
      lstcow.Col = 0
      DB.Execute "insert into RPTCalf (CalfID, HerdID) Values ('" & lstcow.ColList(indx) & "', '" & herdid & "')"
   End If
   indx = indx + 1
Loop
Select Case CBOTab.ListIndex
   Case 0
      SQL = "update cowprof, rptcalf set notes =  left(notes & ' " & txtUpdate & "', 250) where cowprof.cowid = rptcalf.calfid and cowprof.herdid = rptcalf.herdid"
   Case 1
      SQL = "update cowbrd, rptcalf set comments = left(comments & ' " & txtUpdate & "', 250) where cowbrd.cowid = rptcalf.calfid and cowbrd.herdid = rptcalf.herdid and calfdate = #" & Left(cboyear.TEXT, 10) & "# "
End Select
DB.Execute SQL
DB.Close: Set DB = Nothing
Call DeleteTableAttachment(dbfile, "RPTCalf")
Unload Me
End Sub

Private Sub CmdTagAll_Click()
lstcow.SelectAll = 1
End Sub

Private Sub CmdUntag_Click()
lstcow.SelectAll = 0
End Sub

Private Sub Form_Load()
 CBOTab.AddItem "Profile"
 CBOTab.AddItem "Breeding"
 CBOTab.ListIndex = 0
 OptAct(0).Value = True
 Call centermdiform(Me, mdimain, 0, 0)
 Call LoadCows
 Call load_year
 Call load_cow_list(Me!lstcow, " where active = 'A'")
End Sub

Private Sub load_year()
Dim indx As Integer, INDX2 As Integer, OLDDATE$(), CurDate$
Screen.MousePointer = vbHourglass
cboyear.Clear
CurDate = ReturnBullTurnOutDate(herdid$, OLDDATE())
If CurDate = "" Then Exit Sub
Do Until indx = 5
    If OLDDATE(indx) = "--/--/----" Then
      cboyear.AddItem "Not Set"
    Else
      cboyear.AddItem OLDDATE(indx), indx
     End If
    indx = indx + 1
Loop
cboyear.AddItem CurDate, 0
If cboyear.ListCount > 0 Then cboyear.ListIndex = 0
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmsire_list = Nothing
End Sub

Private Sub lstcow_Click()
'
End Sub

Private Sub lstcow_DblClick()
  Call CmdEdit_Click
End Sub

Private Sub OptAct_Click(Index As Integer)
 Call LoadCows
End Sub
