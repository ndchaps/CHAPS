VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form FrmUpdate_Sire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Notes Field"
   ClientHeight    =   3180
   ClientLeft      =   360
   ClientTop       =   1740
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   5310
   Begin MhglbxLib.Mh3dList lstsire 
      Height          =   2535
      Left            =   15
      TabIndex        =   2
      Top             =   90
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "Sire Id"
      BackColor       =   16777215
      TintColor       =   16711935
      Caption         =   "Sire Id"
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
      ColTitle0       =   "Sire ID"
      ColWidth0       =   10
   End
   Begin VB.TextBox txtUpdate 
      Height          =   1215
      Left            =   2445
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1410
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   2430
      TabIndex        =   3
      Top             =   15
      Width           =   2760
      Begin VB.CommandButton CmdTagAll 
         Caption         =   "Tag All"
         Height          =   360
         Left            =   1455
         TabIndex        =   7
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton CmdUntag 
         Caption         =   "Untag All"
         Height          =   360
         Left            =   1455
         TabIndex        =   6
         Top             =   540
         Width           =   1245
      End
      Begin VB.Frame Frame2 
         Height          =   1050
         Left            =   45
         TabIndex        =   5
         Top             =   165
         Width           =   1350
         Begin VB.OptionButton OptAct 
            Caption         =   "Pedigree"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Top             =   750
            Width           =   1230
         End
         Begin VB.OptionButton OptAct 
            Caption         =   "Culled"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   10
            Top             =   465
            Width           =   1230
         End
         Begin VB.OptionButton OptAct 
            Caption         =   "Active"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   180
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "Change Herd"
         Height          =   345
         Left            =   1455
         TabIndex        =   4
         Top             =   930
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   4035
      TabIndex        =   1
      Top             =   2730
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Update"
      Height          =   385
      Left            =   2955
      TabIndex        =   0
      Top             =   2715
      Width           =   1000
   End
End
Attribute VB_Name = "FrmUpdate_Sire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCancel_Click()
Me.Tag = "CANCEL"
Me.Hide
End Sub

Private Sub cmdchange_Click()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
herdid$ = selherd_List.Tag
OptAct(0).Value = True
End Sub

Private Sub CmdEdit_Click()
Dim DB As DAO.database, SQL$, indx%
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
DB.Execute "delete * from rptcalf"
Do Until indx = lstsire.ListCount
   If lstsire.Tagged(indx) Then
      lstsire.Col = 0
      DB.Execute "insert into RPTCalf (CalfID, HerdID) Values ('" & lstsire.ColList(indx) & "', '" & herdid & "')"
   End If
   indx = indx + 1
Loop
SQL = "update sireprof, rptcalf set notes = LEFT(notes & ' " & txtUpdate & "', 250) where sireprof.sireid = rptcalf.calfid and sireprof.herdid = rptcalf.herdid"
DB.Execute SQL
DB.Close: Set DB = Nothing
Call DeleteTableAttachment(dbfile, "RPTCalf")
Unload Me
End Sub

Private Sub CmdTagAll_Click()
lstsire.SelectAll = 1
End Sub

Private Sub CmdUntag_Click()
lstsire.SelectAll = 0
End Sub

Private Sub Form_Load()
Call centermdiform(Me, mdimain, 0, 0)
'Call load_sire_list(Me!lstsire, " where active = 'A'")
OptAct(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmsire_list = Nothing
End Sub

Private Sub LSTSIRE_DblClick()
Call CmdEdit_Click
End Sub


Private Sub OptAct_Click(Index As Integer)
Select Case Index
   Case 0
         Call load_sire_list(Me!lstsire, " where active = 'A' and herdid = '" & herdid & "'")
   Case 1
         Call load_sire_list(Me!lstsire, " where active = 'C' and herdid = '" & herdid & "'")
   Case 2
         Call load_sire_list(Me!lstsire, " where active = 'P' and herdid = '" & herdid & "'")
End Select
End Sub
