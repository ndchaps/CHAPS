VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmUpdate_Calf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Notes Field"
   ClientHeight    =   4155
   ClientLeft      =   1905
   ClientTop       =   1980
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4155
   ScaleWidth      =   5460
   Begin MhglbxLib.Mh3dList lstcalf 
      Height          =   2535
      Left            =   30
      TabIndex        =   2
      Top             =   15
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   4471
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
      ColTitle0       =   "CalfID"
      ColWidth0       =   10
   End
   Begin VB.TextBox txtUpdate 
      Height          =   1125
      Left            =   2505
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2580
      Width           =   2925
   End
   Begin VB.Frame FraSex 
      Height          =   2535
      Left            =   2505
      TabIndex        =   3
      Top             =   15
      Width           =   2925
      Begin VB.CommandButton CmdUntag 
         Caption         =   "Untag All"
         Height          =   375
         Left            =   1665
         TabIndex        =   17
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton CmdTagAll 
         Caption         =   "Tag All"
         Height          =   375
         Left            =   1665
         TabIndex        =   16
         Top             =   615
         Width           =   1095
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "Change Herd"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1425
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBOTab 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   225
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   90
         TabIndex        =   4
         Top             =   120
         Width           =   1335
         Begin VB.OptionButton OptSex 
            Caption         =   "Misc"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton OptSex 
            Caption         =   "Bulls"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   7
            Top             =   420
            Width           =   1155
         End
         Begin VB.OptionButton OptSex 
            Caption         =   "Heifers"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   660
            Width           =   1155
         End
         Begin VB.OptionButton OptSex 
            Caption         =   "Steers"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   5
            Top             =   900
            Width           =   1155
         End
      End
      Begin MSMask.MaskEdBox Dtestart 
         Height          =   285
         Left            =   1620
         TabIndex        =   9
         Top             =   1875
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
         Left            =   1620
         TabIndex        =   10
         Top             =   2190
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
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Birthdate"
         Height          =   255
         Index           =   15
         Left            =   165
         TabIndex        =   12
         Top             =   1875
         Width           =   1350
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Birthdate"
         Height          =   255
         Index           =   25
         Left            =   180
         TabIndex        =   11
         Top             =   2205
         Width           =   1350
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   4395
      TabIndex        =   1
      Top             =   3765
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Update"
      Height          =   385
      Left            =   3315
      TabIndex        =   0
      Top             =   3765
      Width           =   1000
   End
End
Attribute VB_Name = "FrmUpdate_Calf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSex As Integer

Property Let SetSex(ByVal pSex As Integer)
mSex = pSex
End Property

Private Sub CMDCancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub

Private Sub cmdchange_Click()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
herdid$ = selherd_List.Tag
OptSex(0).Value = True
End Sub

Private Sub CmdEdit_Click()
Dim DB As DAO.database, SQL$, indx%
Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
Call CreateTableAttachment(dbfile, repfile, "RPTCalf", "RPTCalf")
DB.Execute "delete * from rptcalf"
Do Until indx = lstcalf.ListCount
   If lstcalf.Tagged(indx) Then
      lstcalf.Col = 0
      DB.Execute "insert into RPTCalf (CalfID, HerdID) Values ('" & lstcalf.ColList(indx) & "', '" & herdid & "')"
   End If
   indx = indx + 1
Loop
Select Case CBOTab.ListIndex
   Case 0
      SQL = "update calfbirth, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calfbirth.calfid = rptcalf.calfid and calfbirth.herdid = rptcalf.herdid"
   Case 1
      SQL = "update calfwean, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calfwean.calfid = rptcalf.calfid and calfwean.herdid = rptcalf.herdid"
   Case 2
      SQL = "update calfback, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calfback.calfid = rptcalf.calfid and calfback.herdid = rptcalf.herdid"
   Case 3
      SQL = "update calfrep, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calfrep.calfid = rptcalf.calfid and calfrep.herdid = rptcalf.herdid"
   Case 4
      SQL = "update calffeed, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calffeed.calfid = rptcalf.calfid and calffeed.herdid = rptcalf.herdid"
   Case 5
      SQL = "update calfcarcass, rptcalf set notes = left(notes & ' " & txtUpdate & "', 250) where calfcarcass.calfid = rptcalf.calfid and calfcarcass.herdid = rptcalf.herdid"
End Select
DB.Execute SQL
DB.Close: Set DB = Nothing
Call DeleteTableAttachment(dbfile, "RPTCalf")
Unload Me
End Sub

Private Sub CmdTagAll_Click()
lstcalf.SelectAll = 1
End Sub

Private Sub CmdUntag_Click()
lstcalf.SelectAll = 0
End Sub

Private Sub Form_Load()
 Dim where$, date1 As String, date2 As String, tmpfile$, tmp%
 Call centermdiform(Me, mdimain, 0, 0)
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "Start date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then date1 = Left(tmpfile$, tmp%)
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "End date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then date2 = Left(tmpfile$, tmp%)
 Dtestart = date1
 dteend = date2
 OptSex(mSex).Value = True
 CBOTab.AddItem "Birth"
 CBOTab.AddItem "Weaning"
 CBOTab.AddItem "Background"
 CBOTab.AddItem "Replacement"
 CBOTab.AddItem "Feed Lot"
 CBOTab.AddItem "Carcass"
 CBOTab.ListIndex = 0
 End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmcalf_list = Nothing
End Sub

Private Sub lstcow_DblClick()
  Call CmdEdit_Click
End Sub

Private Sub OptSex_Click(Index As Integer)
Dim where$
where$ = " where birthdate between #" & Dtestart & "# and #" & dteend & "# and herdid = '" & herdid & "'"
Select Case Index
    Case 0
        where = where & " and sex = '0' "
    Case 1
        where = where & " and sex = '1' "
    Case 2
        where = where & " and sex = '2' "
    Case 3
        where = where & " and sex = '3' "
End Select
Call load_calf_list(Me!lstcalf, where$)
End Sub
