VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form selcalf_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Calf Id"
   ClientHeight    =   3135
   ClientLeft      =   1905
   ClientTop       =   1980
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   Begin MhglbxLib.Mh3dList lstcalf 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   240
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
      MultiSelect     =   0
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
   Begin VB.Frame FraSex 
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   1260
      Width           =   1395
      Begin VB.OptionButton OptSex 
         Caption         =   "Steers"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   1155
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Heifers"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Bulls"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Misc"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Select"
      Height          =   385
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
End
Attribute VB_Name = "selcalf_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSex As Integer
Dim mStartDate As String
Dim mEndDate As String



Property Let SetSex(ByVal pSex As Integer)
mSex = pSex
End Property

Private Sub cmdcancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub

Private Sub CmdEdit_Click()
 lstcalf.Col = 0
 Me.Tag = lstcalf.ColText
 Me.Hide
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
 'If IsDate(date1) And IsDate(date2) Then Where$ = " where birthdate between #" & date1 & "# and #" & date2 & "# "
 mStartDate = date1
 mEndDate = date2
 OptSex(mSex).Value = True
 'Call load_calf_list(Me!lstcalf, Where$)
 End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmcalf_list = Nothing
End Sub

Private Sub lstcow_DblClick()
  Call CmdEdit_Click
End Sub

Private Sub OptSex_Click(Index As Integer)
Dim where$
where$ = " where birthdate between #" & mStartDate & "# and #" & mEndDate & "# "
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
