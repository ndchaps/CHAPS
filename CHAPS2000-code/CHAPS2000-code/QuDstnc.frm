VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form frmListDistinct 
   Caption         =   "Distinct Value Selection"
   ClientHeight    =   3915
   ClientLeft      =   3150
   ClientTop       =   1965
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   3450
   Begin MhglbxLib.Mh3dList lstDistinct 
      Height          =   3750
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2265
      _Version        =   65536
      _ExtentX        =   3995
      _ExtentY        =   6615
      _StockProps     =   79
      Caption         =   "Mh3dList1"
      ForeColor       =   -2147483640
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
      ScrollBars      =   3
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
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   435
      Left            =   2355
      TabIndex        =   2
      Top             =   2535
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00000000&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   2370
      TabIndex        =   1
      Top             =   1410
      Width           =   1050
   End
End
Attribute VB_Name = "frmListDistinct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
Me.Tag = "Cancel"
Me.Hide
End Sub

Private Sub cmdselect_Click()
Me.Tag = lstDistinct.TEXT
If Trim(Me.Tag) = "" Then Me.Tag = "Cancel"
Me.Hide
End Sub

Private Sub Form_Load()
Dim SQL$, rs As Recordset
Dim dbpm As database
SQL$ = FrmHubQuery!txtDistinctSql.TEXT
Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set rs = dbpm.OpenRecordset(SQL$, dbOpenSnapshot)
Me.Tag = "Cancel"

While Not rs.EOF
  lstDistinct.AddItem Field2Str(rs!DV)
 rs.MoveNext
Wend

rs.Close: Set rs = Nothing
dbpm.Close: Set dbpm = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
  Cancel = True
  Me.Tag = "Cancel"
  Me.Hide
End If
End Sub

Private Sub lstDistinct_DblClick()
Me.Tag = lstDistinct.TEXT
If Trim(Me.Tag) = "" Then Me.Tag = "Cancel"
Me.Hide
End Sub

