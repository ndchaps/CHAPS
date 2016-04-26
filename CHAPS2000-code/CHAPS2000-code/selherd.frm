VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form selherd_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Herd"
   ClientHeight    =   2925
   ClientLeft      =   780
   ClientTop       =   645
   ClientWidth     =   4650
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   4650
   Begin MhglbxLib.Mh3dList lstherd 
      Height          =   2655
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   3240
      _Version        =   65536
      _ExtentX        =   5715
      _ExtentY        =   4683
      _StockProps     =   79
      BackColor       =   16777215
      TintColor       =   16711935
      Caption         =   ""
      ColTitleButtons =   0   'False
      BevelStyleInner =   0
      BevelSizeInner  =   0
      BorderType      =   1
      BorderColor     =   0
      Case            =   0
      Col             =   1
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
      ColTitle0       =   "Herd ID"
      ColWidth0       =   12
      ColTitle1       =   "Herd Description"
      ColWidth1       =   30
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Select"
      Height          =   385
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "selherd_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%
Dim oldid$
Dim tbData As Recordset



Private Sub CmdAdd_Click()
 lstherd.Col = 0
 herdid = lstherd.ColText
 Me.Tag = lstherd.ColText
 mdimain.Caption = "Default Herd: " & herdid & "   C.H.A.P.S. (Herd Analysis)"
 Me.Hide
End Sub

Private Sub cmdcancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub




Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 'load the herd list box
 Call loadherd(Me!lstherd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmherd_List = Nothing
End Sub

Private Sub lstherd_DblClick()
Call CmdAdd_Click
End Sub






