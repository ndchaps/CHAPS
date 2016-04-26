VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form selsire_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Sire Id"
   ClientHeight    =   2835
   ClientLeft      =   360
   ClientTop       =   1740
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   4095
   Begin VB.Frame FraType 
      Height          =   900
      Left            =   2625
      TabIndex        =   3
      Top             =   1845
      Width           =   1140
      Begin VB.OptionButton OptType 
         Caption         =   "Active"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Culled"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Pedigree"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   630
         Width           =   945
      End
   End
   Begin MhglbxLib.Mh3dList lstsire 
      Height          =   2505
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   4419
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
      ColTitle0       =   "Sire ID"
      ColWidth0       =   10
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Select"
      Height          =   385
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1000
   End
End
Attribute VB_Name = "selsire_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdcancel_Click()
 Me.Tag = "CANCEL"
Me.Hide
End Sub


Private Sub CmdEdit_Click()
 lstsire.Col = 0
 Me.Tag = lstsire.ColText
 Me.Hide
End Sub


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 Call load_sire_list(Me!lstsire, " where active = 'A'")
 OptType(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmsire_list = Nothing
End Sub



Private Sub LSTSIRE_DblClick()
 Call CmdEdit_Click

End Sub


Private Sub OptType_Click(Index As Integer)
Select Case Index
   Case 0
      Call load_sire_list(Me!lstsire, " where active = 'A'")
   Case 1
      Call load_sire_list(Me!lstsire, " where active = 'C'")
   Case 2
      Call load_sire_list(Me!lstsire, " where active = 'P'")
End Select
End Sub
