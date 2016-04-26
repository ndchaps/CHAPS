VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPopUpCull 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin MhglbxLib.Mh3dList LstCullCodes 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   5001
      _ExtentY        =   4471
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
      MultiSelect     =   1
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
      ColTitle0       =   "Code"
      ColWidth0       =   7
      ColTitle1       =   "Description"
      ColWidth1       =   30
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3300
      TabIndex        =   6
      Top             =   2100
      Width           =   915
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   2100
      Width           =   975
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   270
      Left            =   3945
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   476
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "frmPopUpCull.frx":0000
   End
   Begin MSMask.MaskEdBox txtEndDate 
      Height          =   315
      Left            =   2940
      TabIndex        =   2
      Top             =   780
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   327681
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "-"
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   270
      Left            =   3945
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   476
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "frmPopUpCull.frx":04BE
   End
   Begin MSMask.MaskEdBox txtStartDate 
      Height          =   315
      Left            =   2940
      TabIndex        =   8
      Top             =   420
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   327681
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "-"
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   255
      Left            =   4380
      TabIndex        =   4
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   4380
      TabIndex        =   3
      Top             =   480
      Width           =   915
   End
End
Attribute VB_Name = "frmPopUpCull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Me.Hide
If LstCullCodes.SelectedCount > 0 Then Me.Tag = "True"
End Sub

Private Sub Form_Load()
With LstCullCodes
   .AddItem "G" & vbTab & "Cow Died"
   .AddItem "H" & vbTab & "Sold - Age"
   .AddItem "J" & vbTab & "Sold - Physical Defect"
   .AddItem "K" & vbTab & "Sold - Poor Fertility or Open"
   .AddItem "L" & vbTab & "Sold - Poor Performance"
   .AddItem "R" & vbTab & "Sold - Replacement Stock"
   .AddItem "Y" & vbTab & "Sold - Unknown"
   .ListIndex = 0
End With
End Sub

Private Sub SSCommand1_Click()
gcaldate = txtEndDate.TEXT
 Call GetDate(gcaldate)
 txtEndDate.TEXT = gcaldate
End Sub

Private Sub SSCommand2_Click()
gcaldate = txtStartDate.TEXT
 Call GetDate(gcaldate)
 txtStartDate.TEXT = gcaldate
End Sub

