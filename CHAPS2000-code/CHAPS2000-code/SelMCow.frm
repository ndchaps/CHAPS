VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form FrmSelect_Multi_Cows 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Cows"
   ClientHeight    =   2790
   ClientLeft      =   3855
   ClientTop       =   870
   ClientWidth     =   4800
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Tag All"
      Height          =   360
      Left            =   3720
      TabIndex        =   11
      Top             =   510
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Width           =   3435
      Begin MhglbxLib.Mh3dList lstCows 
         Height          =   2655
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   3420
         _Version        =   65536
         _ExtentX        =   6032
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
         Col             =   0
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
         ColTitle0       =   "Cow ID"
         ColWidth0       =   12
         ColTitle1       =   "Herd Name"
         ColWidth1       =   30
      End
   End
   Begin VB.Frame FraType 
      Height          =   900
      Left            =   3600
      TabIndex        =   5
      Top             =   1845
      Width           =   1140
      Begin VB.OptionButton OptType 
         Caption         =   "Pedigree"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Culled"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   405
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Active"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Untag All"
      Height          =   360
      Left            =   3720
      TabIndex        =   4
      Top             =   915
      Width           =   885
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "&Done"
      Height          =   345
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3750
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label lbltagged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3750
      TabIndex        =   3
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tagged"
      Height          =   210
      Left            =   3735
      TabIndex        =   2
      Top             =   1305
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSelect_Multi_Cows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveletter As String

Private Sub cmdcancel_Click()
 lstCows.Clear
 Me.Tag = ""
 Me.Hide
End Sub

Private Sub CmdSelect_Click()
 lstCows.Col = 1
 Me.Tag = lstCows.ColText
 Me.Hide
End Sub

Private Sub CmdDone_Click()
' lstCows.Col = 1
 Me.Tag = lstCows.ColText
 Me.Hide
End Sub

Private Sub Command1_Click()
lstCows.SelectAll = 0
lbltagged.Caption = ""
End Sub

Property Let SetMode(pMode%)
Select Case pMode
   Case 0
            
   Case 1
      lbltagged.Visible = False
      Label1.Visible = False
      Command1.Visible = False
      lstCows.MultiSelect = mhSelectExtended
End Select
End Property


Private Sub Command2_Click()
lstCows.SelectAll = 1
lbltagged.Caption = lstCows.ListCount
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 OptType(0).Value = True
 Call LoadCows(lstCows, "Active")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Set FrmSelect_Multi_Cows = Nothing
End Sub


Private Sub lstCows_Click()
 lbltagged.Caption = Trim$(Str$(lstCows.SelectedCount))
End Sub

Private Sub OptType_Click(Index As Integer)
lbltagged.Caption = ""
Select Case Index
   Case 0
       Call LoadCows(lstCows, "Active")
   Case 1
       Call LoadCows(lstCows, "Culled")
   Case 2
       Call LoadCows(lstCows, "Pedigree")
End Select
End Sub
