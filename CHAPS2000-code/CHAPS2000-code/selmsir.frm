VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form FrmSelect_Multi_Sires 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Sires"
   ClientHeight    =   2790
   ClientLeft      =   3855
   ClientTop       =   870
   ClientWidth     =   4800
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
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
      TabIndex        =   10
      Top             =   495
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   90
      TabIndex        =   8
      Top             =   30
      Width           =   3435
      Begin MhglbxLib.Mh3dList lstSires 
         Height          =   2655
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   3420
         _Version        =   65536
         _ExtentX        =   6032
         _ExtentY        =   4683
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         ColTitle0       =   "Sire ID"
         ColWidth0       =   12
         ColTitle1       =   "Herd Name"
         ColWidth1       =   30
      End
   End
   Begin VB.Frame FraType 
      Height          =   900
      Left            =   3600
      TabIndex        =   4
      Top             =   1830
      Width           =   1140
      Begin VB.OptionButton OptType 
         Caption         =   "Pedigree"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Culled"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   405
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Active"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Untag All"
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   885
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
      Top             =   2745
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label lbltagged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3750
      TabIndex        =   2
      Top             =   1545
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSelect_Multi_Sires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveletter As String

Private Sub cmdcancel_Click()
 lstSires.Clear
 Me.Tag = ""
 Me.Hide
End Sub

Private Sub CmdSelect_Click()
 lstSires.Col = 1
 Me.Tag = lstSires.ColText
 Me.Hide
End Sub

Private Sub CmdDone_Click()
' lstCows.Col = 1
 Me.Tag = lstSires.ColText
 Me.Hide
End Sub

Private Sub Command1_Click()
lstSires.SelectAll = 0
lbltagged.Caption = ""
End Sub

Property Let SetMode(pMode%)
Select Case pMode
   Case 0
            
   Case 1
      lbltagged.Visible = False
      'label1.Visible = False
      Command1.Visible = False
      lstSires.MultiSelect = mhSelectExtended
End Select
End Property


Private Sub Command2_Click()
lstSires.SelectAll = 1
lbltagged.Caption = lstSires.ListCount
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 OptType(0).Value = True
 'Call LoadCows(lstCows, "Active")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Set FrmSelect_Multi_Cows = Nothing
End Sub


Private Sub lstCows_Click()
 lbltagged.Caption = Trim$(Str$(lstSires.SelectedCount))
End Sub

Private Sub lstSires_Click()
lbltagged.Caption = Trim$(Str$(lstSires.SelectedCount))
End Sub

Private Sub OptType_Click(Index As Integer)
lbltagged.Caption = ""
Select Case Index
   Case 0
       Call LoadSires("Active")
   Case 1
       Call LoadSires("Culled")
   Case 2
       Call LoadSires("Pedigree")
End Select
End Sub

Private Sub LoadSires(pType$)
Dim DB As database, RS As Recordset, strSQL$
Screen.MousePointer = vbHourglass
If pType = "Active" Then pType = " and sireprof.active = 'A' "
If pType = "Culled" Then pType = " and sireprof.active = 'C' "
If pType = "Pedigree" Then pType = " and sireprof.active = 'P' "
strSQL$ = "select sireid, herdid from sireprof where herdid = '" & herdid & "'" & pType
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set tbData = DB.OpenRecordset(strSQL$)
lstSires.Clear
Do Until tbData.EOF
    lstSires.AddItem Field2Str(tbData!sireid) & Chr$(9) & Field2Str(tbData!herdid)
    tbData.MoveNext
Loop
tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub
