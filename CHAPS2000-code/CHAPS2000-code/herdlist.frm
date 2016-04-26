VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form frmherd_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Herd"
   ClientHeight    =   3585
   ClientLeft      =   2385
   ClientTop       =   2220
   ClientWidth     =   4650
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   4650
   Begin MhglbxLib.Mh3dList lstherd 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
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
   Begin VB.TextBox txtherdid 
      Height          =   285
      Left            =   1830
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3015
      Width           =   1525
   End
   Begin VB.CommandButton Cmdreports 
      Caption         =   "&Reports"
      Height          =   385
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1000
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   385
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   385
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   385
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Herd Default"
      Height          =   255
      Left            =   630
      TabIndex        =   7
      Top             =   3045
      Width           =   1095
   End
End
Attribute VB_Name = "frmherd_List"
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
 Load frmherd_data
 frmherd_data.Tag = "A"
 frmherd_data.Show
End Sub

Private Sub CMDCancel_Click()
 lstherd.Clear
 Unload Me
End Sub

Private Sub CmdDelete_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstherd.Col = 0
 theid$ = lstherd.ColText
 Load frmherd_data
 frmherd_data.Tag = "D/" & theid$
 frmherd_data.Show
End Sub

Private Sub CmdEdit_Click()
 Dim theid$
 lstherd.Col = 0
 theid$ = lstherd.ColText
 theid$ = Trim$(theid$)
 'lstherd.Col = 1
 'thename$ = lstherd.ColText
 'thename$ = Trim$(thename$)
 Load frmherd_data
 frmherd_data.Tag = "E/" & theid$
 frmherd_data.Show
End Sub


Private Sub Cmdreports_Click()
herdreps.Show
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 'load the herd list box
 Call loadherd(Me!lstherd)
txtherdid = ReadDefaultHerdID
If gIsDemo Then CmdDelete.Enabled = False Else CmdDelete.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmherd_List = Nothing
Call WriteDefaultHerdID(txtherdid)
End Sub

Private Sub lstherd_DblClick()
Call CmdEdit_Click
End Sub

Private Sub txtherdid_DblClick()
selherd_List.Show vbModal
If selherd_List.Tag = "CANCEL" Then Exit Sub
Call WriteDefaultHerdID(selherd_List.Tag)
txtherdid.TEXT = selherd_List.Tag
End Sub
