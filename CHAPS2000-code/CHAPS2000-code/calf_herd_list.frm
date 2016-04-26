VERSION 4.00
Begin VB.Form FrmCalf_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Calf"
   ClientHeight    =   3240
   ClientLeft      =   1320
   ClientTop       =   2025
   ClientWidth     =   5640
   ClipControls    =   0   'False
   Height          =   3645
   Left            =   1260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   5640
   Top             =   1680
   Width           =   5760
   Begin VB.CommandButton cmdreports 
      Caption         =   "&Reports"
      Height          =   385
      Left            =   3450
      TabIndex        =   4
      Top             =   2760
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   4530
      TabIndex        =   3
      Top             =   2760
      Width           =   1000
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   385
      Left            =   2370
      TabIndex        =   2
      Top             =   2760
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   385
      Left            =   1305
      TabIndex        =   1
      Top             =   2760
      Width           =   1000
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   385
      Left            =   210
      TabIndex        =   0
      Top             =   2760
      Width           =   1000
   End
   Begin MhglbxLib.Mh3dList lstherd 
      Height          =   2535
      Left            =   30
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3285
      _Version        =   65536
      _ExtentX        =   5794
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
      Col             =   2
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
      ColTitle1       =   "Herd Name"
      ColWidth1       =   30
   End
   Begin MhglbxLib.Mh3dList lstcalf 
      Height          =   2535
      Left            =   3375
      TabIndex        =   5
      Top             =   120
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "Mh3dList1"
      BackColor       =   16777215
      TintColor       =   16711935
      Caption         =   "Mh3dList1"
      ColTitleButtons =   0   'False
      BevelStyleInner =   0
      BevelSizeInner  =   0
      BorderType      =   1
      BorderColor     =   0
      Case            =   0
      Col             =   2
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
      Sorted          =   0   'False
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
      SelectedColor   =   8421376
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
      ColTitle0       =   "Calf ID"
      ColWidth0       =   12
   End
End
Attribute VB_Name = "FrmCalf_List"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit




Private Sub CmdAdd_Click()
 Load frmCalf_Data
 frmCalf_Data.Tag = "A"
 frmCalf_Data.Show
End Sub

Private Sub cmdcancel_Click()
 lstcalf.Clear
 Unload Me
End Sub

Private Sub CmdDelete_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcalf.Col = 2
 theid$ = lstcalf.ColText
 Load frmCalf_Data
 frmCalf_Data.Tag = "D/" & theid$
End Sub

Private Sub CmdEdit_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcalf.Col = 2
 theid$ = lstcalf.ColText
 Load frmCalf_Data
 frmCalf_Data.Tag = "E/" & theid$
 'frmcalf_data.Show
End Sub


Private Sub Form_Activate()
   Me.caption = Me.caption & herdid$
End Sub

Private Sub form_Load()
 Call centermdiform(Me, mdichaps, 0, 0)
 ' load the herd list box
End Sub

Private Sub form_Unload(Cancel As Integer)
 Set FrmCalf_List = Nothing
End Sub

Private Sub lstcalf_DblClick()
Call CmdEdit_Click
End Sub






Private Sub lstherd_Click()
' load the calf list box
End Sub


