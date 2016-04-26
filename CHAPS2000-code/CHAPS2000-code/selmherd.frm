VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form FrmSelect_Multi_Herds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Herds"
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
   Begin MhglbxLib.Mh3dList lstherd 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
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
      ColTitle0       =   "Herd ID"
      ColWidth0       =   12
      ColTitle1       =   "Herd Name"
      ColWidth1       =   30
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "&Done"
      Height          =   385
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3750
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label lbltagged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3720
      TabIndex        =   4
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tagged"
      Height          =   240
      Left            =   3720
      TabIndex        =   3
      Top             =   555
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSelect_Multi_Herds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveletter As String

Private Sub cmdcancel_Click()
 lstherd.Clear
 Me.Tag = ""
 Me.Hide
End Sub

Private Sub CmdSelect_Click()
 lstherd.Col = 1
 Me.Tag = lstherd.ColText
 Me.Hide
End Sub

Private Sub CmdDone_Click()
Dim Indx%
Me.Tag = lstherd.ColText
Do Until Indx = lstherd.ListCount
   If lstherd.Tagged(Indx) = True Then herdid = lstherd.ColList(Indx): Exit Do
   Indx = Indx + 1
Loop
Me.Hide
End Sub

Private Sub Form_Load()
 Dim Indx%
 Call centermdiform(Me, mdimain, 0, 0)
 Call loadherd(Me!lstherd)
 If herdid <> "" Then
   lstherd.Col = 0
   Do Until Indx = lstherd.ListCount
      If lstherd.ColList(Indx) = herdid Then
         lstherd.Tagged(Indx) = True
         Exit Do
      End If
      Indx = Indx + 1
   Loop
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set FrmSelect_Multi_Herds = Nothing
End Sub


Private Sub lstherd_Click()
 lbltagged.Caption = Trim$(Str$(lstherd.SelectedCount))
End Sub

Private Sub txtlook_Change()

'Dim SAVELOOK$
'If Left$(txtlook.Text, 1) <> saveletter Then
' SAVELOOK$ = txtlook.Text
' Call Load_Vendor_List(Me!lstvend, txtlook.Text)
' txtlook.Text = SAVELOOK$
 'saveletter = Left$(txtlook.Text, 1)
 ''If lstvend.ListCount < 1 Then
'   txtlook.Text = ""
' End If
' Exit Sub
'End If
'lstvend.ColInstr = 0
'lstvend.FoundIndex = -1
'lstvend.FindInstr = txtlook.Text
'If lstvend.FoundIndex <> -1 Then
'  lstvend.ListIndex = lstvend.FoundIndex
' Else
'  If Len(txtlook.Text) Then
'    txtlook.Text = Left$(txtlook.Text, Len(txtlook.Text) - 1)
'    txtlook.SelStart = Len(txtlook.Text) + 1
'   Else
''    txtlook.Text = ""
'  End If
' End If
'  saveletter = Left$(txtlook.Text, 1)
'
End Sub


