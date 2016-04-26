VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmcalf_list 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Calf Id"
   ClientHeight    =   3915
   ClientLeft      =   4065
   ClientTop       =   2685
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3915
   ScaleWidth      =   4095
   Begin LpLib.fpList lstcalf 
      Height          =   2625
      Left            =   90
      TabIndex        =   11
      Top             =   135
      Width           =   2595
      _Version        =   196608
      _ExtentX        =   4577
      _ExtentY        =   4630
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   2
      Sorted          =   1
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   0
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   1
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   0   'False
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "calflist.frx":0000
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1530
      TabIndex        =   9
      Top             =   2805
      Width           =   1155
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change Herd"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   390
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   385
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   385
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Dtestart 
      Height          =   285
      Left            =   1530
      TabIndex        =   5
      Top             =   3135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox dteend 
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   3450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "-"
   End
   Begin VB.Label lblSearch 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calf ID"
      Height          =   225
      Left            =   255
      TabIndex        =   10
      Top             =   2835
      Width           =   1185
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Birthdate"
      Height          =   255
      Index           =   25
      Left            =   90
      TabIndex        =   8
      Top             =   3465
      Width           =   1350
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Birthdate"
      Height          =   255
      Index           =   15
      Left            =   75
      TabIndex        =   7
      Top             =   3135
      Width           =   1350
   End
End
Attribute VB_Name = "frmcalf_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub CmdAdd_Click()
 Load frmCalf_Data
 frmCalf_Data.Tag = "A"
 frmCalf_Data.ListBoxSelection = lstcalf.ListIndex
 frmCalf_Data.Show
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
 selherd_List.Show vbModal
 If selherd_List.Tag = "CANCEL" Then Exit Sub
 herdid$ = selherd_List.Tag
 Unload Me
 Load frmcalf_list
End Sub

Private Sub CmdDelete_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcalf.Col = 0
 theid$ = lstcalf.ColText
 If frmCalf_Data!txtid.TEXT <> "" Then
  theid$ = frmCalf_Data!txtid.TEXT
 End If
 Load frmCalf_Data
 frmCalf_Data.Tag = "D/" & theid$
 frmCalf_Data.ListBoxSelection = lstcalf.ListIndex
 frmCalf_Data.Show
End Sub

Private Sub CmdEdit_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcalf.Col = 0
 theid$ = lstcalf.ColText
 If frmCalf_Data!txtid.TEXT <> "" Then
  theid$ = frmCalf_Data!txtid.TEXT
 End If
 Load frmCalf_Data
 frmCalf_Data.Tag = "E/" & theid$
 frmCalf_Data.ListBoxSelection = lstcalf.ListIndex
 frmCalf_Data.Show
End Sub


Private Sub dteend_LostFocus()
 Dim where$
 If IsDate(Dtestart.TEXT) And IsDate(dteend.TEXT) Then where$ = " where birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "#"
 Call load_calf_list(Me!lstcalf, where$)
End Sub


Private Sub Dtestart_LostFocus()
 Dim where$
 If IsDate(Dtestart.TEXT) And IsDate(dteend.TEXT) Then where$ = " where birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "#"
 Call load_calf_list(Me!lstcalf, where$)

End Sub


Private Sub Form_Load()
 Dim tmpfile$, tmp%, where$
 Call centermdiform(Me, mdimain, 0, 0)
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "Start date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then Dtestart.TEXT = Left(tmpfile$, tmp%)
 tmpfile$ = Space$(80)
 tmp% = GetPrivateProfileString("chaps", "End date", "", tmpfile$, Len(tmpfile$), "chaps.ini")
 If Left(tmpfile$, tmp%) <> "" Then dteend.TEXT = Left(tmpfile$, tmp%)
 If IsDate(Dtestart.TEXT) And IsDate(dteend.TEXT) Then where$ = " where birthdate between #" & Dtestart.TEXT & "# and #" & dteend.TEXT & "#"
 Call load_calf_list2(Me!lstcalf, where$)
 frmcalf_list.Caption = frmcalf_list.Caption & " for Herd " & herdid$
If gIsDemo Then CMDDelete.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim tmp%
 tmp% = WritePrivateProfileString("chaps", "Start date", Dtestart.TEXT, "chaps.ini")
 tmp% = WritePrivateProfileString("chaps", "End date", dteend.TEXT, "chaps.ini")
 Set frmcalf_list = Nothing
End Sub

Private Sub lstCust_DblClick()
Call CmdEdit_Click
End Sub



Private Sub LSTcalf_DblClick()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcalf.Col = 0
 theid$ = lstcalf.ColText
 If frmCalf_Data!txtid.TEXT <> "" Then
  theid$ = frmCalf_Data!txtid.TEXT
 End If
 Load frmCalf_Data
 frmCalf_Data.Tag = "E/" & theid$
 frmCalf_Data.ListBoxSelection = lstcalf.ListIndex
 frmCalf_Data.Show

End Sub



Private Sub txtSearch_Change()
  Dim Found As Integer
If Len(txtSearch.TEXT) = 0 Then Exit Sub
Found = Find_In_ListPro_Listbox_Col_StringLong(lstcalf, 0, txtSearch.TEXT)
If Found <> -1 Then
  lstcalf.TopIndex = Found
  lstcalf.ListIndex = Found
 Else
  txtSearch.TEXT = Left$(txtSearch.TEXT, Len(txtSearch.TEXT) - 1)
  txtSearch.SelStart = Len(txtSearch.TEXT)
End If

End Sub




