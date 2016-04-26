VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form FrmSelect_Multi_Calves 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Calves"
   ClientHeight    =   2820
   ClientLeft      =   3780
   ClientTop       =   3240
   ClientWidth     =   4800
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin LpLib.fpList lstCalves 
      Height          =   2625
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   3450
      _Version        =   196608
      _ExtentX        =   6085
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
      MultiSelect     =   1
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
      ColDesigner     =   "SelMCalf.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Tag All"
      Height          =   360
      Left            =   3720
      TabIndex        =   5
      Top             =   510
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Untag All"
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   915
      Width           =   885
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "&Done"
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lbltagged 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3750
      TabIndex        =   2
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tagged"
      Height          =   210
      Left            =   3735
      TabIndex        =   1
      Top             =   1305
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSelect_Multi_Calves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveletter As String

Public Sub LoadCalves(ListBox As Control, pType$)
Dim DB As database, RS As Recordset, strSQL$
Screen.MousePointer = vbHourglass
 strSQL$ = "select calfbirth.cowid, calfbirth.herdid, calfbirth.calfid from calfbirth where calfbirth.herdid = '" & herdid & "'"
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset(strSQL$)
 ListBox.Clear
 Do Until tbData.EOF
    'mhListBox.AddItem Field2Str(tbData!CalfID) & Chr$(9) & Field2Str(tbData!herdid)
    ListBox.InsertRow = Field2Str(tbData!calfid) & Chr$(9) & Field2Str(tbData!herdid)
    tbData.MoveNext
 Loop
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub


Private Sub CMDCancel_Click()
 lstCalves.Clear
 Me.Tag = ""
 Me.Hide
End Sub

Private Sub cmdselect_Click()
 lstCalves.Col = 1
 Me.Tag = lstCalves.ColText
 Me.Hide
End Sub

Private Sub CmdDone_Click()
  lstCalves.Col = 1
  Me.Tag = lstCalves.ColText
  Me.Hide
End Sub

Private Sub Command1_Click()
'lstCalves.SelectAll = 0
'lbltagged.Caption = ""
lstCalves.Action = ActionDeselectAll
lbltagged.Caption = lstCalves.SelCount
End Sub

Property Let SetMode(pMode%)
Select Case pMode
   Case 0
            
   Case 1
      lbltagged.Visible = False
      label1.Visible = False
      Command1.Visible = False
      lstCalves.MultiSelect = mhSelectExtended
End Select
End Property


Private Sub Command2_Click()
'lstCalves.SelectAll = 1
'lbltagged.Caption = lstCalves.ListCount
 lstCalves.Action = ActionSelectAll
 lbltagged.Caption = lstCalves.SelCount
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
  Call LoadCalves(lstCalves, "Active")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Set FrmSelect_Multi_Cows = Nothing
End Sub


Private Sub lstCalves_Click()
 'lbltagged.Caption = Trim$(Str$(lstCalves.SelectedCount))
  lbltagged.Caption = lstCalves.SelCount
End Sub

Private Sub OptType_Click(Index As Integer)
lbltagged.Caption = ""
Select Case Index
   Case 0
       Call LoadCalves(lstCalves, "Active")
   Case 1
       Call LoadCalves(lstCalves, "Culled")
   Case 2
       Call LoadCalves(lstCalves, "Pedigree")
End Select
End Sub

