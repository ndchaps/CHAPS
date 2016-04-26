VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form frmcow_list 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Cow Id"
   ClientHeight    =   3825
   ClientLeft      =   1905
   ClientTop       =   1980
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   4230
   Begin MhglbxLib.Mh3dList lstcow 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2880
      _Version        =   65536
      _ExtentX        =   5080
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "Cow Id"
      BackColor       =   16777215
      TintColor       =   16711935
      Caption         =   "Cow Id"
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
      ColTitle0       =   "Cow ID"
      ColWidth0       =   10
      ColTitle1       =   "Breed"
      ColWidth1       =   10
      ColTitle2       =   "Year"
      ColWidth2       =   15
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1275
      TabIndex        =   9
      Top             =   2820
      Width           =   1155
   End
   Begin VB.CheckBox chkPedigree 
      Caption         =   "Pedigree"
      Height          =   300
      Left            =   2295
      TabIndex        =   8
      Top             =   3270
      Width           =   1200
   End
   Begin VB.CheckBox chkculled 
      Caption         =   "Culled"
      Height          =   255
      Left            =   825
      TabIndex        =   7
      Top             =   3540
      Width           =   1230
   End
   Begin VB.CheckBox chkactive 
      Caption         =   "Active"
      Height          =   255
      Left            =   825
      TabIndex        =   6
      Top             =   3255
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change Herd"
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   3060
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   385
      Left            =   3060
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   385
      Left            =   3060
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   385
      Left            =   3060
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblSearch 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cow ID"
      Height          =   225
      Left            =   0
      TabIndex        =   10
      Top             =   2850
      Width           =   1185
   End
End
Attribute VB_Name = "frmcow_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub chkactive_Click()
Dim SQL$
  If chkactive Then
     SQL$ = SQL$ & "where ((active = 'A') "
     If chkculled Then SQL$ = SQL$ & "or (active = 'C' )"
     If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
     SQL$ = SQL$ & ")"
  Else
     If chkculled Then
       SQL$ = SQL$ & "where ((active = 'C' )"
       If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
       SQL$ = SQL$ & ")"
     Else
         If chkPedigree Then SQL$ = SQL$ & "where (active = 'P' )"
     End If
  End If
 Call load_cow_list(Me!lstcow, SQL$)

End Sub

Private Sub chkculled_Click()
  Dim SQL$
  If chkactive Then
     SQL$ = SQL$ & "where ((active = 'A') "
     If chkculled Then SQL$ = SQL$ & "or (active = 'C' )"
     If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
     SQL$ = SQL$ & ")"
  Else
     If chkculled Then
       SQL$ = SQL$ & "where ((active = 'C' )"
       If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
       SQL$ = SQL$ & ")"
     Else
         If chkPedigree Then SQL$ = SQL$ & "where (active = 'P' )"
     End If
  End If
 Call load_cow_list(Me!lstcow, SQL$)

End Sub

Private Sub chkPedigree_Click()
Dim SQL$
  If chkactive Then
     SQL$ = SQL$ & "where ((active = 'A') "
     If chkculled Then SQL$ = SQL$ & "or (active = 'C' )"
     If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
     SQL$ = SQL$ & ")"
  Else
     If chkculled Then
       SQL$ = SQL$ & "where ((active = 'C' )"
       If chkPedigree Then SQL$ = SQL$ & "or (active = 'P' )"
       SQL$ = SQL$ & ")"
     Else
       If chkPedigree Then SQL$ = SQL$ & "where (active = 'P' )"
     End If
  End If
 Call load_cow_list(Me!lstcow, SQL$)

End Sub

Private Sub CmdAdd_Click()
 Load frmcow_data
 frmcow_data.Tag = "A"
 frmcow_data.Show
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdchange_Click()
 selherd_List.Show vbModal
 If selherd_List.Tag = "CANCEL" Then Exit Sub
 herdid$ = selherd_List.Tag
 Unload Me
 Load frmcow_list

End Sub

Private Sub CmdDelete_Click()
  Dim theid$, iRet%
 iRet = MsgBox("Warning -- The Delete Option Is To Remove Mistaken Typing Entries.  If the Animal is No Longer In Production Change the Status To Cull", vbYesNo, Me.Caption)
   If iRet = vbYes Then
   Screen.MousePointer = vbHourglass
   lstcow.Col = 0
   theid$ = lstcow.ColText
   If frmcow_data!txtid.TEXT <> "" Then
   theid$ = frmcow_data!txtid.TEXT
    End If
   Load frmcow_data
   frmcow_data.Tag = "D/" & theid$
   frmcow_data.Show
End If
End Sub

Private Sub CmdEdit_Click()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcow.Col = 0
 theid$ = lstcow.ColText
 If frmcow_data!txtid.TEXT <> "" Then
  theid$ = frmcow_data!txtid.TEXT
 End If
 Load frmcow_data
 frmcow_data.Tag = "E/" & theid$
 frmcow_data.Show
End Sub


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 Call load_cow_list(Me!lstcow, " where active = 'A' ")
 frmcow_list.Caption = frmcow_list.Caption & " for Herd " & herdid$
If gIsDemo Then CMDDelete.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmcow_list = Nothing
End Sub





Private Sub lstcow_DblClick()
 Dim theid$
 Screen.MousePointer = vbHourglass
 lstcow.Col = 0
 theid$ = lstcow.ColText
 If frmcow_data!txtid.TEXT <> "" Then
  theid$ = frmcow_data!txtid.TEXT
 End If
 Load frmcow_data
 frmcow_data.Tag = "E/" & theid$
 frmcow_data.Show

End Sub


Private Sub txtSearch_Change()
  Dim Found As Integer
If Len(txtSearch.TEXT) = 0 Then Exit Sub
Found = Find_In_Listbox_Col_String(lstcow, 0, txtSearch.TEXT)
If Found <> -1 Then
  lstcow.TopIndex = Found
  lstcow.ListIndex = Found
 Else
  txtSearch.TEXT = Left$(txtSearch.TEXT, Len(txtSearch.TEXT) - 1)
  txtSearch.SelStart = Len(txtSearch.TEXT)
End If

End Sub


