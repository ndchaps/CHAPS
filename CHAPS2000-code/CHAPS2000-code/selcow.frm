VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "mhlist32.ocx"
Begin VB.Form selcow_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Cow Id"
   ClientHeight    =   3135
   ClientLeft      =   1905
   ClientTop       =   1980
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   Begin MhglbxLib.Mh3dList lstcow 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
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
      ColTitle0       =   "Cow ID"
      ColWidth0       =   10
   End
   Begin VB.ComboBox cboyear 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1380
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Select"
      Height          =   385
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Cows that appear in RED already have one calf record entered for this breeding season."
      Height          =   1005
      Left            =   2595
      TabIndex        =   4
      Top             =   1815
      Width           =   1455
   End
End
Attribute VB_Name = "selcow_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mMode%


Property Let SetMode(pVal As Integer)
mMode = pVal
End Property

Private Sub LoadCows()
Dim DB As database, RS As Recordset, strSQL$, CurDate$
If cboyear.TEXT = "" Or cboyear.TEXT = "Not Set" Then Exit Sub
Screen.MousePointer = vbHourglass
lstcow.Clear
CurDate = Left(cboyear.TEXT, 10)
 strSQL$ = "SELECT DISTINCTROW cowprof.HerdID, cowprof.CowID, Sum(IIf([calfbirth].[birthdate]>=#" & CurDate & "# And [calfbirth].[birthdate]<=#" & CDate(CurDate) + 365 & "#,1,0)) AS Calf_Count FROM cowprof LEFT JOIN calfbirth ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID)  where cowprof.herdid = '" & herdid & "' and cowprof.active = 'A' "
 'strSQL$ = "SELECT DISTINCTROW cowprof.HerdID, cowprof.CowID, cowbrd.calfdate, Sum(IIf([calfbirth].[birthdate]>=#" & CurDate & "# And [calfbirth].[birthdate]<=#" & CDate(CurDate) + 365 & "#,1,0)) AS Calf_Count FROM (cowprof LEFT JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID)) LEFT JOIN calfbirth ON (cowprof.cowID = calfbirth.CowID) AND (cowprof.HerdID = calfbirth.HerdID) where cowprof.herdid = '" & herdid & "' "
' If cboyear.TEXT <> "Not Set" Then
'   strSQL = strSQL & " and cowbrd.calfdate = #" & CurDate & "# "
' End If
' strSQL = strSQL & " GROUP BY cowprof.HerdID, cowprof.CowID, cowbrd.calfdate, calfbirth.cowid ORDER BY calfbirth.CowID "
 strSQL = strSQL & " GROUP BY cowprof.HerdID, cowprof.cowID, calfbirth.CowID ORDER BY calfbirth.CowID"
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset(strSQL$)
 Do Until tbData.EOF
    Select Case tbData!calf_count
      Case Is = 0
         lstcow.TextColor = vbBlack
      Case Is >= 1
         lstcow.TextColor = vbRed
     End Select
    'lstcow.AddItem Field2Str(tbData!CowID)
    lstcow.AddItem tbData!CowID
    tbData.MoveNext
 Loop
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub


Private Sub cboyear_Click()
Call LoadCows
End Sub

Private Sub cmdcancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub

Private Sub CmdEdit_Click()
 lstcow.Col = 0
 Me.Tag = lstcow.ColText
 Me.Hide
End Sub

Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 Call LoadCows
 If mMode = 0 Then
   Call load_year
   cboyear.Visible = True
   Label1.Visible = True
 Else
   Call load_cow_list(Me!lstcow, " ")
   cboyear.Visible = False
   Label1.Visible = False
 End If
End Sub

Private Sub load_year()
Dim indx As Integer, INDX2 As Integer, OLDDATE$(), CurDate$
Screen.MousePointer = vbHourglass
cboyear.Clear
CurDate = ReturnBullTurnOutDate(herdid$, OLDDATE())
If CurDate = "" Then Exit Sub
Do Until indx = 5
    If OLDDATE(indx) = "--/--/----" Then
      cboyear.AddItem "Not Set"
    Else
      cboyear.AddItem OLDDATE(indx), indx
    End If
    indx = indx + 1
Loop
cboyear.AddItem CurDate & "*", 0
If cboyear.ListCount > 0 Then cboyear.ListIndex = 0
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmsire_list = Nothing
End Sub

Private Sub lstcow_DblClick()
  Call CmdEdit_Click
End Sub

