VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTurnDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Bull Turn Out Date"
   ClientHeight    =   4095
   ClientLeft      =   3300
   ClientTop       =   1815
   ClientWidth     =   7710
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdenable 
      Caption         =   "Edit"
      Height          =   360
      Left            =   2940
      TabIndex        =   20
      Top             =   3645
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Bull Turn Out Date"
      Height          =   3120
      Left            =   4110
      TabIndex        =   10
      Top             =   135
      Width           =   3450
      Begin VB.ComboBox CBOCurrHerd 
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   345
         Width           =   1275
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   270
         Left            =   2790
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   765
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "turndate.frx":0000
      End
      Begin MSMask.MaskEdBox txtTurnDate 
         Height          =   315
         Left            =   1785
         TabIndex        =   12
         Top             =   735
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Herd"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   420
         Width           =   1530
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Height          =   1350
         Left            =   90
         TabIndex        =   15
         Top             =   1740
         Width           =   3195
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "Only One Bull Turn Out Date Is Allowed In A 285 Day Period."
         Height          =   465
         Left            =   195
         TabIndex        =   14
         Top             =   1215
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Bull Turn Out Date"
         Height          =   315
         Left            =   225
         TabIndex        =   13
         Top             =   795
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Edit Previous Bull Turn Out Dates"
      Height          =   3105
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   3630
      Begin VB.ComboBox CBOPrevHerd 
         Height          =   315
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   420
         Width           =   1275
      End
      Begin VB.ComboBox cboyear 
         Height          =   315
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "Save"
         Height          =   330
         Left            =   2895
         TabIndex        =   3
         Top             =   2685
         Width           =   645
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   270
         Left            =   3090
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1230
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "turndate.frx":04BE
      End
      Begin MSMask.MaskEdBox txtNewDate 
         Height          =   315
         Left            =   2085
         TabIndex        =   6
         Top             =   1200
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Herd"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   495
         Width           =   1890
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Previous Turn Out Dates"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   9
         Top             =   855
         Width           =   1890
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Edit Bull Turn Out Date"
         Height          =   315
         Left            =   165
         TabIndex        =   8
         Top             =   1230
         Width           =   1770
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Height          =   465
         Left            =   255
         TabIndex        =   7
         Top             =   1335
         Width           =   3135
      End
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   385
      Left            =   5340
      TabIndex        =   0
      Top             =   3660
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   6540
      TabIndex        =   1
      Top             =   3660
      Width           =   1000
   End
End
Attribute VB_Name = "FrmTurnDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%
Dim oldid$
Dim tbData As Recordset
Dim OLDDATE$(), CurDate$
Dim mFormLoad As Boolean

Private Sub Init_Information()
Call init_form(Me) ' Clear Text Boxes

End Sub

Private Sub LoadCBOHERD(CBO As ComboBox)
Dim DB As DAO.database, RS As DAO.Recordset
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set RS = DB.OpenRecordset("select HerdID from herd order by herdid", dbOpenSnapshot)
CBO.Clear
Do Until RS.EOF
   CBO.AddItem Field2Str(RS!herdid)
   RS.MoveNext
Loop
If CBO.ListCount > 0 Then CBO.ListIndex = 0
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub Load_information()
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set tbData = DB.OpenRecordset("turnoutdate", dbOpenTable)
tbData.Index = "PrimaryKey"
tbData.Seek "=", CBOCurrHerd.TEXT
If Not tbData.NoMatch Then
   txtTurnDate.TEXT = Field2Date(tbData!currentdate)
Else
   txtTurnDate.TEXT = Format(Now, "mm/dd/yyyy")
   If mFormLoad = False Then MsgBox "This Herd Does Not Have A Default Bull Turn Out Date.", vbOKOnly, Me.Caption
End If
tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub save_information()
Dim replace$, save As Boolean, iResponse As Integer, pOldDate As String
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set tbData = DB.OpenRecordset("TurnOutDate", dbOpenTable)
tbData.Index = "primarykey"
tbData.Seek "=", CBOCurrHerd.TEXT
If txtTurnDate.TEXT <> "--/--/----" Then
      If Not tbData.NoMatch Then
         If DateDiff("d", tbData!date1, txtTurnDate.TEXT) < 285 Then
            MsgBox "Current Bull Turn Out Date Can't Be Within 285 Days of Previous Years Bull Turn Out Date", vbOKOnly
            tbData.Close: Set tbData = Nothing
            DB.Close: Set DB = Nothing
            Exit Sub
         End If
      End If
      If Not IsDate(txtTurnDate.TEXT) Then
         MsgBox "Please Enter A Valid Exposed Date", vbOKOnly + vbExclamation, Me.Caption
         txtTurnDate.SetFocus
         Exit Sub
      End If
      Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
      Set tbData = DB.OpenRecordset("turnoutdate", dbOpenTable)
         tbData.Index = "primarykey"
         tbData.Seek "=", CBOCurrHerd.TEXT
      If Not tbData.NoMatch Then
        If DateDiff("d", tbData!currentdate, txtTurnDate.TEXT) < 285 Then
            DB.Execute "UPDATE DISTINCTROW cowprof INNER JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID) SET cowbrd.calfdate = #" & txtTurnDate.TEXT & "# WHERE cowbrd.calfdate = #" & tbData!currentdate & "# and cowprof.herdid = '" & CBOCurrHerd.TEXT & "'"
            DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & CBOCurrHerd.TEXT & "'"
            Call SaveBullTurnOutDate(txtTurnDate.TEXT, CBOCurrHerd.TEXT, "E")
        Else
            DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & CBOCurrHerd.TEXT & "'"
            Call SaveBullTurnOutDate(txtTurnDate.TEXT, CBOCurrHerd.TEXT, "A")
        End If
      Else
         DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & CBOCurrHerd.TEXT & "'"
         Call SaveBullTurnOutDate(txtTurnDate.TEXT, CBOCurrHerd.TEXT, "A")
      End If
   End If
CloseDB:
tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
dirtyflag% = False
Screen.MousePointer = vbDefault
End Sub


Private Sub valid_form(exitcode%)
    Dim iResponse As Integer, pCurDate As String
    exitcode% = 0
    
Exit_Sub:
   '
End Sub

Private Sub CBOCurrHerd_Click()
Call Load_information
End Sub

Private Sub CBOPrevHerd_Click()
Dim indx As Integer, INDX2 As Integer
cboyear.Clear
CurDate = ReturnBullTurnOutDate(CBOPrevHerd.TEXT, OLDDATE())
If CurDate = "" Then Exit Sub
Do Until indx = 5
   If OLDDATE(indx) <> "--/--/----" Then
      cboyear.AddItem OLDDATE(indx), indx
   Else
      cboyear.AddItem "Not Set"
   End If
   indx = indx + 1
Loop
If cboyear.ListCount > 0 Then cboyear.ListIndex = 0
End Sub

Private Sub cboyear_Click()
If cboyear.TEXT <> "" And cboyear.TEXT <> "Not Set" Then txtNewDate.TEXT = cboyear.TEXT
End Sub

Private Sub cmdapply_Click()
Dim Msg$, UpperDate As Date, LowerDate As Date, PassedTest As Boolean
Dim DB As DAO.database
If txtNewDate.TEXT = "--/--/----" Or cboyear.TEXT = "Not Set" Then Exit Sub
Select Case cboyear.ListIndex
   Case 0
      If CDate(CurDate) Then UpperDate = CurDate
      If OLDDATE(1) = "--/--/----" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(1)
   Case 1
      If CDate(OLDDATE(0)) Then UpperDate = OLDDATE(0)
      If OLDDATE(2) = "--/--/----" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(2)
   Case 2
      If CDate(OLDDATE(1)) Then UpperDate = OLDDATE(1)
      If OLDDATE(3) = "--/--/----" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(3)
   Case 3
      If CDate(OLDDATE(2)) Then UpperDate = OLDDATE(2)
      If OLDDATE(4) = "--/--/----" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(4)
   Case 4
      If CDate(OLDDATE(3)) Then UpperDate = OLDDATE(3)
      If OLDDATE(5) = "--/--/----" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(5)
   Case 5
      If CDate(OLDDATE(4)) Then UpperDate = OLDDATE(4)
      LowerDate = #12:00:00 PM#
End Select
Failed_Test:
If LowerDate = #12:00:00 PM# Then
   If CDate(txtNewDate.TEXT) < UpperDate - 285 Then PassedTest = True Else PassedTest = False
Else
   If CDate(txtNewDate.TEXT) < UpperDate - 285 And CDate(txtNewDate.TEXT) > LowerDate + 285 Then PassedTest = True Else PassedTest = False
End If
If PassedTest Then
   'update records here
   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
   DB.Execute "update cowbrd set calfdate = #" & txtNewDate.TEXT & "# where herdid = '" & CBOPrevHerd.TEXT & "' and calfdate = #" & cboyear.TEXT & "#"
   MsgBox DB.RecordsAffected & " Cow Conception/Breeding Records Updated", vbOKOnly, Me.Caption
   DB.Execute "update turnoutdate set date" & cboyear.ListIndex + 1 & "  = #" & txtNewDate.TEXT & "# where herdid = '" & CBOPrevHerd.TEXT & "'"
   DB.Close: Set DB = Nothing
Else
   MsgBox "Warning: This Date Is Within 285 Days" & vbCrLf & "From Previously Set Bull Turn Out Dates.", vbOKOnly + vbCritical, Me.Caption
End If
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub cmdenable_Click()

 Load frmpassword_entry
 
 frmpassword_entry.Tag = BuildChapsPassword
 frmpassword_entry.Show vbModal
 If frmpassword_entry.Tag = "CANCEL" Then
   Unload frmpassword_entry
   Exit Sub
 End If
 Unload frmpassword_entry

  MsgBox "Editing Bull Turn Out Dates is not recommended.  There should be only one date for each calving season.", vbExclamation, "Warning"
  Frame2.Visible = True
  CmdApply.Enabled = True
  Label17.Enabled = True
  Label13(0).Enabled = True
  Label13(1).Enabled = True
  CBOPrevHerd.Enabled = True
  cboyear.Enabled = True
  txtNewDate.Enabled = True
  SSCommand1.Enabled = True
End Sub

Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim TableName$(100)
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
 'If addedflag$ = "D" Then
 '  Call CheckID(dbfile$, "herd", oldid$, TableName$())
 '  RESPONSE% = vbYes
 '  If Val(TableName$(0)) > 0 Then
 '    Beep
 '    RESPONSE% = MsgBox("Warning This Herd Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
 '  End If
 '  If RESPONSE% = vbYes Then Call save_information
 'End If
Call save_information
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
End Sub


Private Sub Form_Load()
Call centermdiform(Me, mdimain, 0, 0)
Label15.Caption = "Selecting A Bull Turn Out Date Creates An Exposed Date Record On The Breeding/Conception Tab For All Active Cows.  If Your Ranching Business Has More Than One Calving Season Per Year Create Only One Bull Turn Out Date Annually"
mFormLoad = True
Call LoadCBOHERD(CBOPrevHerd)
Call LoadCBOHERD(CBOCurrHerd)
Call Load_information
mFormLoad = False
'txtNewDate.TEXT = Format(Now, "mm/dd/yyyy")
  CmdApply.Enabled = False
  Label17.Enabled = False
  Label13(0).Enabled = False
  Label13(1).Enabled = False
  CBOPrevHerd.Enabled = False
  cboyear.Enabled = False
  txtNewDate.Enabled = False
  SSCommand1.Enabled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ' see if infomation needs to be saveed
 Dim RESPONSE%, exitcode%
 If dirtyflag% Then
   Beep
   RESPONSE% = MsgBox("Information Has Been Changed" & vbCrLf & " Do You Wish To Save?", vbYesNoCancel + vbQuestion, Me.Caption)
   Select Case RESPONSE%
    Case vbYes
     Call valid_form(exitcode%)
     If exitcode% <> 0 Then
       Cancel = True
     End If
     Call save_information
    Case vbCancel
     Cancel = True
   End Select
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Cancel Then Exit Sub
 Set frmherd_data = Nothing
End Sub

Private Sub load_year()
'
End Sub

Private Sub SSCommand1_Click()
gcaldate = txtNewDate.TEXT
Call GetDate(gcaldate)
txtNewDate.TEXT = gcaldate
End Sub

Private Sub SSCommand3_Click()
gcaldate = txtTurnDate.TEXT
Call GetDate(gcaldate)
txtTurnDate.TEXT = gcaldate
End Sub
