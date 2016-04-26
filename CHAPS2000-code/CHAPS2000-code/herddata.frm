VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmherd_data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herd Information"
   ClientHeight    =   4095
   ClientLeft      =   3705
   ClientTop       =   2280
   ClientWidth     =   7800
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   7800
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   60
      TabIndex        =   14
      Top             =   75
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6271
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Herd Information"
      TabPicture(0)   =   "herddata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPremiseID"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtmisc2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtmisc1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtherddesc"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtregion"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtdistrict"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtcounty"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtcity"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtstate"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtzip"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtaddress"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtname"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtid"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtPremise"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Other"
      TabPicture(1)   =   "herddata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtPremise 
         Height          =   285
         Left            =   5115
         MaxLength       =   25
         TabIndex        =   29
         Top             =   2640
         Width           =   2460
      End
      Begin VB.TextBox txtid 
         Height          =   315
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtname 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox txtaddress 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1575
         Width           =   6135
      End
      Begin VB.TextBox txtzip 
         Height          =   315
         Left            =   6300
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1920
         Width           =   1275
      End
      Begin VB.TextBox txtstate 
         Height          =   315
         Left            =   5580
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtcity 
         Height          =   315
         Left            =   1425
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtcounty 
         Height          =   315
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2280
         Width           =   1755
      End
      Begin VB.TextBox txtdistrict 
         Height          =   315
         Left            =   3900
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtregion 
         Height          =   315
         Left            =   6060
         MaxLength       =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txtherddesc 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   6135
      End
      Begin VB.TextBox txtmisc1 
         Height          =   315
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   10
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtmisc2 
         Height          =   315
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label lblPremiseID 
         Caption         =   "Premise ID "
         Height          =   270
         Left            =   4305
         TabIndex        =   28
         Top             =   2685
         Width           =   1350
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Height          =   435
         Left            =   -74865
         TabIndex        =   27
         Top             =   390
         Width           =   3675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Herd ID"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   195
         Left            =   5880
         TabIndex        =   23
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   195
         Left            =   4980
         TabIndex        =   22
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "County"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   2340
         Width           =   555
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Herd Description"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comments 1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comments 2"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3060
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   385
      Left            =   5325
      TabIndex        =   12
      Top             =   3675
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   6540
      TabIndex        =   13
      Top             =   3690
      Width           =   1000
   End
End
Attribute VB_Name = "frmherd_data"
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

Private Sub Init_Information()
 Call init_form(Me) ' Clear Text Boxes
 ' load all combo boxes
  Call load_year
End Sub

Private Sub Load_information()
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 
 Set tbData = DB.OpenRecordset("herd", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", oldid$
 If Not tbData.NoMatch Then
   '  sought and found the data now put the data to the controls on the form
   If Not IsNull(tbData!herdid) Then txtid.TEXT = tbData!herdid
   If Not IsNull(tbData!herddesc) Then txtherddesc.TEXT = tbData!herddesc
   If Not IsNull(tbData!herdName) Then txtname.TEXT = tbData!herdName
   If Not IsNull(tbData!address) Then txtaddress.TEXT = tbData!address
   If Not IsNull(tbData!city) Then txtcity.TEXT = tbData!city
   If Not IsNull(tbData!state) Then txtstate.TEXT = tbData!state
   If Not IsNull(tbData!zip) Then txtzip.TEXT = tbData!zip
   If Not IsNull(tbData!county) Then txtcounty.TEXT = tbData!county
   If Not IsNull(tbData!district) Then txtdistrict.TEXT = tbData!district
   If Not IsNull(tbData!region) Then txtregion.TEXT = tbData!region
   If Not IsNull(tbData!misc1) Then txtmisc1.TEXT = tbData!misc1
   If Not IsNull(tbData!misc2) Then txtmisc2.TEXT = tbData!misc2
   If Not IsNull(tbData!Name) Then txtPremise.TEXT = tbData!Name
   
 End If
 Set tbData = DB.OpenRecordset("misc", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", "TurnDate" & Trim(Left(txtid.TEXT, 7))
 If Not tbData.NoMatch Then
   'lblWarning.Caption = lblWarning.Caption & tbData!thetext
   'lblWarning.Visible = True
   'Label16.Visible = True
   'txtTurnDate.TEXT = tbData!thetext
 Else
   'lblWarning.Caption = "Only One Bull Turn Out Date Is Allowed In A 285 Day Period"
   'Label16.Visible = False
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub save_information()
 Dim replace$, save As Boolean, iResponse As Integer, pOldDate As String
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("herd", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", oldid$
 save = True
 If Not tbData.NoMatch Then
   If addedflag$ = "D" Then
     tbData.Delete
     replace$ = ""
     save = False
    Else
     tbData.Edit
   End If
  Else
   tbData.AddNew
 End If
 If save Then
   With tbData
     !herdid = txtid.TEXT
     !herddesc = txtherddesc.TEXT
     !herdName = txtname.TEXT
     !address = txtaddress.TEXT
     !city = txtcity.TEXT
     !state = txtstate.TEXT
     !zip = txtzip.TEXT
     !county = txtcounty.TEXT
     !district = txtdistrict.TEXT
     !region = txtregion.TEXT
     !misc1 = txtmisc1.TEXT
     !misc2 = txtmisc2.TEXT
     !Name = txtPremise.TEXT
     .Update
     replace$ = txtid.TEXT & vbTab & txtname.TEXT
   End With
 End If
'If txtTurnDate.TEXT <> "--/--/----" Then
'      If Not IsDate(txtTurnDate.TEXT) Then
'         MsgBox "Please Enter A Valid Exposed Date", vbOKOnly + vbExclamation, Me.Caption
'         SSTab1.Tab = 1
'         txtTurnDate.SetFocus
'         Exit Sub
'      End If
'      Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
'      Set tbData = DB.OpenRecordset("misc", dbOpenTable)
'         tbData.Index = "primarykey"
'         tbData.Seek "=", "TurnDate" & txtid.TEXT
'      If Not tbData.NoMatch Then
'        If DateDiff("d", tbData!thetext, txtTurnDate.TEXT) < 285 Then
'             DB.Execute ("UPDATE DISTINCTROW cowprof INNER JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID) SET cowbrd.calfdate = #" & txtTurnDate.TEXT & "# WHERE (((cowprof.active)='A')) and cowbrd.calfdate = #" & tbData!thetext & "# and cowprof.herdid = '" & txtid.TEXT & "'")
'            Call SaveBullTurnOutDate(txtTurnDate.TEXT, txtid.TEXT, "E")
'        Else
'            DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate, conceptdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate, #01/01/1900# as conceptdate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & txtid.TEXT & "'", dbFailOnError
'            Call SaveBullTurnOutDate(txtTurnDate.TEXT, txtid.TEXT, "A")
'        End If
'      Else
'         DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate, conceptdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate, #01/01/1900# as conceptdate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & txtid.TEXT & "'", dbFailOnError
'         Call SaveBullTurnOutDate(txtTurnDate.TEXT, txtid.TEXT, "A")
'      End If
'   End If
CloseDB:
 tbData.Close: Set tbData = Nothing
 'delete bull turn out date from misc table
' If addedflag = "D" Then DB.Execute ("delete * from misc where thekey = 'TurnDate" & txtid.TEXT & "' and thekey Like 'Turn" & txtid.TEXT & "*'"), dbFailOnError
' DB.Close: Set DB = Nothing
 dirtyflag% = False
 Call Update_mh_ListBoxes("lstherd", 0, oldid$, replace$)
 Screen.MousePointer = vbDefault
End Sub


Private Sub valid_form(exitcode%)
    Dim iResponse As Integer, pCurDate As String
    exitcode% = 0
    If txtid.TEXT = "" Then
        Beep
        MsgBox "Herd ID Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
        txtid.SetFocus
        exitcode% = 1
        Exit Sub
    End If
    If UCase$(oldid$) <> UCase$(txtid.TEXT) Then
        Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
        Set tbData = DB.OpenRecordset("herd", dbOpenTable)
        tbData.Index = "primarykey"
        tbData.Seek "=", txtid.TEXT
        If Not tbData.NoMatch Then
            Beep
            MsgBox "Herd ID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
            exitcode% = 1
            Exit Sub
            tbData.Close: Set tbData = Nothing
            DB.Close: Set DB = Nothing
        End If
        tbData.Close: Set tbData = Nothing
        DB.Close: Set DB = Nothing
    End If
    'If txtTurnDate.TEXT <> "--/--/----" Then
    '  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
    '  Set tbData = DB.OpenRecordset("misc", dbOpenTable)
    '     tbData.Index = "primarykey"
    '     tbData.Seek "=", "Turn" & txtid.TEXT & "1"
    '  If Not tbData.NoMatch Then
    '     pCurDate = Field2Str(tbData!thetext)
    '     If pCurDate = "Not Set" Then GoTo Exit_Sub
    '     If CDate(txtTurnDate) < CDate(pCurDate) + 285 Then
    '        MsgBox "Warning: Your Entry Is Within 285 Days " & vbCrLf & " of the Existing Bull Turn Out Date.", vbOKOnly + vbCritical, Me.Caption
    '        exitcode% = 1
    '        tbData.Close: Set tbData = Nothing
    '        DB.Close: Set DB = Nothing
    '        Exit Sub
    '     End If
    '  End If
Exit_Sub:
   'tbData.Close: Set tbData = Nothing
   'DB.Close: Set DB = Nothing
'End If
End Sub

Private Sub cboyear_Click()
'If cboyear.TEXT <> "" And cboyear.TEXT <> "Not Set" Then txtNewDate.TEXT = cboyear.TEXT
End Sub

Private Sub cmdapply_Click()
Dim Msg$, UpperDate As Date, LowerDate As Date, PassedTest As Boolean
Dim DB As DAO.database
'If txtNewDate.TEXT = "--/--/----" Or cboyear.TEXT = "Not Set" Then Exit Sub
'Select Case cboyear.ListIndex
'   Case 0
'      If CDate(CurDate) Then UpperDate = CurDate
'      If OLDDATE(1) = "Not Set" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(1)
'   Case 1
'      If CDate(OLDDATE(0)) Then UpperDate = OLDDATE(0)
'      If OLDDATE(2) = "Not Set" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(2)
'   Case 2
'      If CDate(OLDDATE(1)) Then UpperDate = OLDDATE(1)
'      If OLDDATE(3) = "Not Set" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(3)
'   Case 3
'      If CDate(OLDDATE(2)) Then UpperDate = OLDDATE(2)
'      If OLDDATE(4) = "Not Set" Then LowerDate = #12:00:00 PM# Else LowerDate = OLDDATE(4)
'   Case 4
'      If CDate(OLDDATE(3)) Then UpperDate = OLDDATE(3)
'      LowerDate = #12:00:00 PM#
'End Select
'Failed_Test:
'If LowerDate = #12:00:00 PM# Then
'   If CDate(txtNewDate.TEXT) < UpperDate - 285 Then PassedTest = True Else PassedTest = False
'Else
'   If CDate(txtNewDate.TEXT) < UpperDate - 285 And CDate(txtNewDate.TEXT) > LowerDate + 285 Then PassedTest = True Else PassedTest = False
'End If
'If PassedTest Then
'   'update records here
'   Set DB = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%)
'   DB.Execute "update cowbrd set calfdate = #" & txtNewDate.TEXT & "# where herdid = '" & txtid.TEXT & "' and calfdate = #" & cboyear.TEXT & "#"
'   MsgBox DB.RecordsAffected & " Cow Conception/Breeding Records Updated", vbOKOnly, Me.Caption
'   DB.Execute "update misc set thetext = '" & txtNewDate.TEXT & "' where thekey = 'Turn" & txtid.TEXT & cboyear.ListIndex + 1 & "'"
'   DB.Close: Set DB = Nothing
'Else
'   MsgBox "Warning: This Date Is Within 285 Days" & vbCrLf & "From Previously Set Bull Turn Out Dates.", vbOKOnly + vbCritical, Me.Caption
'End If
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim TableName$(100)
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
 If addedflag$ = "D" Then
   Call CheckID(dbfile$, "herd", oldid$, TableName$())
   RESPONSE% = vbYes
   If Val(TableName$(0)) > 0 Then
     Beep
     RESPONSE% = MsgBox("Warning This Herd Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
   End If
   If RESPONSE% = vbYes Then Call save_information
 End If
 If addedflag$ <> "D" Then Call save_information
' If addedflag$ = "A" Then
'   Me.Tag = "A"
'   Call Form_Activate
'  Else
   Unload Me
' End If
End Sub

Private Sub Form_Activate()
 If Me.Tag = "" Then Exit Sub
 Dim tabchr
'  Label16.Visible = False
'  lblWarning.Visible = False
 addedflag$ = Left$(Me.Tag, 1)
 Screen.MousePointer = vbHourglass
 Call Init_Information
 If addedflag$ = "A" Then
   Me.Caption = "Add a Herd"
   oldid$ = ""
 End If
 If addedflag$ = "E" Or addedflag$ = "D" Then
   'tabchr = InStr(Me.Tag, Chr$(9))
   'If tabchr > 0 Then
      oldid$ = Trim$(Mid$(Me.Tag, 3))
      oldid$ = Trim$(oldid$)
   'End If
   Me.Caption = "Edit herd " & oldid$
   Call Load_information
   If addedflag$ = "D" Then
     Me.Caption = "Delete"
     Call disable_controls(Me)
     Cmdsave.Caption = "&Delete"
     Cmdsave.Enabled = True
     cmdcancel.Enabled = True
   End If
 End If
 Me.Tag = ""
 Screen.MousePointer = vbDefault
  
 Me.Enabled = True
 SSTab1.Tab = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
End Sub


Private Sub Form_Load()
Call centermdiform(Me, mdimain, 0, 0)
'Label15.Caption = "Selecting A Bull Turn Out Date Creates An Exposed Date Record On The Breeding/Conception Tab For All Active Cows." & vbCrLf & vbCrLf & "If Your Ranching Business Has More Than One Calving Season Per Year Create Only One Bull Turn Out Date Annually"

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
Dim indx As Integer, INDX2 As Integer
'Screen.MousePointer = vbHourglass
'cboyear.Clear
'CurDate = ReturnBullTurnOutDate(herdid$, OLDDATE())
'If CurDate = "" Then Exit Sub
'Do Until indx = 5
'    cboyear.AddItem OLDDATE(indx), indx
'    indx = indx + 1
'Loop
'If cboyear.ListCount > 0 Then cboyear.ListIndex = 0
'Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub SSCommand1_Click()
'gcaldate = txtNewDate.TEXT
'Call GetDate(gcaldate)
'txtNewDate.TEXT = gcaldate
End Sub

Private Sub SSCommand3_Click()
'gcaldate = txtTurnDate.TEXT
'Call GetDate(gcaldate)
'txtTurnDate.TEXT = gcaldate
End Sub
