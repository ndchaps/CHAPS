VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm mdimain 
   BackColor       =   &H8000000C&
   Caption         =   "C.H.A.P.S. (Herd Analysis)"
   ClientHeight    =   3255
   ClientLeft      =   2865
   ClientTop       =   2070
   ClientWidth     =   6690
   Icon            =   "mdimain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdimain.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDG 
      Left            =   7770
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6690
      _Version        =   65536
      _ExtentX        =   11800
      _ExtentY        =   714
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel pnlhint 
         Height          =   375
         Left            =   3595
         TabIndex        =   1
         Top             =   15
         Width           =   6020
         _Version        =   65536
         _ExtentX        =   10619
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "mdimain.frx":C775
      End
      Begin Threed.SSCommand cmdopen 
         Height          =   375
         Left            =   345
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "mdimain.frx":C887
      End
      Begin Threed.SSCommand CMDSetup 
         Height          =   375
         Left            =   1485
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "mdimain.frx":C981
      End
      Begin Threed.SSCommand CMDForms 
         Height          =   375
         Left            =   1845
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "mdimain.frx":CDD3
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileopen 
         Caption         =   "&Open"
         Begin VB.Menu mnucalves 
            Caption         =   "C&alves"
         End
         Begin VB.Menu mnuopencows 
            Caption         =   "&Cows"
         End
         Begin VB.Menu mnuopensire 
            Caption         =   "&Sires"
         End
      End
      Begin VB.Menu mnufileprint 
         Caption         =   "&Print Screen"
      End
      Begin VB.Menu mnufileprintsetup 
         Caption         =   "Printer &Setup"
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuforms 
      Caption         =   "&Forms"
   End
   Begin VB.Menu MNUReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnurepcalf 
         Caption         =   "C&alf Reports"
      End
      Begin VB.Menu mnuYReps 
         Caption         =   "Yearling Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu mnurepcows 
         Caption         =   "&Cow Reports"
      End
      Begin VB.Menu mnurepsires 
         Caption         =   "&Sire Reports"
      End
      Begin VB.Menu mnurepherd 
         Caption         =   "&Herd Reports"
      End
      Begin VB.Menu mnuusda 
         Caption         =   "Administrative Reports"
      End
      Begin VB.Menu MNUReportsMake 
         Caption         =   "Make Your Own"
      End
   End
   Begin VB.Menu mnuimport 
      Caption         =   "&Import"
      Begin VB.Menu MNUDos2Win 
         Caption         =   "&Dos To Windows Conversion"
      End
      Begin VB.Menu mnuimportspread 
         Caption         =   "Import Spreadsheet"
      End
   End
   Begin VB.Menu mnuexport 
      Caption         =   "&Export "
   End
   Begin VB.Menu MNUSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnusetupherd 
         Caption         =   "&Herd"
      End
      Begin VB.Menu MNUTurnOutDates 
         Caption         =   "Bull Turn Out Dates"
      End
      Begin VB.Menu mnuptod 
         Caption         =   "Prior Turn Out Date"
      End
      Begin VB.Menu mnusetupprefs 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuspaprfe 
         Caption         =   "&SPA Preferences"
      End
   End
   Begin VB.Menu mnuutils 
      Caption         =   "&Utilities"
      Begin VB.Menu MNUNotes 
         Caption         =   "&Update Notes"
         Begin VB.Menu MNUNotesCalves 
            Caption         =   "Calves"
         End
         Begin VB.Menu MNUNotesCows 
            Caption         =   "Cows"
         End
         Begin VB.Menu MNUNotesSires 
            Caption         =   "Sires"
         End
      End
      Begin VB.Menu MNUDatabase 
         Caption         =   "Database"
         Begin VB.Menu mnucompact 
            Caption         =   "&Compact Database"
         End
         Begin VB.Menu mnurepair 
            Caption         =   "&Repair Database"
         End
         Begin VB.Menu mnuview 
            Caption         =   "&View Database"
         End
         Begin VB.Menu MNUtilitiesArchive 
            Caption         =   "&Archive Database"
         End
      End
      Begin VB.Menu MNUDirectory 
         Caption         =   "Directories"
         Begin VB.Menu MNUSetupDir 
            Caption         =   "&Setup Directories"
         End
         Begin VB.Menu MNUChangeDirectory 
            Caption         =   "Change &Directory"
         End
      End
      Begin VB.Menu mnucopyData 
         Caption         =   "Copy Data"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu MNUGettingStarted 
         Caption         =   "&Getting Started"
      End
      Begin VB.Menu mnuhelponhelp 
         Caption         =   "&How to Use Help"
      End
      Begin VB.Menu mnuhelpsearch 
         Caption         =   "&Search for Help On"
      End
      Begin VB.Menu mnuhelpindex 
         Caption         =   "&Index"
      End
      Begin VB.Menu MnuRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuhelpsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "&Chaps Web"
      End
      Begin VB.Menu mnuhelpabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdimain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub writeuser()
 Dim temp$, tmp%
 temp$ = current_USER$
 tmp% = WritePrivateProfileString("Chaps User Information", "User ID", temp$, "chaps.ini")
End Sub

Private Sub CMDforms_Click()
Call mnuforms_Click
End Sub

Private Sub cmdopen_Click()
  PopupMenu mnufileopen
End Sub


Private Sub cmdprint_Click()
   PopupMenu MNUReports
End Sub


Private Sub Cmdreports_Click()

End Sub

Private Sub CMDSetup_Click()
 PopupMenu MNUSetup
End Sub


Private Sub MDIForm_Activate()
Static bFirstTime As Boolean
If bFirstTime Then Exit Sub
bFirstTime = True
Call MNUChangeDirectory_Click
End Sub

Private Sub MDIForm_Load()
Dim x%, TheDate$, theyear$
5  On Error GoTo ehandle
  
TheDate = Date
theyear = Right(TheDate, 4)
If InStr(1, theyear, "/") > 0 Then
   MsgBox "Please Set the Date Format to MM/DD/YYYY In the Control Panel Under Regional Settings", vbOKOnly
   End
End If
  
10 If IsRegistered Then
 
 If gIsDemo = True Then GoTo Show_Register
  
 DoEvents
 AbSplash.SplashOn Me, 3000
 'AbSplash.About Me, "", "Developed by:" & vbCrLf & "North Dakota Beef Cattle Improvement Association" & vbCrLf & "North Dakota State University Extension Service" & vbCrLf & vbCrLf & "Copyright 1985, 1987, 1988, 1990, 1993, 1999"
20 Call read_ini
Call UpdateDataBase2
21 AbSplash.SplashOff

Else

22 Show_Register:

23 FrmRegister.Show vbModal


24  DoEvents
25 AbSplash.SplashOn Me, 3000
 'AbSplash.About Me, "", "Developed by:" & vbCrLf & "North Dakota Beef Cattle Improvement Association" & vbCrLf & "North Dakota State University Extension Service" & vbCrLf & vbCrLf & "Copyright 1985, 1987, 1988, 1990, 1993, 1999"
30 Call read_ini
 AbSplash.SplashOff

End If

40 MnuRegister.Visible = gIsDemo
Me.pnlhint.Caption = dbfile$
Exit Sub
 
ehandle:
 If Err.Number = 3428 Or Err.Number = 3049 Or Err.Number = 3343 Then
   Load frmrepairdatabase
   frmrepairdatabase.Tag = "Repair"
   frmrepairdatabase.DB = dbfile$
   frmrepairdatabase.Show vbModal
   If frmrepairdatabase.Tag = "CANCEL" Then End
   Unload frmrepairdatabase
   Resume
 End If
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = Me.Name & " MDIForm_Load"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnucalves_Click()
  Load frmcalf_list
End Sub


Private Sub MNUChangeDirectory_Click()
Dim sPath$
Call SetupDirectoryForm(True)
Show_Modal:
frmchangedir.Show vbModal
If UCase(frmchangedir.Tag) <> "CANCEL" Then
   sPath$ = frmchangedir.Tag
   If Right(sPath$, 1) <> "\" Then sPath$ = sPath$ & "\"
   dbfile$ = sPath$ & "chaps.mdb"
   If Dir(dbfile$) = "" Then
      MsgBox "No chaps.mdb found in '" & frmchangedir.Tag & "' "
      GoTo Show_Modal
   End If
   Me.pnlhint.Caption = dbfile$
End If
Unload frmchangedir
End Sub

Private Sub mnucompact_Click()
 Load frmpassword_entry
 frmpassword_entry.Tag = "CNIISS"
 frmpassword_entry.Show vbModal
 If frmpassword_entry.Tag = "CANCEL" Then
   Unload frmpassword_entry
   Exit Sub
 End If
 Unload frmpassword_entry
 Load frmcompactdatabase
 frmcompactdatabase.DB = dbfile$ ' set the database to compact
 frmcompactdatabase.Show vbModal
 Unload frmcompactdatabase
End Sub

Private Sub mnucopyData_Click()
  FrmCopyData.Show vbModal
End Sub

Private Sub MNUDos2Win_Click()
Dim acctprogram
' Call writeuser
' AcctProgram = LoadProgram("Agvance Accounting", App.path & "\Agvacct.exe")
 Dim sProgram As String
 Call writeuser
' Select Case gPersonality
'  Case "HIS"
'   sProgram = "H.I.S. Accounting"
'  Case Else
'   sProgram = "Agvance Accounting"
' End Select
 sProgram = "DOS to Windows Conversion"
 acctprogram = LoadProgram(sProgram, App.Path & "\chapconv.exe")
End Sub

Private Sub mnuexport_Click()
 MsgBox "Please go to Reports menu and run the report that prints the information you would like to export. After running the chosen report in Preview mode, click on the icon that looks like an open envelope. Choose the Format and Destination. Click OK. Choose the Separator and Delimiter. Click OK. Choose the Number and Date formats you want. Click OK. Choose from the listbox the name for the file and where you want to store it. Click Save. The file will begin building. When it is done, the status box will disappear. You can close the report and find the file in the folder you saved it to.", vbInformation, Me.Caption
End Sub

Private Sub mnufileexit_Click()
Close
End
End Sub

Private Sub mnufileprint_Click()
On Error GoTo ehandle
 mdimain.ActiveForm.PrintForm
 Exit Sub
ehandle:
  Err.Clear: Exit Sub

End Sub

Private Sub mnufileprintsetup_Click()
 mdimain!CDG.Flags = cdlPDPrintSetup
 mdimain!CDG.Action = 5
 'MdiMain!cdiprintset.ShowPrinter
 cmdPrint.ToolTipText = Printer.DeviceName

End Sub


Private Sub mnuforms_Click()
  Input_forms.Show
End Sub

Private Sub MNUGettingStarted_Click()
 On Error GoTo ehandle
 
 GMODNAME$ = Me.Name & " mnuhelpcontents_click"
 HelpFile$ = App.Path & "\chaps.hlp"
 CDG.HelpFile = HelpFile$
 CDG.HelpCommand = &H3&   'cdlHelpContents
 CDG.Action = 6
 Exit Sub
 
ehandle:
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())


End Sub

Private Sub mnuhelpabout_Click()
   AbSplash.About Me
End Sub

Private Sub mnuimportspread_Click()
  frmImportData.Show vbModal
End Sub

Private Sub MNUNotesCalves_Click()
FrmUpdate_Calf.Show vbModal
End Sub

Private Sub MNUNotesCows_Click()
FrmUpdate_Cow.Show vbModal
End Sub

Private Sub MNUNotesSires_Click()
FrmUpdate_Sire.Show vbModal
End Sub

Private Sub mnuopencows_Click()
Load frmcow_list
End Sub


Private Sub mnuopensire_Click()
Load frmsire_list
End Sub




Private Sub mnuptod_Click()
  FrmoldDate.Show
End Sub

Private Sub MnuRegister_Click()
 Call FrmRegister.Show
End Sub

Private Sub mnurepair_Click()
Load frmpassword_entry
 frmpassword_entry.Tag = "CNIISS"
 frmpassword_entry.Show 1
 If frmpassword_entry.Tag = "CANCEL" Then
   Unload frmpassword_entry
   Exit Sub
 End If
 Unload frmpassword_entry
 Load frmrepairdatabase
 frmrepairdatabase.lblWarning.Caption = "!!!!Warning!!!! " & vbCrLf & "This Process Will Rewrite Your Database!" & vbCrLf & "Please Make A Backup Of Data Directory Before Continuing" & vbCrLf & " "
 frmrepairdatabase.Tag = "Repair"
 frmrepairdatabase.DB = dbfile$
 frmrepairdatabase.Show vbModal
 Unload frmrepairdatabase
End Sub

Private Sub mnurepcalf_Click()
   calfreps.Show
End Sub

Private Sub mnurepcows_Click()
  cowreps.Show
End Sub

Private Sub mnurepherd_Click()
  herdreps.Show
End Sub

Private Sub MNUReportsMake_Click()
  Call UpdateDataBase
  Load FrmHubQuery
End Sub

Private Sub mnurepsires_Click()
  sirereps.Show
End Sub

Private Sub MNUSetupDir_Click()
FRMSetupDirectories.Show vbModal
End Sub

Private Sub mnusetupherd_Click()
Load frmherd_List
End Sub


Private Sub mnusetupprefs_Click()
   Load frmpref
End Sub


Private Sub mnuspaprfe_Click()
    frmprefspa.Show
End Sub



Private Sub MNUtilitiesArchive_Click()
FrmArchive.Show
End Sub


Private Sub MNUTurnOutDates_Click()
FrmTurnDate.Show
End Sub

Private Sub mnuusda_Click()
 Load frmpassword_entry
 
 frmpassword_entry.Tag = BuildChapsPassword
 frmpassword_entry.Show vbModal
 If frmpassword_entry.Tag = "CANCEL" Then
   Unload frmpassword_entry
   Exit Sub
 End If
 Unload frmpassword_entry
 Load usdareps
End Sub

Private Sub mnuview_Click()
 Load frmpassword_entry
 frmpassword_entry.Tag = "CNIISS"
 frmpassword_entry.Show 1
 If frmpassword_entry.Tag = "CANCEL" Then
   Unload frmpassword_entry
   Exit Sub
 End If
  
 Unload frmpassword_entry
  
 frmdisplay_database.Show
End Sub


Private Sub mnuweb_Click()
Dim AgvanceProgram
Dim RESPONSE%
'AgvanceProgram = LoadProgram(sProgram, App.Path & "\Agvance.exe")
'AgvanceProgram = LoadProgram(sProgram, "C:\bbeef\welcome.htm")
RESPONSE% = MsgBox("If you have Internet access, you can view the Web page online. Would you like to view Web pages online?", vbYesNo)
If RESPONSE% = vbYes Then
  'Shell "Start " & "http://www.chaps2000.org/", vbHide
  Call HyperJump("www.chaps2000.org")
 Else
  'Shell "Start " & dbdir$ & "ChapsWeb\default.htm", vbHide
  'MsgBox App.Path & "\ChapsWeb\default.htm"
  'Shell "Start " & App.Path & "\ChapsWeb\default.htm", vbHide
  Call HyperJump(App.Path & "\ChapsWeb\default.htm")
End If
End Sub

Private Sub mnuYReps_Click()
YearlingsReps.Show
End Sub
