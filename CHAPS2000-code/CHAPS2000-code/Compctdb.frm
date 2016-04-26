VERSION 5.00
Begin VB.Form frmcompactdatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compact Database"
   ClientHeight    =   2295
   ClientLeft      =   1320
   ClientTop       =   2610
   ClientWidth     =   7005
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   5895
      TabIndex        =   1
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   4800
      TabIndex        =   0
      Top             =   1815
      Width           =   1000
   End
   Begin VB.Label lblwarning 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmcompactdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Warningstring As String
Dim DbName As String
Property Let db(db As String)
 DbName = db
End Property

Private Sub cmdCancel_Click()
 Me.tag = "CANCEL"
 Me.Hide
End Sub

Private Sub cmdOK_Click()
 Dim megfree&, FileSize&, drive$, FILECMP&
 On Local Error GoTo LeHandle
 'Mdimain!StatusBar1.Panels(6).TEXT = "Compacting Database...."
 DoEvents
 Screen.MousePointer = vbHourglass
 'check to see if enough space to do this
 If Mid$(dbdir$, 2, 1) = ":" Then
   drive$ = left$(dbdir$, 1)
  Else
   drive$ = left$(CurDir$, 1)
 End If
 megfree& = get_free_space(drive$)
10 Open DbName For Input As #1: FileSize& = LOF(1): Close 1
11 Open left$(DbName, Len(DbName) - 4) & ".cmp" For Output As #1: FILECMP& = LOF(1): Close 1: Kill left$(DbName, Len(DbName) - 4) & ".cmp"
12 If (FileSize& * 2) - FILECMP& > megfree& Then
     MsgBox "You Do Not Have Enough Room To Do This Operation" & vbCrLf & "Clear Off Some Drive Space And Try Again.", vbOKOnly + vbCritical, Me.caption
     'Mdimain!StatusBar1.Panels(6).TEXT = ""
     Screen.MousePointer = vbDefault
     Me.tag = "CANCEL"
     Me.Hide
     Exit Sub
   End If
  'Call ClearTableAttachemnts(dbname, "")
13 Open left$(DbName, Len(DbName) - 4) & ".cmp" For Output As #1: Close 1: Kill left$(DbName, Len(DbName) - 4) & ".cmp"
14 FileCopy DbName, left$(DbName, Len(DbName) - 4) & ".cmp"
15 Kill DbName
16 DBEngine.CompactDatabase left$(DbName, Len(DbName) - 4) & ".cmp", DbName
17 Screen.MousePointer = vbDefault
18 'Mdimain!StatusBar1.Panels(6).TEXT = ""
19 Me.Hide
20 Exit Sub
 
LeHandle:
 If Err = 70 Then
   MsgBox "Database is open, check that all user are out of " & App.productname & "." & vbCrLf & " You must have everyone else close then database and click OK", vbOKOnly, Me.caption
   Resume
 End If
 'Text$(1) = ""
 'Text$(2) = ""
 'Text$(3) = ""
 'Text$(4) = ""
 'Text$(5) = ""
 'GERRNUM$ = Str$(Err.number)
 'GERRSOURCE$ = Err.Source
 'Call POP_ERROR(Text$())
End Sub

Private Sub Form_Load()
' Call centermdiform(Me, Mdimain, 0, 0)
 lblwarning.caption = "!!!!Warning!!!! " & vbCrLf & "This process will rewrite your dataBase!" & vbCrLf & "Please make a backup of data directory before continuing." & vbCrLf & " Please make sure all other network stations are out of " & App.productname & "."
End Sub

Public Function get_free_space(drive$) As Long
 Dim SaveDrive$
 SaveDrive$ = left$(CurDir$, 1)
 ChDrive left$(drive$, 1)
 get_free_space = DiskSpaceFree()
 ChDrive SaveDrive$
End Function
