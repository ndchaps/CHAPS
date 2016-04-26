VERSION 5.00
Begin VB.Form frmrepairdatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repair Database"
   ClientHeight    =   2295
   ClientLeft      =   1140
   ClientTop       =   1515
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
      Left            =   5880
      TabIndex        =   1
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   385
      Left            =   4785
      TabIndex        =   0
      Top             =   1800
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
Attribute VB_Name = "frmrepairdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DbName As String
Property Let DB(database As String)
 DbName = database
End Property


Private Sub CMDCancel_Click()
 Me.Tag = "CANCEL"
 Me.Hide
End Sub

Private Sub CMDOk_Click()
 Dim megfree&, FileSize&, drive$, FILECMP&, FILECMP2&, Path$, i%
 On Local Error GoTo LeHandle
 'Mdimain!StatusBar1.Panels(6).TEXT = "Repairing Database...."
 DoEvents
 Screen.MousePointer = vbHourglass
 Path = ""
 For i% = Len(DbName) To 1 Step -1
   If Mid$(DbName, i%, 1) = "\" Then
     Path = Left$(DbName, i%)
     Exit For
   End If
 Next i%
 'check to see if enough space to do this
 If Mid$(Path$, 2, 1) = ":" Then
   drive$ = Left$(Path$, 1)
  Else
   drive$ = Left$(CurDir$, 1)
 End If
 megfree& = get_free_space(drive$)
 Open DbName For Input As #1: FileSize& = LOF(1): Close 1
 Open Left$(DbName, Len(DbName) - 4) & ".RPR" For Output As #1: FILECMP& = LOF(1): Close 1
 Open Left$(DbName, Len(DbName) - 4) & ".RP2" For Output As #1: FILECMP2& = LOF(1): Close 1
 If (FileSize& * 2) - FILECMP& - FILECMP2& > megfree& Then
   MsgBox "You Do Not Have Enough Room To Do This Operation" & vbCrLf & "Clear Off Some Drive Space And Try Again.", vbOKOnly + vbCritical, Me.Caption
   'Mdimain!StatusBar1.Panels(6).TEXT = ""
   Screen.MousePointer = vbDefault
   Me.Tag = "CANCEL"
   Me.Hide
   Exit Sub
 End If
 Open Left$(DbName, Len(DbName) - 4) & ".RPR" For Output As #1: Close 1: Kill Path$ & "*.RPR"
 FileCopy DbName, Left$(DbName, Len(DbName) - 4) & ".RPR"
 DBEngine.RepairDatabase DbName
 Open Left$(DbName, Len(DbName) - 4) & ".rp2" For Output As #1: Close 1: Kill Path$ & "*.rp2"
 'Mdimain!StatusBar1.Panels(6).TEXT = "Compacting Database...."
 DoEvents
 FileCopy DbName, Left$(DbName, Len(DbName) - 4) & ".rp2"
 Kill DbName
 DBEngine.CompactDatabase Left$(DbName, Len(DbName) - 4) & ".rp2", DbName
 Screen.MousePointer = vbDefault
 'Mdimain!StatusBar1.Panels(6).TEXT = ""
 Me.Hide
Exit Sub
 
LeHandle:
 If Err = 70 Then
   MsgBox "Database is open, check that all user are out of " & App.ProductName & "." & vbCrLf & " You must have everyone else close then database and click OK", vbOKOnly, Me.Caption
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
 lblWarning.Caption = "!!!!Warning!!!! " & vbCrLf & "This Process Will Rewrite Your DataBase!" & vbCrLf & "Please Make A Backup Of Data Directory Before Continuing"

End Sub
