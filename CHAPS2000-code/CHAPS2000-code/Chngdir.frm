VERSION 5.00
Begin VB.Form frmchangedir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Directory"
   ClientHeight    =   2925
   ClientLeft      =   3150
   ClientTop       =   2790
   ClientWidth     =   3015
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDirName 
      Height          =   285
      Left            =   660
      TabIndex        =   4
      Top             =   1800
      Width           =   2250
   End
   Begin VB.CommandButton cmdchangedir 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Index           =   1
      Left            =   1635
      TabIndex        =   3
      Top             =   2445
      Width           =   1000
   End
   Begin VB.DriveListBox drvchangedrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdchangedir 
      Caption         =   "&OK"
      Height          =   385
      Index           =   0
      Left            =   315
      TabIndex        =   2
      Top             =   2445
      Width           =   1000
   End
   Begin VB.DirListBox dirchangedir 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   2775
   End
   Begin VB.Label LBLDir 
      Alignment       =   1  'Right Justify
      Caption         =   "Dir"
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   1815
      Width           =   495
   End
End
Attribute VB_Name = "frmchangedir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveDrive As String

Private Sub cmdchangedir_Click(Index As Integer)
 Select Case Index
  Case 0
   Me.Tag = dirchangedir.Path
  Case 1
   Me.Tag = "CANCEL"
 End Select
 Me.Hide
End Sub

Private Sub drvchangedrive_Change()
 On Local Error GoTo LeHandle
 dirchangedir.Path = drvchangedrive.drive
 SaveDrive = drvchangedrive.drive
Exit Sub
 
LeHandle:
 If Err.Number = 68 Then
   MsgBox "Drive not avaliable please choose another drive", vbOKOnly + vbCritical, Me.Caption
   drvchangedrive.drive = SaveDrive
   Resume
 End If
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = Me.Name & " DrvChangeDrive"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub


Private Sub drvchangedrive_GotFocus()
 SaveDrive = drvchangedrive
End Sub


Private Sub Form_Activate()
If Me.Tag = "" Then
   drvchangedrive.drive = Left$(CurDir$, 1)
   dirchangedir.Path = CurDir$
   Exit Sub
 End If
'lynn Me.Tag = ""
' frmchangedir.Show vbModal
 'lynn drvchangedrive.drive = Left$(CurDir$, 1)
 'lynn dirchangedir.Path = CurDir$
 drvchangedrive.drive = Left$(Me.Tag, 1)
 If Me.Tag <> "CANCEL" Then dirchangedir.Path = Me.Tag
End Sub

Private Sub Form_Load()
 'If Me.Tag = "" Then
 '  drvchangedrive.drive = Left$(CurDir$, 1)
 '  dirchangedir.Path = CurDir$
 '  Exit Sub
 'End If
'lynn Me.Tag = ""
' frmchangedir.Show vbModal
 'lynn drvchangedrive.drive = Left$(CurDir$, 1)
 'lynn dirchangedir.Path = CurDir$
 'drvchangedrive.drive = Left$(Me.Tag, 1)
 'dirchangedir.Path = Me.Tag
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then
   Me.Tag = "CANCEL"
   Cancel = True
   Me.Hide
 End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Set frmchangedir = Nothing
End Sub


