VERSION 5.00
Begin VB.Form FRMSetupDirectories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup CHAPS Directories"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDOk 
      Caption         =   "OK"
      Height          =   405
      Left            =   5745
      TabIndex        =   7
      Top             =   2145
      Width           =   990
   End
   Begin VB.CommandButton CMDNewDir 
      Caption         =   "New Dir"
      Height          =   405
      Left            =   4530
      TabIndex        =   6
      Top             =   2145
      Width           =   990
   End
   Begin VB.CommandButton CMDMove2 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2985
      TabIndex        =   5
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton CMDMove 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2985
      TabIndex        =   4
      Top             =   1530
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   4020
      TabIndex        =   3
      Top             =   675
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4020
      TabIndex        =   2
      Top             =   300
      Width           =   2775
   End
   Begin VB.DirListBox dirchangedir 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   2775
   End
   Begin VB.DriveListBox drvchangedrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   285
      Width           =   2775
   End
End
Attribute VB_Name = "FRMSetupDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function MDBExists(sPath$) As Boolean
If Right(sPath$, 1) <> "\" Then sPath$ = sPath$ & "\"
If Dir(sPath$ & "chaps.mdb") = "" Then MDBExists = False Else MDBExists = True
End Function

Private Sub CMDMove_Click()
Dim sToPath$, sFromPath$
sToPath$ = Dir1.Path
sFromPath$ = dirchangedir.Path
If Right(sToPath$, 1) <> "\" Then sToPath$ = sToPath$ & "\"
If Right(sFromPath$, 1) <> "\" Then sFromPath$ = sFromPath$ & "\"
If Not MDBExists(sFromPath$) Then MsgBox "No chaps.mdb found in '" & sFromPath$ & "' ": Exit Sub
If MDBExists(sToPath$) Then MsgBox "chaps.mdb found in '" & sToPath$ & "' ": Exit Sub
FileCopy sFromPath$ & "chaps.mdb", sToPath$ & "chaps.mdb"
End Sub

Private Sub CMDMove2_Click()
Dim sToPath$, sFromPath$
If Not MDBExists(Dir1.Path) Then MsgBox "No chaps.mdb found in '" & Dir1.Path & "' ": Exit Sub
If MDBExists(dirchangedir.Path) Then MsgBox "chaps.mdb found in '" & dirchangedir.Path & "' ": Exit Sub
sToPath$ = dirchangedir.Path
sFromPath$ = Dir1.Path
If Right(sToPath$, 1) <> "\" Then sToPath$ = sToPath$ & "\"
If Right(sFromPath$, 1) <> "\" Then sFromPath$ = sFromPath$ & "\"
FileCopy sFromPath$ & "chaps.mdb", sToPath$ & "chaps.mdb"
End Sub

Private Sub CMDNewDir_Click()
Dim s$
Call SetupDirectoryForm(False)
frmchangedir.Show vbModal, Me
If UCase(frmchangedir.Tag) <> "CANCEL" Then
   s$ = frmchangedir.Tag
   If Right$(s$, 1) <> "\" Then s$ = s$ & "\"
   If Dir(s$ & frmchangedir.txtDirName.TEXT, vbDirectory) = "" Then
      Call MkDir(s$ & frmchangedir.txtDirName.TEXT)
      MsgBox "Directory '" & s$ & frmchangedir.txtDirName.TEXT & "' created successfully"
   Else
      MsgBox "Directory '" & s$ & frmchangedir.txtDirName.TEXT & "' already exists"
   End If
End If
Unload frmchangedir
End Sub

Private Sub CMDOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call centermdiform(Me, mdimain, 0, 0)
End Sub
