VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmCopyData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Database"
   ClientHeight    =   2445
   ClientLeft      =   2160
   ClientTop       =   3525
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&Browse"
      Height          =   300
      Left            =   4860
      TabIndex        =   4
      Top             =   495
      Width           =   930
   End
   Begin VB.TextBox Txtdb 
      Height          =   300
      Left            =   2100
      TabIndex        =   2
      Top             =   450
      Width           =   2460
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "C&opy"
      Height          =   385
      Left            =   4080
      TabIndex        =   1
      Top             =   1980
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   5400
      TabIndex        =   0
      Top             =   1980
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog cdlfilename 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Database To"
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1770
   End
End
Attribute VB_Name = "FrmCopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbrowse_Click()
frmchangedir.Show vbModal
If UCase(frmchangedir.Tag) <> "CANCEL" Then
   Txtdb.TEXT = frmchangedir.Tag
   If Right(Txtdb.TEXT, 1) <> "\" Then Txtdb.TEXT = Txtdb.TEXT & "\"
End If
Unload frmchangedir

End Sub

Private Sub cmdcancel_Click()
  Unload Me
End Sub

Private Sub cmdchangedir_Click(Index As Integer)

End Sub


Private Sub cmdcopy_Click()
 Dim OLdDb$
 If Txtdb.TEXT = "" Then Exit Sub
 OLdDb$ = Txtdb.TEXT & "cHAPS.MDB"
 If FileExist(OLdDb$) Then Kill OLdDb$
 FileCopy dbfile$, OLdDb$
 Unload Me
 
End Sub


