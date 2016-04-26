VERSION 5.00
Begin VB.Form frmgetdisk 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1140
   ClientLeft      =   2685
   ClientTop       =   1965
   ClientWidth     =   3000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1140
   ScaleWidth      =   3000
   Begin VB.CommandButton cmdcancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   1890
      TabIndex        =   3
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton cmdcontinue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   135
      TabIndex        =   2
      Top             =   600
      Width           =   1000
   End
   Begin VB.Label LBLDRIVE2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "345345"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -15
      TabIndex        =   1
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please put diskette in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -15
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmgetdisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
 DISKCANCEL% = True
 Unload Me
End Sub

Private Sub cmdcontinue_Click()
 Dim RESPONSE%
 On Local Error GoTo lehandle

checkdisk:
 If Right(DISKDRIVE$, 1) = ":" Then
    Open DISKDRIVE$ + "\TEMPfile" For Output As #1
 Else
    Open DISKDRIVE$ + ":\TEMPfile" For Output As #1
 End If
 Close 1
 Unload Me
 Exit Sub

lehandle:
 If Err = 52 Or Err = 68 Or Err = 57 Or Err = 71 Or Err = 72 Then
   RESPONSE% = MsgBox("Disk Error Please Try Again", vbOKCancel, diskcaption$)
   If RESPONSE% = vbCancel Then
     DISKCANCEL% = True
     Call cmdcancel_Click
     Exit Sub
    Else
     Resume checkdisk
   End If
 End If
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = Me.Name & " cmdcontinue_click"
 GERRNUM$ = Str$(Err.Number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())

End Sub

Private Sub form_Load()
 Call centerform(Me, 0, 0)
 Me.caption = diskcaption$
 LBLDRIVE2.caption = "Drive " + DISKDRIVE$
End Sub

