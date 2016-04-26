VERSION 5.00
Begin VB.Form frmpassword_entry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Entry"
   ClientHeight    =   1680
   ClientLeft      =   2760
   ClientTop       =   2640
   ClientWidth     =   3060
   ClipControls    =   0   'False
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
   ScaleHeight     =   1680
   ScaleWidth      =   3060
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdpassword 
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
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton cmdpassword 
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
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label lblinstructions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the password below"
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
      Height          =   195
      Index           =   2
      Left            =   15
      TabIndex        =   2
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label lblinstructions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "to gain access"
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
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label lblinstructions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "This area needs a password"
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
      Height          =   195
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmpassword_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ATTEMPT%

Private Sub cmdpassword_Click(Index As Integer)
 Dim response%
 On Error GoTo ehandle
 Select Case Index
  Case 0
   If UCase$(txtpassword.TEXT) <> UCase$(Me.Tag) Then
     ATTEMPT% = ATTEMPT% + 1
     If ATTEMPT% > 3 Then
       MsgBox "You Have Exceeded The Number Of Password Attempts", vbOKOnly + vbExclamation, Me.caption
       Me.Tag = "CANCEL"
       Me.Hide
       Exit Sub
     End If
     response% = MsgBox("Incorrect Password", vbOKCancel + vbExclamation, Me.caption)
     If response% = vbCancel Then
       Me.Tag = "CANCEL"
       Me.Hide
       Exit Sub
     End If
     txtpassword.SetFocus
     Call txtpassword_GotFocus
     Exit Sub
   End If
   Me.Tag = "OK"
   Me.Hide
  Case 1
  Me.Tag = "CANCEL"
  Me.Hide
 End Select
 Exit Sub
 
ehandle:
 TEXT$(1) = ""
 TEXT$(2) = ""
 TEXT$(3) = ""
 TEXT$(4) = ""
 TEXT$(5) = ""
 GMODNAME$ = Me.name & " cmdpassword_click"
 GERRNUM$ = Str$(Err.number)
 GERRSOURCE$ = Err.Source
 Call POP_ERROR(TEXT$())
End Sub

Private Sub form_Load()
 Screen.MousePointer = vbHourglass
  Call centerform(Me, 0, 0)
 Screen.MousePointer = vbDefault
 ATTEMPT% = 0
End Sub

Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then
   Me.Tag = "CANCEL"
   Me.Hide
   Cancel = True
 End If
End Sub


Private Sub txtpassword_GotFocus()
 txtpassword.SelStart = 0
 txtpassword.SelLength = Len(txtpassword)
End Sub


