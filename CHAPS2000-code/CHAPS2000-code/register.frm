VERSION 5.00
Begin VB.Form FrmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register C.H.A.P.S."
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4245
      TabIndex        =   4
      Top             =   2175
      Width           =   840
   End
   Begin VB.CommandButton CMDContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   5130
      TabIndex        =   3
      Top             =   2175
      Width           =   840
   End
   Begin VB.TextBox txtActivationKey 
      Height          =   285
      Left            =   2190
      TabIndex        =   2
      Top             =   1395
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Activation Key"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label LBLWelcome 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Insert welcome message here."
      Height          =   1230
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5880
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCancel_Click()
Unload Me
End
End Sub

Private Sub CMDContinue_Click()
If IsValidKey(txtActivationKey.TEXT) = False Then MsgBox "Please enter a valid activation key", vbOKOnly, Me.Caption: txtActivationKey.SetFocus: Exit Sub
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
txtActivationKey.TEXT = GetActivationKey
End Sub
