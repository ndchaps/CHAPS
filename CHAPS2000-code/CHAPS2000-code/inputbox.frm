VERSION 5.00
Begin VB.Form frminput_box 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   3375
   ClientTop       =   2055
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1095
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtinput 
      Height          =   285
      Left            =   2850
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   1815
   End
   Begin VB.CommandButton cmdinput_box 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2475
      TabIndex        =   3
      Top             =   630
      Width           =   1000
   End
   Begin VB.CommandButton cmdinput_box 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1410
      TabIndex        =   2
      Top             =   630
      Width           =   1000
   End
   Begin VB.Label lblinput 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   225
      Width           =   2730
   End
End
Attribute VB_Name = "frminput_box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdinput_box_Click(Index As Integer)
 Select Case Index
  Case 0
   If Len(txtinput.text) < 1 Then
    MsgBox "You Must Enter In A " & lblinput.caption, vbOK, Me.caption
    Exit Sub
   End If
   Me.tag = txtinput.text
  Case 1
  Me.tag = "Cancel"
 End Select
 Me.Hide
End Sub


Private Sub Form_Activate()
 txtinput.SetFocus
 'txtinput.TEXT = ""
 Me.tag = ""
End Sub

Private Sub Form_Load()
 Call centerform(Me, 0, 0)
 txtinput.text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
' Me.Hide
End Sub


