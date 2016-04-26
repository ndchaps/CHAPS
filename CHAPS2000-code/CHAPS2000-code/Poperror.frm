VERSION 5.00
Begin VB.Form frmpop_error 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   435
   ClientTop       =   1155
   ClientWidth     =   8220
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
   ScaleHeight     =   4635
   ScaleWidth      =   8220
   Begin VB.TextBox txtcomments 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4905
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5745
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdcontinue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   21
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "PH#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lbluser_message 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4905
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblcompany_phone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company Phone"
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
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label lblcompany_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "company name"
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
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PLEASE FAX TO NDBCIA  (701) 483-2005"
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
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label lblerror_line 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "label19"
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
      Left            =   1080
      TabIndex        =   17
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label lblerror_number 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label18"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label lblerror_time 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label17"
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
      Height          =   200
      Left            =   5730
      TabIndex        =   2
      Top             =   545
      Width           =   1400
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Line:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Error #:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time:"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   540
      Width           =   705
   End
   Begin VB.Label lblevent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label12"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Label lblmodule 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   3000
      Width           =   3645
   End
   Begin VB.Label lblerror_date 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
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
      Height          =   200
      Left            =   5730
      TabIndex        =   3
      Top             =   325
      Width           =   1400
   End
   Begin VB.Label lblerror_message 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Label lblmessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
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
      Height          =   615
      Left            =   15
      TabIndex        =   7
      Top             =   960
      Width           =   6990
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Source:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Module:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date:"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   330
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   870
   End
End
Attribute VB_Name = "frmpop_error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDContinue_Click()
 Dim dbpm As database
 Dim TBUSERS As Recordset
 On Local Error GoTo LeHandle
 Set dbpm = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
 Set TBUSERS = dbpm.OpenRecordset("USERS", dbOpenTable)
 If Not TBUSERS.BOF And Not TBUSERS.EOF Then
   TBUSERS.Index = "PRIMARYKEY"
   TBUSERS.Seek "=", current_USER$
   If Not TBUSERS.NoMatch Then
     TBUSERS.Edit
     TBUSERS!USERLOGGEDON = False
     TBUSERS!logondate = Null
     TBUSERS!logontime = Null
     TBUSERS.Update
   End If
 End If
 TBUSERS.Close: Set TBUSERS = Nothing
 dbpm.Close: Set dbpm = Nothing
LeHandle:
 Close
 End
End Sub

Private Sub cmdprint_Click()
frmpop_error.BackColor = QBColor(7)
'mdimain.ActiveForm.PrintForm
frmpop_error.PrintForm
End Sub

Private Sub Form_Load()
 Dim tme$
 Dim hr
 Dim AMPM$
' Call centermdiform(Me, mdimain, 0, 0)
 Me.Caption = "Error Encountered For User Number " & current_USER$
 If CURRENT_USER_NAME$ <> "" Then Me.Caption = Me.Caption & " " & CURRENT_USER_NAME$
 lblerror_message.Caption = Error$
 lblerror_date = Left$(Date$, 6) & Right$(Date$, 2)
 tme$ = Time$
 hr = Val(Left$(tme$, 2)): AMPM$ = " AM"
 If hr >= 12 Then AMPM$ = " PM"
 If hr > 12 Then hr = hr - 12
 If hr = 0 Then hr = 12
 tme$ = LTrim$(Str$(hr)) + Mid$(tme$, 3, 3) + AMPM$

 lblerror_time = tme$
 lblerror_number = GERRNUM$
 lblerror_line = LTrim$(Str$(Erl))
 lblmessage.Caption = ""
' lblform.caption = mdimain.ActiveForm.caption
 lblevent.Caption = GERRSOURCE$
 lblmodule.Caption = GMODNAME$
 'lblreference_number.caption = ""

 lblcompany_name.Caption = APP_COMPANY_name
 lblcompany_phone.Caption = APP_COMPANY_phone
 Screen.MousePointer = vbDefault
End Sub

