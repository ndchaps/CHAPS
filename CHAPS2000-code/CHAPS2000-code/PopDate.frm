VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPopUpDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Enter New Exposed Date"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3300
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2340
      TabIndex        =   0
      Top             =   1200
      Width           =   915
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   270
      Left            =   4185
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   690
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   476
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "PopDate.frx":0000
   End
   Begin MSMask.MaskEdBox txtExpDate 
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   660
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   327681
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "-"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   435
      Left            =   1860
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Exposed Date"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "If the New Exposed Date Is Within the 285 Day Period the Existing Exposed Date Will Be Overwritten"
      Height          =   1155
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmPopUpDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cow$

Private Sub cmdcancel_Click()
Unload Me
End Sub

Property Let CowID(CowID$)
Cow$ = CowID$
End Property

Private Sub cmdok_Click()
Dim tbData As Recordset, iResponse
If txtExpDate.TEXT <> "--/--/----" Then
   Set db = DBEngine(0).OpenDatabase(dbfile$, False, False)
      Set tbData = db.OpenRecordset("misc", dbOpenTable)
         tbData.Index = "primarykey"
         tbData.Seek "=", "TurnDate" & herdid$
      If Not tbData.NoMatch Then
        If DateDiff("d", tbData!thetext, txtExpDate.TEXT) < 285 Then
            iResponse = MsgBox("Do You Wish To Overwrite The Current Exposed Date For This Herd?", vbYesNo + vbQuestion, Me.Caption)
            If iResponse = vbYes Then
               db.Execute ("UPDATE DISTINCTROW cowprof INNER JOIN cowbrd ON (cowprof.cowID = cowbrd.CowID) AND (cowprof.HerdID = cowbrd.HerdID) SET cowbrd.calfdate = #" & txtExpDate.TEXT & "# WHERE (((cowprof.active)='A')) and cowbrd.calfdate = #" & tbData!thetext & "# and cowprof.herdid = '" & herdid$ & "'")
               tbData.Edit: tbData!thetext = txtExpDate.TEXT: tbData.Update
               'Update all active (cowbrd!active = 'A') cows' breed date (cowbrd!calfdate) to new bull turn out date
               GoTo CloseDB
            Else
               GoTo CloseDB
            End If
         End If
         tbData.Edit: tbData!thetext = txtExpDate.TEXT: tbData.Update
      Else
         tbData.AddNew: tbData!thekey = "TurnDate" & herdid$: tbData!thetext = txtExpDate.TEXT: tbData.Update
      End If
Else
   MsgBox "Please Enter A Valid Date", vbOKOnly + vbCritical
   txtExpDate.SetFocus
   Exit Sub
End If
Exit Sub
CloseDB:
tbData.Close: Set tbData = Nothing
db.Close: Set db = Nothing
Me.Hide
frmcow_data.cboyear.RemoveItem (0)
frmcow_data.cboyear.AddItem "*" & txtExpDate.TEXT, 0
frmcow_data.cboyear.ListIndex = 0
End Sub

Private Sub Form_Load()
If TurnDate = "12:00:00 AM" Then
   Label2.Caption = "The Default Exposed Date Isn't Set"
Else
   Label2.Caption = "The Current Exposed Date Is " & CStr(TurnDate)
End If
End Sub

Private Sub SSCommand4_Click()
gcaldate = txtExpDate.TEXT
 Call GetDate(gcaldate)
 txtExpDate.TEXT = gcaldate
End Sub
