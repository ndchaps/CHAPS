VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmoldDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Prior Bull Turn Out Date"
   ClientHeight    =   4095
   ClientLeft      =   3300
   ClientTop       =   1815
   ClientWidth     =   3735
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      Caption         =   "Prior Bull Turn Out Date"
      Height          =   3120
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   3450
      Begin VB.ComboBox CBOCurrHerd 
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   345
         Width           =   1275
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   270
         Left            =   2790
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   765
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   476
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "olddate.frx":0000
      End
      Begin MSMask.MaskEdBox txtTurnDate 
         Height          =   315
         Left            =   1785
         TabIndex        =   4
         Top             =   735
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "-"
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Herd"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   1530
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Height          =   1350
         Left            =   90
         TabIndex        =   7
         Top             =   1740
         Width           =   3195
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "Only One Bull Turn Out Date Is Allowed In A 285 Day Period."
         Height          =   465
         Left            =   195
         TabIndex        =   6
         Top             =   1215
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Bull Turn Out Date"
         Height          =   315
         Left            =   225
         TabIndex        =   5
         Top             =   795
         Width           =   1455
      End
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   385
      Left            =   1365
      TabIndex        =   0
      Top             =   3660
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   2565
      TabIndex        =   1
      Top             =   3660
      Width           =   1000
   End
End
Attribute VB_Name = "FrmoldDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%
Dim oldid$
Dim tbData As Recordset
Dim OLDDATE$(), CurDate$
Dim mFormLoad As Boolean

Private Sub Init_Information()
Call init_form(Me) ' Clear Text Boxes

End Sub

Private Sub LoadCBOHERD(CBO As ComboBox)
Dim DB As DAO.database, RS As DAO.Recordset
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set RS = DB.OpenRecordset("select HerdID from herd order by herdid", dbOpenSnapshot)
CBO.Clear
Do Until RS.EOF
   CBO.AddItem Field2Str(RS!herdid)
   RS.MoveNext
Loop
If CBO.ListCount > 0 Then CBO.ListIndex = 0
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub Load_information()
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set tbData = DB.OpenRecordset("turnoutdate", dbOpenTable)
tbData.Index = "PrimaryKey"
tbData.Seek "=", CBOCurrHerd.TEXT
If Not tbData.NoMatch Then
   txtTurnDate.TEXT = Field2Date(tbData!currentdate)
Else
   txtTurnDate.TEXT = Format(Now, "mm/dd/yyyy")
   If mFormLoad = False Then MsgBox "This Herd Does Not Have A Default Bull Turn Out Date.", vbOKOnly, Me.Caption
End If
tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub save_information()
'Dim replace$, save As Boolean, iResponse As Integer, pOldDate As String
Screen.MousePointer = vbHourglass
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
If txtTurnDate.TEXT <> "--/--/----" Then
   DB.Execute "INSERT INTO cowbrd ( Stat, HerdID, CowID, calfdate) SELECT DISTINCTROW 'P' as Stat, cowprof.HerdID, cowprof.cowID, #" & txtTurnDate.TEXT & "# AS TurnDate From cowbrd, cowprof Where (((cowprof.active) = 'A')) and cowprof.herdid = '" & CBOCurrHerd.TEXT & "'"
End If
'CloseDB:
'tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
dirtyflag% = False
Screen.MousePointer = vbDefault
End Sub


Private Sub valid_form(exitcode%)
    Dim iResponse As Integer, pCurDate As String
    exitcode% = 0
    
Exit_Sub:
   '
End Sub

Private Sub CBOCurrHerd_Click()
Call Load_information
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim TableName$(100)
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
Call save_information
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
End Sub


Private Sub Form_Load()
Call centermdiform(Me, mdimain, 0, 0)
Label15.Caption = "Selecting A Bull Turn Out Date Creates An Exposed Date Record On The Breeding/Conception Tab For All Active Cows.  If Your Ranching Business Has More Than One Calving Season Per Year Create Only One Bull Turn Out Date Annually"
mFormLoad = True
Call LoadCBOHERD(CBOCurrHerd)
Call Load_information
mFormLoad = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 ' see if infomation needs to be saveed
 Dim RESPONSE%, exitcode%
 If dirtyflag% Then
   Beep
   RESPONSE% = MsgBox("Information Has Been Changed" & vbCrLf & " Do You Wish To Save?", vbYesNoCancel + vbQuestion, Me.Caption)
   Select Case RESPONSE%
    Case vbYes
     Call valid_form(exitcode%)
     If exitcode% <> 0 Then
       Cancel = True
     End If
     Call save_information
    Case vbCancel
     Cancel = True
   End Select
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Cancel Then Exit Sub
 Set frmherd_data = Nothing
End Sub

Private Sub load_year()
'
End Sub

Private Sub SSCommand3_Click()
gcaldate = txtTurnDate.TEXT
Call GetDate(gcaldate)
txtTurnDate.TEXT = gcaldate
End Sub
