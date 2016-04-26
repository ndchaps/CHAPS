VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmCrPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   30
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnufileprintersetup 
         Caption         =   "P&rinter Setup"
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnu1 
         Caption         =   "Page &Layout"
      End
      Begin VB.Menu Mnu2 
         Caption         =   "Fit &Screen"
      End
      Begin VB.Menu Mnu3 
         Caption         =   "Full &Page"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmCrPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height
Me.Width = Screen.Width
'Call centerform(Me, 0, 0)
Screen.MousePointer = vbHourglass
End Sub




Private Sub Form_Unload(Cancel As Integer)
 Set frmCrPrint = Nothing
 Screen.MousePointer = vbDefault
End Sub


Private Sub mnu1_Click()
  report.fullpage
End Sub

Private Sub Mnu2_Click()
   report.oneside
End Sub

Private Sub Mnu3_Click()
  report.bothside
End Sub


Private Sub mnuabout_Click()
   report.rptinfo
End Sub




Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnufileprint_Click()
report.goprint
End Sub


Private Sub mnufileprintersetup_Click()
 Dim Result%
 CommonDialog1.CancelError = True
 CommonDialog1.Flags = cdlPDPrintSetup
' CommonDialog1.action = 5
  On Error GoTo done
  CommonDialog1.ShowPrinter
' If CommonDialog1.CancelError = True Then Exit Sub
 'If cdlCancel Then Exit Sub
 

 Unload Me
 
 report.PrintReport
 'report.goprint
done:
End Sub


