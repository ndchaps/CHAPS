VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSACAL70.OCX"
Begin VB.Form frmcalendar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2025
   ClientLeft      =   4035
   ClientTop       =   2925
   ClientWidth     =   1860
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand Command2 
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Top             =   45
      Width           =   300
      _Version        =   65536
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "calendar.frx":0000
   End
   Begin Threed.SSCommand Command1 
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   300
      _Version        =   65536
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "calendar.frx":04BE
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2055
      Left            =   -15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1890
      _Version        =   458752
      _ExtentX        =   3334
      _ExtentY        =   3625
      _StockProps     =   1
      BackColor       =   12632256
      Year            =   1997
      Month           =   2
      Day             =   5
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DayLength       =   0
      GridCellEffect  =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthLength     =   0
      ShowDateSelectors=   0   'False
      ShowHorizontalGrid=   0   'False
      ShowVerticalGrid=   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmcalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calendar1_Click()
gcaldate = SSIDate(Calendar1)
Unload Me

End Sub

Private Sub Command1_Click()
If Calendar1.Month > 1 Then
 Calendar1.Month = Calendar1.Month - 1
Else
 Calendar1.Month = 12
 Calendar1.Year = Calendar1.Year - 1
End If
End Sub

Private Sub Command2_Click()
If Calendar1.Month < 12 Then
 Calendar1.Month = Calendar1.Month + 1
 Else
 Calendar1.Month = 1
 Calendar1.Year = Calendar1.Year + 1
End If
End Sub

Private Sub Form_Load()
Calendar1 = gcaldate

End Sub


