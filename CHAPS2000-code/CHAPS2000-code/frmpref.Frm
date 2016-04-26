VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpref 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   4395
   ClientLeft      =   285
   ClientTop       =   1425
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   7755
   Begin VB.ComboBox CBOHerdID 
      Height          =   315
      Left            =   2865
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   3825
      Width           =   1665
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3465
      Left            =   135
      TabIndex        =   79
      Top             =   150
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6112
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Calf Misc"
      TabPicture(0)   =   "frmpref.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Epd Misc"
      TabPicture(1)   =   "frmpref.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Sort Orders"
      TabPicture(2)   =   "frmpref.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   3015
         Left            =   90
         TabIndex        =   64
         Top             =   330
         Width           =   7110
         Begin VB.Frame Frame5 
            Caption         =   "Sort Order of Sires"
            Height          =   1875
            Left            =   3705
            TabIndex        =   71
            Top             =   375
            Width           =   2445
            Begin VB.OptionButton Optsire 
               Caption         =   "Breed, Year, Sire ID"
               Height          =   195
               Index           =   4
               Left            =   225
               TabIndex        =   76
               Top             =   1500
               Width           =   1830
            End
            Begin VB.OptionButton Optsire 
               Caption         =   "Breed, Sire ID, Year"
               Height          =   195
               Index           =   3
               Left            =   225
               TabIndex        =   75
               Top             =   1200
               Width           =   1830
            End
            Begin VB.OptionButton Optsire 
               Caption         =   "Year, Breed, Sire ID"
               Height          =   195
               Index           =   2
               Left            =   225
               TabIndex        =   74
               Top             =   915
               Width           =   1830
            End
            Begin VB.OptionButton Optsire 
               Caption         =   "ID"
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   72
               Top             =   300
               Width           =   1830
            End
            Begin VB.OptionButton Optsire 
               Caption         =   "Year, Sire ID, Breed"
               Height          =   195
               Index           =   1
               Left            =   210
               TabIndex        =   73
               Top             =   600
               Width           =   1830
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sort Order of Cows"
            Height          =   1875
            Left            =   660
            TabIndex        =   65
            Top             =   375
            Width           =   2445
            Begin VB.OptionButton optcow 
               Caption         =   "Breed, Year, Cow ID"
               Height          =   195
               Index           =   4
               Left            =   225
               TabIndex        =   70
               Top             =   1500
               Width           =   1830
            End
            Begin VB.OptionButton optcow 
               Caption         =   "Breed, Cow ID, Year"
               Height          =   195
               Index           =   3
               Left            =   225
               TabIndex        =   69
               Top             =   1200
               Width           =   1830
            End
            Begin VB.OptionButton optcow 
               Caption         =   "Year, Breed, Cow ID"
               Height          =   195
               Index           =   2
               Left            =   225
               TabIndex        =   68
               Top             =   900
               Width           =   1830
            End
            Begin VB.OptionButton optcow 
               Caption         =   "ID"
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   66
               Top             =   300
               Width           =   1830
            End
            Begin VB.OptionButton optcow 
               Caption         =   "Year, Cow ID, Breed"
               Height          =   195
               Index           =   1
               Left            =   225
               TabIndex        =   67
               Top             =   600
               Width           =   1830
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Calf Misc Code Labels"
         Height          =   2940
         Left            =   -74910
         TabIndex        =   0
         Top             =   360
         Width           =   7095
         Begin VB.TextBox txtcalf1 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   3
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtcalf2 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   5
            Top             =   795
            Width           =   1335
         End
         Begin VB.TextBox txtcalf3 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   7
            Top             =   1125
            Width           =   1335
         End
         Begin VB.TextBox txtcalf4 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   10
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtcalf5 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   12
            Top             =   2115
            Width           =   1335
         End
         Begin VB.TextBox txtcalf6 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   14
            Top             =   2445
            Width           =   1335
         End
         Begin VB.TextBox txtback1 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   17
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtback2 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   19
            Top             =   795
            Width           =   1335
         End
         Begin VB.TextBox txtback3 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   21
            Top             =   1125
            Width           =   1335
         End
         Begin VB.TextBox txtfeed1 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtfeed2 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   33
            Top             =   795
            Width           =   1335
         End
         Begin VB.TextBox txtfeed3 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   35
            Top             =   1125
            Width           =   1335
         End
         Begin VB.TextBox txtrep1 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   24
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtrep2 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   26
            Top             =   2115
            Width           =   1335
         End
         Begin VB.TextBox txtrep3 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   28
            Top             =   2445
            Width           =   1335
         End
         Begin VB.TextBox txtcarc1 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   38
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtcarc2 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   40
            Top             =   2115
            Width           =   1335
         End
         Begin VB.TextBox txtcarc3 
            Height          =   285
            Left            =   5640
            MaxLength       =   12
            TabIndex        =   42
            Top             =   2445
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   105
            TabIndex        =   11
            Top             =   2115
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   105
            TabIndex        =   13
            Top             =   2445
            Width           =   495
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Profile"
            Height          =   255
            Left            =   645
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Weaning"
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Background"
            Height          =   255
            Left            =   3240
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Replacement"
            Height          =   255
            Left            =   3120
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Feed Lot"
            Height          =   255
            Left            =   5640
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Carcass"
            Height          =   255
            Left            =   5760
            TabIndex        =   36
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   480
            Width           =   500
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   795
            Width           =   500
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   2520
            TabIndex        =   20
            Top             =   1125
            Width           =   500
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   2520
            TabIndex        =   23
            Top             =   1800
            Width           =   500
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   2520
            TabIndex        =   25
            Top             =   2115
            Width           =   500
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   2520
            TabIndex        =   27
            Top             =   2445
            Width           =   500
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   5040
            TabIndex        =   30
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   5040
            TabIndex        =   32
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   5040
            TabIndex        =   34
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Left            =   5040
            TabIndex        =   37
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Left            =   5040
            TabIndex        =   39
            Top             =   2115
            Width           =   495
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   5040
            TabIndex        =   41
            Top             =   2445
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Epd Labels"
         Height          =   2910
         Left            =   -74910
         TabIndex        =   43
         Top             =   420
         Width           =   7095
         Begin VB.TextBox txtepd10 
            Height          =   285
            Left            =   2760
            MaxLength       =   12
            TabIndex        =   63
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtepd9 
            Height          =   285
            Left            =   2760
            MaxLength       =   12
            TabIndex        =   61
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtepd8 
            Height          =   285
            Left            =   2760
            MaxLength       =   12
            TabIndex        =   59
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtepd7 
            Height          =   285
            Left            =   2775
            MaxLength       =   12
            TabIndex        =   57
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtepd6 
            Height          =   285
            Left            =   2775
            MaxLength       =   12
            TabIndex        =   55
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtepd5 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   53
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtepd4 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   51
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtepd3 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   49
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtepd2 
            Height          =   285
            Left            =   645
            MaxLength       =   12
            TabIndex        =   47
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtepd1 
            Height          =   285
            Left            =   660
            MaxLength       =   12
            TabIndex        =   45
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 10"
            Height          =   255
            Left            =   1995
            TabIndex        =   62
            Top             =   1800
            Width           =   705
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 9"
            Height          =   255
            Left            =   1995
            TabIndex        =   60
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 8"
            Height          =   255
            Left            =   1995
            TabIndex        =   58
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 7"
            Height          =   255
            Left            =   1995
            TabIndex        =   56
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 6"
            Height          =   255
            Left            =   2010
            TabIndex        =   54
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 5"
            Height          =   255
            Left            =   105
            TabIndex        =   52
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 4"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 3"
            Height          =   255
            Left            =   105
            TabIndex        =   48
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 2"
            Height          =   255
            Left            =   105
            TabIndex        =   46
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Epd 1"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5985
      TabIndex        =   78
      Top             =   3795
      Width           =   1000
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4905
      TabIndex        =   77
      Top             =   3795
      Width           =   1000
   End
End
Attribute VB_Name = "frmpref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%
Dim oldid$
Dim tbData As Recordset

Private Sub valid_form(exitcode%)
    exitcode% = 0
    'If txtherdid.TEXT = "" Then
    '    Beep
    '    MsgBox "Herd ID Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
    '    txtherdid.SetFocus
    '    exitcode% = 1
    '    Exit Sub
    'End If
    Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
    Set tbData = DB.OpenRecordset("herd", dbOpenTable)
    tbData.Index = "primarykey"
    tbData.Seek "=", CBOHerdID.TEXT
    If tbData.NoMatch Then
       Beep
       MsgBox "Must have valid Herd ID", vbOKOnly + vbCritical, Me.Caption
       exitcode% = 1
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
       Exit Sub
    End If
    tbData.Close: Set tbData = Nothing
    DB.Close: Set DB = Nothing
End Sub

Private Sub save_information()
 Dim replace$, save As Boolean, i%
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("select * from prefer where herdid = '" & CBOHerdID.TEXT & "'", dbOpenDynaset)
 save = True
 If Not tbData.EOF Then
     tbData.Edit
  Else
   tbData.AddNew
 End If
 If save Then
   With tbData
     !herdid = CBOHerdID.TEXT
     !epdhead1 = txtepd1.TEXT: epdhead1$ = txtepd1.TEXT
     !epdhead2 = txtepd2.TEXT: epdhead2$ = txtepd2.TEXT
     !epdhead3 = txtepd3.TEXT: epdhead3$ = txtepd3.TEXT
     !epdhead4 = txtepd4.TEXT: epdhead4$ = txtepd4.TEXT
     !epdhead5 = txtepd5.TEXT: epdhead5$ = txtepd5.TEXT
     !epdhead6 = txtepd6.TEXT: epdhead6$ = txtepd6.TEXT
     !epdhead7 = txtepd7.TEXT: epdhead7$ = txtepd7.TEXT
     !epdhead8 = txtepd8.TEXT: epdhead8$ = txtepd8.TEXT
     !epdhead9 = txtepd9.TEXT: epdhead9$ = txtepd9.TEXT
     !epdhead10 = txtepd10.TEXT: epdhead10$ = txtepd10.TEXT
     !prof1 = txtcalf1.TEXT: calfhead(1) = txtcalf1.TEXT
     !prof2 = txtcalf2.TEXT: calfhead(2) = txtcalf2.TEXT
     !prof3 = txtcalf3.TEXT: calfhead(3) = txtcalf3.TEXT
     !wean1 = txtcalf4.TEXT: calfhead(4) = txtcalf4.TEXT
     !wean2 = txtcalf5.TEXT: calfhead(5) = txtcalf5.TEXT
     !wean3 = txtcalf6.TEXT: calfhead(6) = txtcalf6.TEXT
     !background1 = txtback1.TEXT: calfhead(7) = txtback1.TEXT
     !background2 = txtback2.TEXT: calfhead(8) = txtback2.TEXT
     !background3 = txtback3.TEXT: calfhead(9) = txtback3.TEXT
     !replacement1 = txtrep1.TEXT: calfhead(10) = txtrep1.TEXT
     !replacement2 = txtrep2.TEXT: calfhead(11) = txtrep2.TEXT
     !replacement3 = txtrep3.TEXT: calfhead(12) = txtrep3.TEXT
     !feedlot1 = txtfeed1.TEXT: calfhead(13) = txtfeed1.TEXT
     !feedlot2 = txtfeed2.TEXT: calfhead(14) = txtfeed2.TEXT
     !feedlot3 = txtfeed3.TEXT: calfhead(15) = txtfeed3.TEXT
     !carcass1 = txtcarc1.TEXT: calfhead(16) = txtcarc1.TEXT
     !carcass2 = txtcarc2.TEXT: calfhead(17) = txtcarc2.TEXT
     !carcass3 = txtcarc3.TEXT: calfhead(18) = txtcarc3.TEXT
     For i% = 0 To 4
       If optcow(i%).Value = True Then
         !cowsort = i%
         Select Case i%
          Case 0
            sortcows$ = "order by cowprof.cowid"
          Case 1
            sortcows$ = "order by cowprof.Birthdate, cowprof.cowid, cowprof.breed"
          Case 2
            sortcows$ = "order by cowprof.Birthdate,cowprof.breed, cowprof.cowid"
          Case 3
            sortcows$ = "order by cowprof.breed, cowprof.cowid, cowprof.Birthdate"
          Case 4
            sortcows$ = "order by cowprof.breed, cowprof.Birthdate, cowprof.cowid"
         End Select
       End If
       If Optsire(i%).Value = True Then
         !siresort = i%
         Select Case i%
          Case 0
            sortsires$ = "order by sireprof.sireid"
          Case 1
            sortsires$ = "order by sireprof.Birthdate, sireprof.sireid, sireprof.breed"
          Case 2
            sortsires$ = "order by sireprof.Birthdate, sireprof.breed, sireprof.sireid"
          Case 3
            sortsires$ = "order by sireprof.breed, sireprof.sireid, sireprof.Birthdate"
          Case 4
            sortsires$ = "order by sireprof.breed, sireprof.Birthdate, sireprof.sireid"
         End Select
       
       End If
     Next i%
     .Update
   End With
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 herdid$ = CBOHerdID.TEXT
 dirtyflag% = False
 Screen.MousePointer = vbDefault
End Sub

Private Sub LoadCBOHERD()
Dim DB As database
Dim RS As Recordset
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select distinct herdid from herd order by herdid", dbOpenSnapshot)
CBOHerdID.Clear
Do Until RS.EOF
   CBOHerdID.AddItem Field2Str(RS!herdid)
   RS.MoveNext
Loop
If CBOHerdID.ListCount > 0 Then CBOHerdID.ListIndex = 0
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End Sub


Private Sub Load_information()
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 'herdid = ReadDefaultHerdID
 Set tbData = DB.OpenRecordset("select * from prefer where herdid = '" & CBOHerdID.TEXT & "'", dbOpenSnapshot)
 If Not tbData.EOF Then
    
   'If Not IsNull(tbData!herdid) Then CBOHerdID.TEXT = tbData!herdid
   If Not IsNull(tbData!epdhead1) Then txtepd1.TEXT = tbData!epdhead1
   If Not IsNull(tbData!epdhead2) Then txtepd2.TEXT = tbData!epdhead2
   If Not IsNull(tbData!epdhead3) Then txtepd3.TEXT = tbData!epdhead3
   If Not IsNull(tbData!epdhead4) Then txtepd4.TEXT = tbData!epdhead4
   If Not IsNull(tbData!epdhead5) Then txtepd5.TEXT = tbData!epdhead5
   If Not IsNull(tbData!epdhead6) Then txtepd6.TEXT = tbData!epdhead6
   If Not IsNull(tbData!epdhead7) Then txtepd7.TEXT = tbData!epdhead7
   If Not IsNull(tbData!epdhead8) Then txtepd8.TEXT = tbData!epdhead8
   If Not IsNull(tbData!epdhead9) Then txtepd9.TEXT = tbData!epdhead9
   If Not IsNull(tbData!epdhead10) Then txtepd10.TEXT = tbData!epdhead10
     txtcalf1.TEXT = Field2Str(tbData!prof1)
     txtcalf2.TEXT = Field2Str(tbData!prof2)
     txtcalf3.TEXT = Field2Str(tbData!prof3)
     txtcalf4.TEXT = Field2Str(tbData!wean1)
     txtcalf5.TEXT = Field2Str(tbData!wean2)
     txtcalf6.TEXT = Field2Str(tbData!wean3)
     txtback1.TEXT = Field2Str(tbData!background1)
     txtback2.TEXT = Field2Str(tbData!background2)
     txtback3.TEXT = Field2Str(tbData!background3)
     txtrep1.TEXT = Field2Str(tbData!replacement1)
     txtrep2.TEXT = Field2Str(tbData!replacement2)
     txtrep3.TEXT = Field2Str(tbData!replacement3)
     txtfeed1.TEXT = Field2Str(tbData!feedlot1)
     txtfeed2.TEXT = Field2Str(tbData!feedlot2)
     txtfeed3.TEXT = Field2Str(tbData!feedlot3)
     txtcarc1.TEXT = Field2Str(tbData!carcass1)
     txtcarc2.TEXT = Field2Str(tbData!carcass2)
     txtcarc3.TEXT = Field2Str(tbData!carcass3)
     If Field2Str(tbData!cowsort) <> "" Then optcow(tbData!cowsort).Value = True
     If Field2Str(tbData!siresort) <> "" Then Optsire(tbData!siresort).Value = True
     If Field2Num(tbData!cowsort) = 0 Then optcow(0).Value = True
     If Field2Num(tbData!siresort) = 0 Then Optsire(0).Value = True
     If Field2Str(tbData!cowsort) = "" Then optcow(0).Value = True
     If Field2Str(tbData!siresort) = "" Then Optsire(0).Value = True
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub Init_Information()
' Call init_form(Me) ' Clear Text Boxes
Dim txtBox As Control
For Each txtBox In Me.Controls
   If TypeName(txtBox) = "TextBox" Then txtBox.TEXT = ""
Next
End Sub

Private Sub CBOHerdID_Click()
Call Init_Information
Call Load_information
End Sub

Private Sub CMDCancel_Click()
  Unload Me
End Sub

Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim TableName$(100)
 Call valid_form(exitcode%)
 If exitcode% = 1 Then Exit Sub
 Call save_information
 Unload Me
End Sub

Private Sub Form_Activate()
  SSTab1.Tab = 0
  txtcalf1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True

End Sub


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 Screen.MousePointer = vbHourglass
 Call LoadCBOHERD
 'Call Init_Information
 'Call Load_information
 Me.Tag = ""
 Screen.MousePointer = vbDefault
 Me.Enabled = True
 Me.Show
 txtepd1.SetFocus
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
 Set frmpref = Nothing

End Sub


Private Sub txtherdid_DblClick()

End Sub


