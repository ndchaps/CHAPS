VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCalf_Data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calf Information"
   ClientHeight    =   5535
   ClientLeft      =   5865
   ClientTop       =   5670
   ClientWidth     =   7905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   7905
   Begin VB.CommandButton cmddefault 
      Caption         =   "&Default"
      Height          =   375
      Left            =   135
      TabIndex        =   280
      Top             =   5115
      Width           =   1000
   End
   Begin VB.CommandButton CMDprev 
      Caption         =   "&Prev"
      Height          =   375
      Left            =   3075
      TabIndex        =   265
      Top             =   5115
      Width           =   1000
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4200
      TabIndex        =   266
      Top             =   5115
      Width           =   1000
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Height          =   385
      Left            =   5280
      TabIndex        =   264
      Top             =   5100
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   6360
      TabIndex        =   267
      Top             =   5100
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4830
      Left            =   45
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   210
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8520
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   529
      TabCaption(0)   =   "Birth"
      TabPicture(0)   =   "calfdata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Weaning"
      TabPicture(1)   =   "calfdata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Background"
      TabPicture(2)   =   "calfdata.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Replacement"
      TabPicture(3)   =   "calfdata.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Feed Lot"
      TabPicture(4)   =   "calfdata.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Carcass"
      TabPicture(5)   =   "calfdata.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame1 
         ClipControls    =   0   'False
         Height          =   4335
         Left            =   105
         TabIndex        =   129
         Top             =   375
         Width           =   7590
         Begin VB.ComboBox Cboeid 
            Height          =   315
            Left            =   5010
            TabIndex        =   11
            Text            =   "Cboeid"
            Top             =   1230
            Width           =   2085
         End
         Begin VB.TextBox txtpmisc3 
            Height          =   285
            Left            =   5010
            TabIndex        =   14
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox txtnotes 
            Height          =   705
            Left            =   5010
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Text            =   "calfdata.frx":00A8
            Top             =   2895
            Width           =   2490
         End
         Begin VB.TextBox TxtCowAge 
            Height          =   285
            Left            =   1470
            TabIndex        =   2
            Top             =   1170
            Width           =   675
         End
         Begin VB.TextBox txtregname 
            Height          =   285
            Left            =   5010
            MaxLength       =   40
            TabIndex        =   10
            Top             =   855
            Width           =   1900
         End
         Begin VB.TextBox Txtsireid 
            Height          =   285
            Left            =   1470
            MaxLength       =   8
            TabIndex        =   3
            Top             =   1500
            Width           =   1100
         End
         Begin VB.TextBox txtbirthwt 
            Height          =   285
            Left            =   1470
            TabIndex        =   7
            Top             =   2955
            Width           =   1140
         End
         Begin VB.TextBox txtid 
            Height          =   285
            Left            =   1470
            MaxLength       =   8
            TabIndex        =   0
            Top             =   420
            Width           =   1100
         End
         Begin VB.TextBox txtcowid 
            Height          =   285
            Left            =   1470
            MaxLength       =   8
            TabIndex        =   1
            Top             =   840
            Width           =   1100
         End
         Begin VB.TextBox txtregistration 
            Height          =   285
            Left            =   5010
            MaxLength       =   20
            TabIndex        =   9
            Top             =   420
            Width           =   1900
         End
         Begin VB.ComboBox cbosex 
            Height          =   315
            Left            =   1470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2190
            Width           =   615
         End
         Begin VB.ComboBox cboease 
            Height          =   315
            Left            =   1470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3285
            Width           =   615
         End
         Begin VB.TextBox txtcalfbreed 
            Height          =   285
            Left            =   1470
            MaxLength       =   8
            TabIndex        =   4
            Top             =   1845
            Width           =   1095
         End
         Begin VB.TextBox txtpmisc1 
            Height          =   285
            Left            =   5010
            TabIndex        =   12
            Top             =   1665
            Width           =   1095
         End
         Begin VB.TextBox txtpmisc2 
            Height          =   285
            Left            =   5010
            TabIndex        =   13
            Top             =   2040
            Width           =   1095
         End
         Begin MSMask.MaskEdBox dtebirth 
            Height          =   285
            Left            =   1470
            TabIndex        =   6
            Top             =   2580
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow Age"
            Height          =   315
            Left            =   510
            TabIndex        =   281
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Name"
            Height          =   255
            Index           =   6
            Left            =   3750
            TabIndex        =   271
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sire ID"
            Height          =   270
            Index           =   0
            Left            =   210
            TabIndex        =   131
            Top             =   1515
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   180
            Index           =   9
            Left            =   405
            TabIndex        =   136
            Top             =   2625
            Width           =   1005
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Weight"
            Height          =   195
            Index           =   10
            Left            =   405
            TabIndex        =   137
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sex"
            Height          =   180
            Index           =   11
            Left            =   375
            TabIndex        =   138
            Top             =   2250
            Width           =   1005
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calving Ease"
            Height          =   225
            Index           =   12
            Left            =   405
            TabIndex        =   139
            Top             =   3330
            Width           =   1005
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calf ID"
            Height          =   270
            Index           =   1
            Left            =   210
            TabIndex        =   130
            Top             =   420
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow ID"
            Height          =   270
            Index           =   2
            Left            =   210
            TabIndex        =   132
            Top             =   870
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registration #"
            Height          =   255
            Index           =   4
            Left            =   3735
            TabIndex        =   134
            Top             =   420
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Electronic ID"
            Height          =   255
            Index           =   5
            Left            =   3735
            TabIndex        =   135
            Top             =   1215
            Width           =   1200
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Profile Notes"
            Height          =   375
            Index           =   13
            Left            =   3900
            TabIndex        =   143
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calf Breed"
            Height          =   270
            Index           =   3
            Left            =   195
            TabIndex        =   133
            Top             =   1845
            Width           =   1200
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Index           =   0
            Left            =   3615
            TabIndex        =   140
            Top             =   1665
            Width           =   1365
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Index           =   1
            Left            =   3615
            TabIndex        =   141
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   2
            Left            =   3615
            TabIndex        =   142
            Top             =   2415
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   135
         Left            =   -74400
         TabIndex        =   269
         Top             =   900
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Height          =   4335
         Left            =   -74925
         TabIndex        =   144
         Top             =   375
         Width           =   7605
         Begin VB.CommandButton cmdapply 
            Caption         =   "Apply"
            Height          =   315
            Left            =   3600
            TabIndex        =   157
            Tag             =   "Weaning"
            Top             =   3240
            Width           =   615
         End
         Begin VB.CheckBox chkwstatus 
            Caption         =   "Status"
            Height          =   225
            Left            =   135
            TabIndex        =   16
            Top             =   150
            Width           =   900
         End
         Begin VB.CheckBox chkedit 
            Caption         =   "Edit"
            Height          =   195
            Left            =   3600
            TabIndex        =   152
            Tag             =   "Weaning"
            Top             =   3960
            Width           =   765
         End
         Begin VB.Frame Frame17 
            Caption         =   "Calculated Data"
            Height          =   1155
            Left            =   225
            TabIndex        =   153
            Top             =   3090
            Width           =   3210
            Begin VB.TextBox txtratio 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   26
               Top             =   765
               Width           =   1260
            End
            Begin VB.TextBox txtadj205 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   25
               Top             =   465
               Width           =   1260
            End
            Begin VB.TextBox txtframe 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   24
               Top             =   165
               Width           =   1260
            End
            Begin VB.Label lblratio 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ratio"
               Enabled         =   0   'False
               Height          =   225
               Left            =   210
               TabIndex        =   156
               Top             =   810
               Width           =   1500
            End
            Begin VB.Label lbl205 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Adj. 205 Day Wt."
               Enabled         =   0   'False
               Height          =   220
               Left            =   210
               TabIndex        =   155
               Top             =   510
               Width           =   1500
            End
            Begin VB.Label lblscore 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calf Frame Score"
               Enabled         =   0   'False
               Height          =   225
               Left            =   195
               TabIndex        =   154
               Top             =   210
               Width           =   1500
            End
         End
         Begin VB.ComboBox Cbograde 
            Height          =   315
            Left            =   1755
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Weaning"
            Top             =   2715
            Width           =   900
         End
         Begin VB.TextBox txtactwt 
            Height          =   285
            Left            =   1755
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Weaning"
            Top             =   450
            Width           =   1140
         End
         Begin VB.ComboBox cbomancode 
            Height          =   315
            Left            =   1755
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Tag             =   "Weaning"
            Top             =   1200
            Width           =   660
         End
         Begin VB.TextBox txtgroup 
            Height          =   285
            Left            =   1755
            MaxLength       =   1
            TabIndex        =   22
            Tag             =   "Weaning"
            Top             =   2370
            Width           =   420
         End
         Begin VB.TextBox txtmisc1 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   27
            Tag             =   "Weaning"
            Top             =   480
            Width           =   1300
         End
         Begin VB.TextBox txtmisc2 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "Weaning"
            Top             =   855
            Width           =   1300
         End
         Begin VB.TextBox txtmisc3 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   29
            Tag             =   "Weaning"
            Top             =   1230
            Width           =   1300
         End
         Begin VB.TextBox txtmisc4 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   30
            Tag             =   "Weaning"
            Top             =   1605
            Width           =   1300
         End
         Begin VB.TextBox txtmisc5 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Weaning"
            Top             =   1980
            Width           =   1300
         End
         Begin VB.TextBox txtmisc6 
            Height          =   285
            Left            =   5115
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "Weaning"
            Top             =   2355
            Width           =   1300
         End
         Begin VB.TextBox txthipht 
            Height          =   285
            Left            =   1750
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "Weaning"
            Top             =   1605
            Width           =   1125
         End
         Begin VB.TextBox txtwnotes 
            Height          =   1515
            Left            =   5100
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Tag             =   "Weaning"
            Text            =   "calfdata.frx":00AC
            Top             =   2745
            Width           =   2100
         End
         Begin MSMask.MaskEdBox Dtewt 
            Height          =   285
            Left            =   1755
            TabIndex        =   18
            Tag             =   "Weaning"
            Top             =   840
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin MSMask.MaskEdBox dtemeas 
            Height          =   285
            Left            =   1755
            TabIndex        =   21
            Tag             =   "Weaning"
            Top             =   1965
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Weight"
            Height          =   240
            Index           =   14
            Left            =   360
            TabIndex        =   145
            Tag             =   "Weaning"
            Top             =   450
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Weighed"
            Height          =   255
            Index           =   15
            Left            =   300
            TabIndex        =   146
            Tag             =   "Weaning"
            Top             =   840
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Manage Code"
            Height          =   285
            Index           =   23
            Left            =   300
            TabIndex        =   147
            Tag             =   "Weaning"
            Top             =   1230
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contemp. Group"
            Height          =   270
            Index           =   26
            Left            =   300
            TabIndex        =   150
            Tag             =   "Weaning"
            Top             =   2385
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Muscle Grade"
            Height          =   255
            Index           =   27
            Left            =   285
            TabIndex        =   151
            Tag             =   "Weaning"
            Top             =   2760
            Width           =   1350
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Index           =   3
            Left            =   3105
            TabIndex        =   158
            Tag             =   "Weaning"
            Top             =   510
            Width           =   1905
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Index           =   4
            Left            =   3105
            TabIndex        =   159
            Tag             =   "Weaning"
            Top             =   855
            Width           =   1905
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   5
            Left            =   3105
            TabIndex        =   160
            Tag             =   "Weaning"
            Top             =   1230
            Width           =   1905
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 4"
            Height          =   255
            Index           =   6
            Left            =   3105
            TabIndex        =   161
            Tag             =   "Weaning"
            Top             =   1605
            Width           =   1905
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 5"
            Height          =   255
            Index           =   7
            Left            =   3105
            TabIndex        =   162
            Tag             =   "Weaning"
            Top             =   1980
            Width           =   1905
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 6"
            Height          =   255
            Index           =   8
            Left            =   3105
            TabIndex        =   163
            Tag             =   "Weaning"
            Top             =   2355
            Width           =   1905
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hip Height"
            Height          =   255
            Index           =   24
            Left            =   300
            TabIndex        =   148
            Tag             =   "Weaning"
            Top             =   1605
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Measured"
            Height          =   255
            Index           =   25
            Left            =   300
            TabIndex        =   149
            Tag             =   "Weaning"
            Top             =   1980
            Width           =   1350
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Weaning Notes"
            Height          =   255
            Index           =   22
            Left            =   3780
            TabIndex        =   164
            Tag             =   "Weaning"
            Top             =   2745
            Width           =   1200
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4335
         Left            =   -74910
         TabIndex        =   268
         Top             =   345
         Width           =   7620
         Begin VB.CommandButton Command1 
            Caption         =   "Apply"
            Height          =   315
            Left            =   120
            TabIndex        =   180
            Tag             =   "Back"
            Top             =   1440
            Width           =   615
         End
         Begin VB.CheckBox Chkbackstatus 
            Caption         =   "Status"
            Height          =   225
            Left            =   120
            TabIndex        =   34
            Top             =   210
            Width           =   900
         End
         Begin VB.TextBox txtcontgrp 
            Height          =   285
            Left            =   2700
            MaxLength       =   1
            TabIndex        =   51
            Tag             =   "Back"
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtbackmisc1 
            Height          =   285
            Left            =   5145
            MaxLength       =   10
            TabIndex        =   47
            Tag             =   "Back"
            Top             =   2070
            Width           =   1335
         End
         Begin VB.TextBox txtbackmisc2 
            Height          =   285
            Left            =   5145
            MaxLength       =   10
            TabIndex        =   48
            Tag             =   "Back"
            Top             =   2430
            Width           =   1335
         End
         Begin VB.TextBox txtbackmisc3 
            Height          =   285
            Left            =   5145
            MaxLength       =   10
            TabIndex        =   49
            Tag             =   "Back"
            Top             =   2790
            Width           =   1335
         End
         Begin VB.TextBox txtbnotes 
            Height          =   870
            Left            =   5130
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Tag             =   "Back"
            Text            =   "calfdata.frx":00B0
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Frame Frame7 
            Caption         =   "Receiving"
            Height          =   1815
            Left            =   1095
            TabIndex        =   165
            Top             =   135
            Width           =   3135
            Begin VB.TextBox txtbackrecframe 
               Height          =   285
               Left            =   1755
               TabIndex        =   38
               Tag             =   "Back"
               Top             =   1290
               Width           =   1250
            End
            Begin VB.TextBox txtbackhh 
               Height          =   285
               Left            =   1755
               TabIndex        =   37
               Tag             =   "Back"
               Top             =   945
               Width           =   1250
            End
            Begin VB.TextBox txtbackrecwt 
               Height          =   285
               Left            =   1755
               TabIndex        =   36
               Tag             =   "Back"
               Top             =   585
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtebackrec 
               Height          =   285
               Left            =   1755
               TabIndex        =   35
               Tag             =   "Back"
               Top             =   210
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   166
               Tag             =   "Back"
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   255
               Left            =   120
               TabIndex        =   169
               Tag             =   "Back"
               Top             =   1290
               Width           =   1500
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip Height"
               Height          =   255
               Left            =   120
               TabIndex        =   168
               Tag             =   "Back"
               Top             =   945
               Width           =   1500
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   120
               TabIndex        =   167
               Tag             =   "Back"
               Top             =   585
               Width           =   1500
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Interim"
            Height          =   1935
            Left            =   1095
            TabIndex        =   170
            Top             =   1935
            Width           =   2520
            Begin VB.TextBox txtbackintframe 
               Height          =   285
               Left            =   1140
               TabIndex        =   42
               Tag             =   "Back"
               Top             =   1410
               Width           =   1250
            End
            Begin VB.TextBox txtbackinthh 
               Height          =   285
               Left            =   1140
               TabIndex        =   41
               Tag             =   "Back"
               Top             =   1050
               Width           =   1250
            End
            Begin VB.TextBox txtbackintwt 
               Height          =   285
               Left            =   1140
               TabIndex        =   40
               Tag             =   "Back"
               Top             =   690
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtebackint 
               Height          =   285
               Left            =   1140
               TabIndex        =   39
               Tag             =   "Back"
               Top             =   315
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   255
               Left            =   60
               TabIndex        =   174
               Tag             =   "Back"
               Top             =   1380
               Width           =   930
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip Height"
               Height          =   255
               Left            =   60
               TabIndex        =   173
               Tag             =   "Back"
               Top             =   1035
               Width           =   960
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   180
               TabIndex        =   172
               Tag             =   "Back"
               Top             =   705
               Width           =   810
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   270
               TabIndex        =   171
               Tag             =   "Back"
               Top             =   330
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Final"
            Height          =   1815
            Left            =   4320
            TabIndex        =   175
            Top             =   135
            Width           =   3135
            Begin VB.TextBox txtbackfinframe 
               Height          =   285
               Left            =   1725
               TabIndex        =   46
               Tag             =   "Back"
               Top             =   1290
               Width           =   1250
            End
            Begin VB.TextBox txtbackfinhh 
               Height          =   285
               Left            =   1725
               TabIndex        =   45
               Tag             =   "Back"
               Top             =   945
               Width           =   1250
            End
            Begin VB.TextBox txtbackfinwt 
               Height          =   285
               Left            =   1725
               TabIndex        =   44
               Tag             =   "Back"
               Top             =   585
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtebackfinal 
               Height          =   285
               Left            =   1725
               TabIndex        =   43
               Tag             =   "Back"
               Top             =   210
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   255
               Left            =   120
               TabIndex        =   179
               Tag             =   "Back"
               Top             =   1290
               Width           =   1500
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip Height"
               Height          =   255
               Left            =   120
               TabIndex        =   178
               Tag             =   "Back"
               Top             =   945
               Width           =   1500
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   120
               TabIndex        =   177
               Tag             =   "Back"
               Top             =   585
               Width           =   1500
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   176
               Tag             =   "Back"
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Label Label96 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contemporary Group"
            Height          =   255
            Left            =   975
            TabIndex        =   181
            Tag             =   "Back"
            Top             =   3975
            Width           =   1605
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Index           =   9
            Left            =   3570
            TabIndex        =   182
            Tag             =   "Back"
            Top             =   2070
            Width           =   1500
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Index           =   10
            Left            =   3555
            TabIndex        =   183
            Tag             =   "Back"
            Top             =   2430
            Width           =   1500
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   11
            Left            =   3555
            TabIndex        =   184
            Tag             =   "Back"
            Top             =   2775
            Width           =   1500
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Background Notes"
            Height          =   255
            Left            =   3510
            TabIndex        =   185
            Tag             =   "Back"
            Top             =   3120
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4395
         Left            =   -74910
         TabIndex        =   186
         Top             =   330
         Width           =   7635
         Begin VB.TextBox txtweight365 
            Height          =   285
            Left            =   6120
            MaxLength       =   10
            TabIndex        =   81
            Tag             =   "REPL"
            Top             =   2835
            Width           =   1250
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Apply"
            Height          =   345
            Left            =   2625
            TabIndex        =   191
            Tag             =   "REPL"
            Top             =   2865
            Width           =   720
         End
         Begin VB.CheckBox chkrstatus 
            Caption         =   "Status"
            Height          =   225
            Left            =   60
            TabIndex        =   52
            Top             =   165
            Width           =   765
         End
         Begin VB.Frame Frame15 
            Caption         =   "Yearling"
            Height          =   2655
            Left            =   5100
            TabIndex        =   214
            Top             =   165
            Width           =   2415
            Begin VB.TextBox txt365marb 
               Height          =   285
               Left            =   1035
               TabIndex        =   76
               Tag             =   "REPL"
               Top             =   2295
               Width           =   1250
            End
            Begin VB.TextBox txt365rib 
               Height          =   285
               Left            =   1035
               TabIndex        =   75
               Tag             =   "REPL"
               Top             =   1995
               Width           =   1250
            End
            Begin VB.TextBox txt365fat 
               Height          =   285
               Left            =   1035
               TabIndex        =   74
               Tag             =   "REPL"
               Top             =   1680
               Width           =   1250
            End
            Begin VB.TextBox txt365cond 
               Height          =   285
               Left            =   1035
               TabIndex        =   71
               Tag             =   "REPL"
               Top             =   780
               Width           =   1250
            End
            Begin VB.TextBox txt365wt 
               Height          =   285
               Left            =   1035
               TabIndex        =   70
               Tag             =   "REPL"
               Top             =   480
               Width           =   1250
            End
            Begin VB.TextBox txt365hh 
               Height          =   285
               Left            =   1035
               MaxLength       =   8
               TabIndex        =   72
               Tag             =   "REPL"
               Top             =   1080
               Width           =   1250
            End
            Begin VB.TextBox txt365score 
               Height          =   285
               Left            =   1035
               MaxLength       =   10
               TabIndex        =   73
               Tag             =   "REPL"
               Top             =   1365
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dte365 
               Height          =   285
               Left            =   1035
               TabIndex        =   69
               Tag             =   "REPL"
               Top             =   165
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label106 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   30
               TabIndex        =   222
               Tag             =   "REPL"
               Top             =   2295
               Width           =   945
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   30
               TabIndex        =   221
               Tag             =   "REPL"
               Top             =   1995
               Width           =   945
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   30
               TabIndex        =   220
               Tag             =   "REPL"
               Top             =   1680
               Width           =   945
            End
            Begin VB.Label Label103 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   15
               TabIndex        =   216
               Tag             =   "REPL"
               Top             =   480
               Width           =   945
            End
            Begin VB.Label Label102 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   30
               TabIndex        =   217
               Tag             =   "REPL"
               Top             =   795
               Width           =   945
            End
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip height"
               Height          =   255
               Left            =   30
               TabIndex        =   218
               Tag             =   "REPL"
               Top             =   1080
               Width           =   945
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   225
               Left            =   30
               TabIndex        =   215
               Tag             =   "REPL"
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   285
               Left            =   30
               TabIndex        =   219
               Tag             =   "REPL"
               Top             =   1365
               Width           =   945
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Interim"
            Height          =   2655
            Left            =   2595
            TabIndex        =   205
            Top             =   165
            Width           =   2415
            Begin VB.TextBox txtintmarb 
               Height          =   285
               Left            =   1035
               TabIndex        =   68
               Tag             =   "REPL"
               Top             =   2295
               Width           =   1245
            End
            Begin VB.TextBox txtintrib 
               Height          =   285
               Left            =   1035
               TabIndex        =   67
               Tag             =   "REPL"
               Top             =   1995
               Width           =   1245
            End
            Begin VB.TextBox txtintfat 
               Height          =   285
               Left            =   1035
               TabIndex        =   66
               Tag             =   "REPL"
               Top             =   1680
               Width           =   1245
            End
            Begin VB.TextBox txtintcond 
               Height          =   285
               Left            =   1035
               TabIndex        =   63
               Tag             =   "REPL"
               Top             =   780
               Width           =   1245
            End
            Begin VB.TextBox txtintwt 
               Height          =   285
               Left            =   1035
               TabIndex        =   62
               Tag             =   "REPL"
               Top             =   465
               Width           =   1245
            End
            Begin VB.TextBox txtinthh 
               Height          =   285
               Left            =   1035
               MaxLength       =   8
               TabIndex        =   64
               Tag             =   "REPL"
               Top             =   1080
               Width           =   1245
            End
            Begin VB.TextBox txtintscore 
               Height          =   285
               Left            =   1035
               MaxLength       =   10
               TabIndex        =   65
               Tag             =   "REPL"
               Top             =   1380
               Width           =   1245
            End
            Begin MSMask.MaskEdBox dteint 
               Height          =   285
               Left            =   1035
               TabIndex        =   61
               Tag             =   "REPL"
               Top             =   150
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   30
               TabIndex        =   213
               Tag             =   "REPL"
               Top             =   2295
               Width           =   945
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   30
               TabIndex        =   212
               Tag             =   "REPL"
               Top             =   1995
               Width           =   945
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   30
               TabIndex        =   211
               Tag             =   "REPL"
               Top             =   1680
               Width           =   945
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   30
               TabIndex        =   207
               Tag             =   "REPL"
               Top             =   480
               Width           =   945
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   30
               TabIndex        =   208
               Tag             =   "REPL"
               Top             =   795
               Width           =   945
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip height"
               Height          =   210
               Left            =   30
               TabIndex        =   209
               Tag             =   "REPL"
               Top             =   1110
               Width           =   945
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   225
               Left            =   30
               TabIndex        =   206
               Tag             =   "REPL"
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   285
               Left            =   30
               TabIndex        =   210
               Tag             =   "REPL"
               Top             =   1395
               Width           =   945
            End
         End
         Begin VB.TextBox txtymisc1 
            Height          =   285
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   78
            Tag             =   "REPL"
            Top             =   3195
            Width           =   1100
         End
         Begin VB.TextBox txtpelvicsz 
            Height          =   285
            Left            =   6120
            MaxLength       =   10
            TabIndex        =   84
            Tag             =   "REPL"
            Top             =   3765
            Width           =   1250
         End
         Begin VB.TextBox txtscrotumcir 
            Height          =   285
            Left            =   6120
            MaxLength       =   10
            TabIndex        =   82
            Tag             =   "REPL"
            Top             =   3135
            Width           =   1250
         End
         Begin VB.Frame Frame10 
            Caption         =   "Receiving"
            Height          =   2655
            Left            =   75
            TabIndex        =   196
            Top             =   405
            Width           =   2415
            Begin VB.TextBox txtrecmarb 
               Height          =   285
               Left            =   1035
               TabIndex        =   60
               Tag             =   "REPL"
               Top             =   2300
               Width           =   1250
            End
            Begin VB.TextBox txtrecrib 
               Height          =   285
               Left            =   1035
               TabIndex        =   59
               Tag             =   "REPL"
               Top             =   1990
               Width           =   1250
            End
            Begin VB.TextBox txtrecfat 
               Height          =   285
               Left            =   1035
               TabIndex        =   58
               Tag             =   "REPL"
               Top             =   1695
               Width           =   1250
            End
            Begin VB.TextBox txtyframe 
               Height          =   285
               Left            =   1035
               MaxLength       =   10
               TabIndex        =   57
               Tag             =   "REPL"
               Top             =   1380
               Width           =   1250
            End
            Begin VB.TextBox txtyhipht 
               Height          =   285
               Left            =   1035
               MaxLength       =   8
               TabIndex        =   56
               Tag             =   "REPL"
               Top             =   1095
               Width           =   1250
            End
            Begin VB.TextBox txtrecwt 
               Height          =   285
               Left            =   1035
               TabIndex        =   54
               Tag             =   "REPL"
               Top             =   480
               Width           =   1250
            End
            Begin VB.TextBox txtreccond 
               Height          =   285
               Left            =   1035
               TabIndex        =   55
               Tag             =   "REPL"
               Top             =   800
               Width           =   1250
            End
            Begin MSMask.MaskEdBox Dterec 
               Height          =   285
               Left            =   1035
               TabIndex        =   53
               Tag             =   "REPL"
               Top             =   165
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   50
               TabIndex        =   204
               Tag             =   "REPL"
               Top             =   2300
               Width           =   945
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   50
               TabIndex        =   203
               Tag             =   "REPL"
               Top             =   1990
               Width           =   945
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   50
               TabIndex        =   202
               Tag             =   "REPL"
               Top             =   1680
               Width           =   945
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   285
               Left            =   45
               TabIndex        =   201
               Tag             =   "REPL"
               Top             =   1395
               Width           =   945
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   225
               Left            =   50
               TabIndex        =   197
               Tag             =   "REPL"
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip height"
               Height          =   255
               Left            =   45
               TabIndex        =   200
               Tag             =   "REPL"
               Top             =   1110
               Width           =   945
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   50
               TabIndex        =   199
               Tag             =   "REPL"
               Top             =   800
               Width           =   945
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   50
               TabIndex        =   198
               Tag             =   "REPL"
               Top             =   500
               Width           =   945
            End
         End
         Begin VB.TextBox txtymisc2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   79
            Tag             =   "REPL"
            Top             =   3570
            Width           =   1095
         End
         Begin VB.TextBox txtymisc3 
            Height          =   285
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   80
            Tag             =   "REPL"
            Top             =   3945
            Width           =   1095
         End
         Begin VB.TextBox txtynotes 
            Height          =   960
            Left            =   135
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Tag             =   "REPL"
            Top             =   3330
            Width           =   2265
         End
         Begin MSMask.MaskEdBox dtescr 
            Height          =   285
            Left            =   6120
            TabIndex        =   83
            Tag             =   "REPL"
            Top             =   3450
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin MSMask.MaskEdBox dtepelv 
            Height          =   285
            Left            =   6120
            TabIndex        =   85
            Tag             =   "REPL"
            Top             =   4065
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "365 D Weight"
            Height          =   255
            Left            =   5070
            TabIndex        =   272
            Tag             =   "REPL"
            Top             =   2865
            Width           =   1005
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   285
            Index           =   12
            Left            =   2475
            TabIndex        =   192
            Tag             =   "REPL"
            Top             =   3225
            Width           =   1170
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pelvic Area"
            Height          =   285
            Left            =   5070
            TabIndex        =   189
            Tag             =   "REPL"
            Top             =   3780
            Width           =   975
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Scrotum Cir"
            Height          =   255
            Left            =   5160
            TabIndex        =   187
            Tag             =   "REPL"
            Top             =   3150
            Width           =   900
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Scrotum Date"
            Height          =   255
            Left            =   5085
            TabIndex        =   188
            Tag             =   "REPL"
            Top             =   3465
            Width           =   975
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pelvic Date"
            Height          =   255
            Left            =   5070
            TabIndex        =   190
            Tag             =   "REPL"
            Top             =   4095
            Width           =   975
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   210
            Index           =   13
            Left            =   2460
            TabIndex        =   193
            Tag             =   "REPL"
            Top             =   3615
            Width           =   1170
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   14
            Left            =   2460
            TabIndex        =   194
            Tag             =   "REPL"
            Top             =   3975
            Width           =   1170
         End
         Begin VB.Label Label90 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Replacement Notes"
            Height          =   255
            Left            =   15
            TabIndex        =   195
            Tag             =   "REPL"
            Top             =   3090
            Width           =   1545
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4455
         Left            =   -74910
         TabIndex        =   244
         Top             =   345
         Width           =   7560
         Begin VB.ComboBox CBOConformance 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Tag             =   "Carc"
            Top             =   3480
            Width           =   1245
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Apply"
            Height          =   345
            Left            =   360
            TabIndex        =   279
            Tag             =   "Carc"
            Top             =   2580
            Width           =   705
         End
         Begin VB.TextBox txtcarcmisc3 
            Height          =   285
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   127
            Tag             =   "Carc"
            Top             =   3945
            Width           =   1300
         End
         Begin VB.Frame Frame19 
            Caption         =   "Quality Grade"
            Height          =   2325
            Left            =   4260
            TabIndex        =   251
            Top             =   135
            Width           =   3120
            Begin VB.ComboBox CBOCarcQual 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   117
               Tag             =   "Carc"
               Top             =   300
               Width           =   1200
            End
            Begin VB.TextBox txtqmaturity 
               Height          =   285
               Left            =   1485
               MaxLength       =   15
               TabIndex        =   121
               Tag             =   "Carc"
               Top             =   1935
               Width           =   1200
            End
            Begin VB.TextBox txtqlean 
               Height          =   315
               Left            =   1485
               MaxLength       =   15
               TabIndex        =   120
               Tag             =   "Carc"
               Top             =   1545
               Width           =   1200
            End
            Begin VB.TextBox txtqcolor 
               Height          =   285
               Left            =   1485
               MaxLength       =   15
               TabIndex        =   119
               Tag             =   "Carc"
               Top             =   1155
               Width           =   1200
            End
            Begin VB.TextBox txtqscore 
               Height          =   285
               Left            =   1485
               MaxLength       =   15
               TabIndex        =   118
               Tag             =   "Carc"
               Top             =   780
               Width           =   1200
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Maturity"
               Height          =   300
               Left            =   195
               TabIndex        =   256
               Tag             =   "Carc"
               Top             =   1935
               Width           =   1200
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Texture of Lean"
               Height          =   225
               Left            =   195
               TabIndex        =   255
               Tag             =   "Carc"
               Top             =   1545
               Width           =   1200
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
               Height          =   270
               Left            =   195
               TabIndex        =   254
               Tag             =   "Carc"
               Top             =   1155
               Width           =   1200
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling Score"
               Height          =   255
               Left            =   195
               TabIndex        =   253
               Tag             =   "Carc"
               Top             =   810
               Width           =   1200
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Quality Grade"
               Height          =   255
               Left            =   195
               TabIndex        =   252
               Tag             =   "Carc"
               Top             =   390
               Width           =   1200
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Yield Grade"
            Height          =   2325
            Left            =   1005
            TabIndex        =   245
            Top             =   135
            Width           =   3120
            Begin VB.TextBox txtcarcyield 
               Height          =   285
               Left            =   1455
               TabIndex        =   111
               Tag             =   "Carc"
               Top             =   390
               Width           =   1200
            End
            Begin VB.TextBox txtcarccarcwt 
               Height          =   285
               Left            =   1455
               TabIndex        =   112
               Tag             =   "Carc"
               Top             =   780
               Width           =   1200
            End
            Begin VB.TextBox txtcarcfatthick 
               Height          =   285
               Left            =   1455
               TabIndex        =   113
               Tag             =   "Carc"
               Top             =   1155
               Width           =   1200
            End
            Begin VB.TextBox txtcarckph 
               Height          =   285
               Left            =   1455
               TabIndex        =   114
               Tag             =   "Carc"
               Top             =   1545
               Width           =   1200
            End
            Begin VB.TextBox txtcarcrib 
               Height          =   285
               Left            =   1455
               TabIndex        =   115
               Tag             =   "Carc"
               Top             =   1935
               Width           =   1200
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Yield Grade"
               Height          =   255
               Left            =   375
               TabIndex        =   246
               Tag             =   "Carc"
               Top             =   390
               Width           =   1005
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hot Carcass Wt"
               Height          =   255
               Left            =   135
               TabIndex        =   247
               Tag             =   "Carc"
               Top             =   780
               Width           =   1230
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fat Thickness"
               Height          =   255
               Left            =   210
               TabIndex        =   248
               Tag             =   "Carc"
               Top             =   1155
               Width           =   1155
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kidney (KPH)"
               Height          =   255
               Left            =   315
               TabIndex        =   249
               Tag             =   "Carc"
               Top             =   1545
               Width           =   1050
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   495
               TabIndex        =   250
               Tag             =   "Carc"
               Top             =   1935
               Width           =   885
            End
         End
         Begin VB.TextBox txtcarcmisc1 
            Height          =   285
            Left            =   3975
            MaxLength       =   10
            TabIndex        =   125
            Tag             =   "Carc"
            Top             =   3030
            Width           =   1300
         End
         Begin VB.TextBox txtcarcmisc2 
            Height          =   285
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   126
            Tag             =   "Carc"
            Top             =   3480
            Width           =   1300
         End
         Begin VB.CheckBox chkcstatus 
            Caption         =   "Status"
            Height          =   225
            Left            =   165
            TabIndex        =   110
            Top             =   225
            Width           =   900
         End
         Begin VB.TextBox txtcarcnotes 
            Height          =   1665
            Left            =   5415
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   128
            Tag             =   "Carc"
            Top             =   2595
            Width           =   2040
         End
         Begin VB.TextBox txtcarcmuscle 
            Height          =   285
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   124
            Tag             =   "Carc"
            Top             =   3945
            Width           =   1245
         End
         Begin MSMask.MaskEdBox dtecarc 
            Height          =   285
            Left            =   1230
            TabIndex        =   122
            Tag             =   "Carc"
            Top             =   3030
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "-"
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   20
            Left            =   2565
            TabIndex        =   262
            Tag             =   "Carc"
            Top             =   3945
            Width           =   1350
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Index           =   18
            Left            =   2565
            TabIndex        =   260
            Tag             =   "Carc"
            Top             =   3030
            Width           =   1350
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Index           =   19
            Left            =   2565
            TabIndex        =   261
            Tag             =   "Carc"
            Top             =   3480
            Width           =   1350
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Left            =   2535
            TabIndex        =   270
            Top             =   39451
            Width           =   1005
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carcass Notes"
            Height          =   255
            Left            =   4185
            TabIndex        =   263
            Tag             =   "Carc"
            Top             =   2595
            Width           =   1125
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Conformance"
            Height          =   255
            Left            =   105
            TabIndex        =   258
            Tag             =   "Carc"
            Top             =   3480
            Width           =   1065
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Muscle Score"
            Height          =   255
            Left            =   75
            TabIndex        =   259
            Tag             =   "Carc"
            Top             =   3945
            Width           =   1095
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Carcass Date"
            Height          =   255
            Left            =   105
            TabIndex        =   257
            Tag             =   "Carc"
            Top             =   3030
            Width           =   1050
         End
      End
      Begin VB.Frame Frame11 
         Height          =   4365
         Left            =   -74895
         TabIndex        =   223
         Top             =   345
         Width           =   7545
         Begin VB.CommandButton Command4 
            Caption         =   "Apply"
            Height          =   345
            Left            =   180
            TabIndex        =   282
            Tag             =   "F"
            Top             =   3300
            Width           =   705
         End
         Begin VB.CheckBox Chkfstatus 
            Caption         =   "Status"
            Height          =   225
            Left            =   105
            TabIndex        =   86
            Top             =   210
            Width           =   900
         End
         Begin VB.TextBox txtfeedmisc3 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   108
            Tag             =   "F"
            Top             =   3975
            Width           =   1300
         End
         Begin VB.TextBox txtfeedmisc2 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   107
            Tag             =   "F"
            Top             =   3600
            Width           =   1300
         End
         Begin VB.TextBox txtfeedmisc1 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   106
            Tag             =   "F"
            Top             =   3240
            Width           =   1300
         End
         Begin VB.TextBox txtfnotes 
            Height          =   975
            Left            =   4800
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Tag             =   "F"
            Text            =   "calfdata.frx":00B4
            Top             =   3225
            Width           =   2295
         End
         Begin VB.Frame Frame13 
            Caption         =   "Final"
            Height          =   2085
            Left            =   5190
            TabIndex        =   234
            Top             =   510
            Width           =   2265
            Begin VB.TextBox txtfinfinrea 
               Height          =   285
               Left            =   885
               TabIndex        =   104
               Tag             =   "F"
               Top             =   1260
               Width           =   1250
            End
            Begin VB.TextBox txtfinfinmarbl 
               Height          =   285
               Left            =   885
               TabIndex        =   105
               Tag             =   "F"
               Top             =   1605
               Width           =   1250
            End
            Begin VB.TextBox txtfinfinfat 
               Height          =   285
               Left            =   885
               TabIndex        =   103
               Tag             =   "F"
               Top             =   915
               Width           =   1250
            End
            Begin VB.TextBox txtfinfinwt 
               Height          =   285
               Left            =   880
               TabIndex        =   102
               Tag             =   "F"
               Top             =   550
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtefinfin 
               Height          =   285
               Left            =   885
               TabIndex        =   101
               Tag             =   "F"
               Top             =   195
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   105
               TabIndex        =   239
               Tag             =   "F"
               Top             =   1260
               Width           =   705
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   90
               TabIndex        =   238
               Tag             =   "F"
               Top             =   1605
               Width           =   705
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   90
               TabIndex        =   237
               Tag             =   "F"
               Top             =   915
               Width           =   705
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   100
               TabIndex        =   236
               Tag             =   "F"
               Top             =   550
               Width           =   705
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   100
               TabIndex        =   235
               Tag             =   "F"
               Top             =   200
               Width           =   705
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Interim"
            Height          =   2655
            Left            =   2715
            TabIndex        =   229
            Top             =   510
            Width           =   2400
            Begin VB.TextBox txtflscore 
               Height          =   285
               Left            =   1050
               TabIndex        =   97
               Tag             =   "F"
               Top             =   1260
               Width           =   1250
            End
            Begin VB.TextBox txtflmar 
               Height          =   285
               Left            =   1050
               TabIndex        =   100
               Tag             =   "F"
               Top             =   2295
               Width           =   1250
            End
            Begin VB.TextBox txtflrea 
               Height          =   285
               Left            =   1050
               TabIndex        =   99
               Tag             =   "F"
               Top             =   1950
               Width           =   1250
            End
            Begin VB.TextBox txtfinint2wt 
               Height          =   285
               Left            =   1050
               TabIndex        =   95
               Tag             =   "F"
               Top             =   555
               Width           =   1250
            End
            Begin VB.TextBox txtfinint2cond 
               Height          =   285
               Left            =   1050
               TabIndex        =   96
               Tag             =   "F"
               Top             =   900
               Width           =   1250
            End
            Begin VB.TextBox txtfinint2fat 
               Height          =   285
               Left            =   1050
               TabIndex        =   98
               Tag             =   "F"
               Top             =   1605
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtefinint2 
               Height          =   285
               Left            =   1050
               TabIndex        =   94
               Tag             =   "F"
               Top             =   195
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   45
               TabIndex        =   278
               Tag             =   "F"
               Top             =   1605
               Width           =   960
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   60
               TabIndex        =   277
               Tag             =   "F"
               Top             =   2295
               Width           =   960
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   60
               TabIndex        =   276
               Tag             =   "F"
               Top             =   1950
               Width           =   960
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   240
               TabIndex        =   230
               Tag             =   "F"
               Top             =   195
               Width           =   780
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   120
               TabIndex        =   231
               Tag             =   "F"
               Top             =   555
               Width           =   900
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip Height"
               Height          =   255
               Left            =   30
               TabIndex        =   232
               Tag             =   "F"
               Top             =   915
               Width           =   990
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   255
               Left            =   105
               TabIndex        =   233
               Tag             =   "F"
               Top             =   1245
               Width           =   900
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Recieving"
            Height          =   2640
            Left            =   105
            TabIndex        =   224
            Top             =   510
            Width           =   2550
            Begin VB.TextBox txtrecscore 
               Height          =   285
               Left            =   1065
               TabIndex        =   90
               Tag             =   "F"
               Top             =   1230
               Width           =   1250
            End
            Begin VB.TextBox txtrecmar 
               Height          =   285
               Left            =   1080
               TabIndex        =   93
               Tag             =   "F"
               Top             =   2280
               Width           =   1250
            End
            Begin VB.TextBox txtrecrea 
               Height          =   285
               Left            =   1065
               TabIndex        =   92
               Tag             =   "F"
               Top             =   1935
               Width           =   1250
            End
            Begin VB.TextBox txtfinintfat 
               Height          =   285
               Left            =   1065
               TabIndex        =   91
               Tag             =   "F"
               Top             =   1575
               Width           =   1250
            End
            Begin VB.TextBox txtfinintcond 
               Height          =   285
               Left            =   1065
               TabIndex        =   89
               Tag             =   "F"
               Top             =   900
               Width           =   1250
            End
            Begin VB.TextBox txtfinintwt 
               Height          =   285
               Left            =   1080
               TabIndex        =   88
               Tag             =   "F"
               Top             =   550
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtefinint 
               Height          =   285
               Left            =   1080
               TabIndex        =   87
               Tag             =   "F"
               Top             =   210
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "mm/dd/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "-"
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   45
               TabIndex        =   275
               Tag             =   "F"
               Top             =   1590
               Width           =   960
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Marbling"
               Height          =   255
               Left            =   60
               TabIndex        =   274
               Tag             =   "F"
               Top             =   2280
               Width           =   960
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rib Eye"
               Height          =   255
               Left            =   60
               TabIndex        =   273
               Tag             =   "F"
               Top             =   1935
               Width           =   960
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Score"
               Height          =   255
               Left            =   120
               TabIndex        =   228
               Tag             =   "F"
               Top             =   1245
               Width           =   900
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hip Height"
               Height          =   255
               Left            =   45
               TabIndex        =   227
               Tag             =   "F"
               Top             =   900
               Width           =   1005
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   150
               TabIndex        =   226
               Tag             =   "F"
               Top             =   555
               Width           =   885
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   240
               TabIndex        =   225
               Tag             =   "F"
               Top             =   195
               Width           =   780
            End
         End
         Begin VB.Label Label95 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Feed Lot Notes"
            Height          =   255
            Left            =   3600
            TabIndex        =   243
            Tag             =   "F"
            Top             =   3300
            Width           =   1095
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 3"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   242
            Tag             =   "F"
            Top             =   3960
            Width           =   1500
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 2"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   241
            Tag             =   "F"
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label lblmisc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Misc 1"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   240
            Tag             =   "F"
            Top             =   3240
            Width           =   1500
         End
      End
   End
End
Attribute VB_Name = "frmCalf_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%, dirtyflg0%, dirtyflg1%, dirtyflg2%, dirtyflg3%, dirtyflg4%, dirtyflg5%
Dim oldid$, oldid1$
Dim tbData As Recordset
Dim save%
Dim hipht$, framscor$
Dim d, numofday
Dim j As Double
Dim jj$
Dim listspot As Long
Private Sub calc_205()
  Dim wt205 As Double, Actual_Birth_Weight As Double, Dam_adj As Double
  Dim Adj_Birth_Wt As Double, birthdate, cowage As Long, calfage As Long
 If Val(txtactwt.TEXT) = 0 Then MsgBox "Actual weight must greater then zero to calculate 205 day weight.", vbOKOnly: Exit Sub
 'Adj Birth Weight
 'If NO Actual Birth Weight provided then Adj. Birth Wt. defaults are
 'Sex = 1 Or 3 Adj Birth Wt. = 75
 'Sex = 2 Adj Birth Wt. =70
 'If Actual Birth Weight provided then
 'Adj. Birth Weight=Adjustment+Birth Weight Sex=1,2, or  3
 'Age of Dam=2 then add 8
 '           3 then add 5
 '           4 then add 2
 '           5-10 then add 0
 '           11 and older then add 3
  Dim DB As database
    Dim tbData As Recordset
     Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
      Set tbData = DB.OpenRecordset("COWprof", dbOpenTable)
      tbData.Index = "primarykey"
      tbData.Seek "=", herdid$, txtcowid.TEXT
      If Not tbData.NoMatch Then
          birthdate = Field2Date(tbData!birthdate)
      End If
      tbData.Close: Set tbData = Nothing
      DB.Close: Set DB = Nothing
      If IsDate(birthdate) Then cowage = DateDiff("m", birthdate, Dtewt.TEXT)
      cowage = cowage \ 12
     
  Actual_Birth_Weight = Val(txtbirthwt.TEXT)
  If Actual_Birth_Weight = 0 Then
      If cbosex.TEXT = 2 Then
        Adj_Birth_Wt = 70
      Else
        Adj_Birth_Wt = 75
      End If
  Else
      If cowage = 2 Then Adj_Birth_Wt = Actual_Birth_Weight + 8
      If cowage = 3 Then Adj_Birth_Wt = Actual_Birth_Weight + 5
      If cowage = 4 Then Adj_Birth_Wt = Actual_Birth_Weight + 2
      If cowage >= 5 And cowage <= 10 Then Adj_Birth_Wt = Actual_Birth_Weight + 0
      If cowage >= 11 Then Adj_Birth_Wt = Actual_Birth_Weight + 3
   End If

'Dam Adjustment Factor
'Sex = 2
'Age of Dam=2 then add 54
'           3 then add 36
'           4 then add 18
'           5-10 then add 0
'           11 and older then add 18
'Sex = 1 Or 3
'Age of Dam=2 then add 60
'              3 then add 40
'              4 then add 20
'              5-10 then add 0
'              11 and older then add 20

     If cbosex.TEXT = 2 Then
        If cowage = 2 Then Dam_adj = 54
        If cowage = 3 Then Dam_adj = 36
        If cowage = 4 Then Dam_adj = 18
        If cowage >= 5 And cowage <= 10 Then Dam_adj = 0
        If cowage >= 11 Then Dam_adj = 18
      Else
        If cowage = 2 Then Dam_adj = 60
        If cowage = 3 Then Dam_adj = 40
        If cowage = 4 Then Dam_adj = 20
        If cowage >= 5 And cowage <= 10 Then Dam_adj = 0
        If cowage >= 11 Then Dam_adj = 20
      End If

'This is the formula for calculating 205 day weight.
'Adjusted 205 Wt=((((Actual Wt.-Adj Birth Wt.)/Age in days)*205)+Adj
'Birth Wt.+Age of Dam Adjustment Factor)
           
   calfage = DateDiff("d", dtebirth.TEXT, Dtewt.TEXT)


   wt205 = (((Val(txtactwt.TEXT) - Adj_Birth_Wt) / calfage) * 205 + Adj_Birth_Wt + Dam_adj)

   txtadj205.TEXT = funround2(2, wt205)



End Sub

Private Sub calc_frame(hipht$, framscor$, numofday)
  If Val(cbosex.TEXT) = 2 Then
     j = -11.7086 + (0.4723 * Val(hipht$)) - (0.0239 * numofday) + (0.0000146 * (numofday * numofday)) + (0.0000759 * (Val(hipht$)) * numofday)
  End If
  If Val(cbosex.TEXT) <> 2 Then
     j = -11.548 + (0.4878 * Val(hipht$)) - (0.0289 * numofday) + (0.00001947 * (numofday * numofday)) + (0.0000334 * (Val(hipht$)) * numofday)
  End If
  jj$ = funround2(1, j)
  framscor$ = jj$
End Sub


Private Sub Init_Information()
 Dim t%
 Call init_form(Me) ' Clear Text Boxes
 ' load all combo boxes
 cbosex.AddItem "0"
 cbosex.AddItem "1"
 cbosex.AddItem "2"
 cbosex.AddItem "3"
 cboease.AddItem "0"
 cboease.AddItem "1"
 cboease.AddItem "2"
 cboease.AddItem "3"
 cboease.AddItem "4"
 cboease.AddItem "5"
 For t% = 0 To 9
     cbomancode.AddItem Trim$(Str$(t%))
 Next t%
 cbomancode.AddItem "A"
 cbomancode.AddItem "B"
 cbomancode.AddItem "C"
 cbomancode.AddItem "D"
 cbomancode.AddItem "E"
 cbomancode.AddItem "F"
 cbomancode.AddItem "K"
 cbomancode.AddItem "N"
 cbomancode.AddItem "S"
 cbomancode.AddItem "T"
 cbomancode.AddItem "X"
 cbomancode.AddItem "P"
 Cbograde.AddItem "0"
 Cbograde.AddItem "L1"
 Cbograde.AddItem "L2"
 Cbograde.AddItem "L3"
 Cbograde.AddItem "S1"
 Cbograde.AddItem "S2"
 Cbograde.AddItem "S3"
 Cbograde.AddItem "M1"
 Cbograde.AddItem "M2"
 Cbograde.AddItem "M3"
 Cbograde.AddItem "I"
 With CBOCarcQual
   .AddItem " "
   .AddItem "Prime+"
   .AddItem "Prime"
   .AddItem "Prime-"
   .AddItem "Choice+"
   .AddItem "CAB"
   .AddItem "STS"
   .AddItem "Choice"
   .AddItem "Choice-"
   .AddItem "AAA"
   .AddItem "Select+"
   .AddItem "Select"
   .AddItem "AA"
   .AddItem "Select-"
   .AddItem "Standard+"
   .AddItem "A"
   .AddItem "Standard"
   .AddItem "Standard-"
   .AddItem "B1"
   
   .AddItem "HRI"
   .AddItem "NoRoll"
   .AddItem "B2"
   .AddItem "B3"
   .AddItem "B4"
   .AddItem "D1"
   .AddItem "D2"
   .AddItem "D3"
   .AddItem "D4"
   .AddItem "C"
   .AddItem "Dark"
   .AddItem "Stag"
   .AddItem "Comm"
   .AddItem "Other"
   
 End With
 With CBOConformance
   .AddItem " "
   .AddItem "Yes"
   .AddItem "No"
 End With
 
 
 Dim SQL$
 Dim dbtest As database
 Dim tbCowBrd As Recordset
 Set dbtest = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select EID from eidlist  "
 Set tbCowBrd = dbtest.OpenRecordset(SQL$, dbOpenDynaset)
 Do Until tbCowBrd.EOF
   Cboeid.AddItem tbCowBrd!eid
   tbCowBrd.MoveNext
 Loop
 tbCowBrd.Close: Set tbCowBrd = Nothing
 dbtest.Close: Set dbtest = Nothing
 
 
End Sub

Private Sub Load_information()
 Dim tbwean As Recordset
 Dim tbback As Recordset
 Dim tbrep As Recordset
 Dim tbfeed As Recordset
 Dim tbcarc As Recordset
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("calfbirth", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 If Not tbData.NoMatch Then
     txtid.TEXT = Field2Str(tbData!calfid)
     Txtsireid.TEXT = Field2Str(tbData!sireid)
     txtcowid.TEXT = Field2Str(tbData!CowID)
     TxtCowAge.TEXT = Field2Str(tbData!cowage)
     txtcalfbreed.TEXT = Field2Str(tbData!breed)
     txtregistration.TEXT = Field2Str(tbData!registration)
     txtregname.TEXT = Field2Str(tbData!regname)
     Cboeid.TEXT = Field2Str(tbData!elecid)
     txtbirthwt.TEXT = Field2Str(tbData!birthwt)
     'If IsDate(tbdata!birthdate) Then dtebirth.TEXT = Field2str(tbdata!birthdate)
     dtebirth.TEXT = Field2Date(tbData!birthdate)
     Call set_combo(Me!cbosex, Trim$(Field2Str(tbData!Sex)))
     Call set_combo(Me!cboease, Trim$(Field2Str(tbData!calvingease)))
     txtnotes.TEXT = Field2Str(tbData!notes)
     txtpmisc1.TEXT = Field2Str(tbData!misc1)
     txtpmisc2.TEXT = Field2Str(tbData!misc2)
     txtpmisc3.TEXT = Field2Str(tbData!misc3)
 End If
 Set tbwean = DB.OpenRecordset("calfwean", dbOpenTable)
 tbwean.Index = "primarykey"
 tbwean.Seek "=", herdid$, oldid$
 If Not tbwean.NoMatch Then
     If tbwean!Status = True Then chkwstatus.Value = vbChecked Else chkwstatus.Value = vbUnchecked
     txtactwt = Field2Str(tbwean!actweight)
     Dtewt.TEXT = Field2Date(tbwean!dateweighed)
     Call set_combo(Me!cbomancode, Trim$(Field2Str(tbwean!managecode)))
     txthipht.TEXT = Field2Str(tbwean!chipheight)
     dtemeas.TEXT = Field2Date(tbwean!cdatemeas)
     txtgroup.TEXT = Field2Str(tbwean!group)
     Call set_combo(Me!Cbograde, Trim$(Field2Str(tbwean!grade)))
     txtframe.TEXT = Field2Str(tbwean!score)
     txtadj205.TEXT = Field2Str(tbwean!wt205)
     txtratio.TEXT = Field2Str(tbwean!ratio)
     txtmisc1.TEXT = Field2Str(tbwean!misc1)
     txtmisc2.TEXT = Field2Str(tbwean!misc2)
     txtmisc3.TEXT = Field2Str(tbwean!misc3)
     txtmisc4.TEXT = Field2Str(tbwean!misc4)
     txtmisc5.TEXT = Field2Str(tbwean!misc5)
     txtmisc6.TEXT = Field2Str(tbwean!misc6)
     txtwnotes.TEXT = Field2Str(tbwean!notes)
 End If
 Set tbback = DB.OpenRecordset("calfback", dbOpenTable)
 tbback.Index = "primarykey"
 tbback.Seek "=", herdid$, oldid$
 If Not tbback.NoMatch Then
     If tbback!Status Then Chkbackstatus.Value = vbChecked Else Chkbackstatus.Value = vbUnchecked
     dtebackrec.TEXT = Field2Date(tbback!recdate)
     txtbackrecwt.TEXT = Field2Str(tbback!recweight)
     txtbackhh.TEXT = Field2Str(tbback!recheight)
     txtbackrecframe.TEXT = Field2Str(tbback!recscore)
     dtebackint.TEXT = Field2Date(tbback!intdate)
     txtbackintwt.TEXT = Field2Str(tbback!intweight)
     txtbackinthh.TEXT = Field2Str(tbback!intheight)
     txtbackintframe = Field2Str(tbback!intscore)
     dtebackfinal = Field2Date(tbback!findate)
     txtbackfinwt = Field2Str(tbback!finweight)
     txtbackfinhh = Field2Str(tbback!finheight)
     txtbackfinframe = Field2Str(tbback!finscore)
     txtbnotes.TEXT = Field2Str(tbback!notes)
     txtbackmisc1.TEXT = Field2Str(tbback!misc1)
     txtbackmisc2.TEXT = Field2Str(tbback!misc2)
     txtbackmisc3.TEXT = Field2Str(tbback!misc3)
     txtcontgrp.TEXT = Field2Str(tbback!group)
 End If
 Set tbrep = DB.OpenRecordset("calfrep", dbOpenTable)
 tbrep.Index = "primarykey"
 tbrep.Seek "=", herdid$, oldid$
 If Not tbrep.NoMatch Then
     If tbrep!Status Then chkrstatus.Value = vbChecked Else chkrstatus.Value = vbUnchecked
     txtscrotumcir = Field2Str(tbrep!scrotumcir)
     dtescr = Field2Date(tbrep!scrotumdate)
     txtpelvicsz = Field2Str(tbrep!pelvic)
     dtepelv = Field2Date(tbrep!pelvicdate)
     Dterec.TEXT = Field2Date(tbrep!recdate)
     txtrecwt.TEXT = Field2Str(tbrep!recwt)
     txtreccond.TEXT = Field2Str(tbrep!reccond)
     txtyframe.TEXT = Field2Str(tbrep!recscore)
     txtyhipht.TEXT = Field2Str(tbrep!rechip)
     txtrecfat.TEXT = Field2Str(tbrep!recfat)
     txtrecrib.TEXT = Field2Str(tbrep!recribEYE)
     txtrecmarb.TEXT = Field2Str(tbrep!recmarbling)
     dteint.TEXT = Field2Date(tbrep!intdate)
     txtintwt.TEXT = Field2Str(tbrep!intwt)
     txtintcond.TEXT = Field2Str(tbrep!intcond)
     txtintscore.TEXT = Field2Str(tbrep!intscore)
     txtinthh.TEXT = Field2Str(tbrep!inthip)
     txtintfat.TEXT = Field2Str(tbrep!intfat)
     txtintrib.TEXT = Field2Str(tbrep!intribEYE)
     txtintmarb.TEXT = Field2Str(tbrep!intmarbling)
     dte365.TEXT = Field2Date(tbrep!daydate)
     txt365wt.TEXT = Field2Str(tbrep!daywt)
     txt365cond.TEXT = Field2Str(tbrep!daycond)
     txt365score.TEXT = Field2Str(tbrep!dayscore)
     txt365hh.TEXT = Field2Str(tbrep!dayhip)
     txt365fat.TEXT = Field2Str(tbrep!dayfat)
     txt365rib.TEXT = Field2Str(tbrep!dayribEYE)
     txt365marb.TEXT = Field2Str(tbrep!daymarbling)
     txtynotes.TEXT = Field2Str(tbrep!notes)
     txtymisc1.TEXT = Field2Str(tbrep!misc1)
     txtymisc2.TEXT = Field2Str(tbrep!misc2)
     txtymisc3.TEXT = Field2Str(tbrep!misc3)
     txtweight365.TEXT = Field2Str(tbrep!w365)
 End If
 Set tbfeed = DB.OpenRecordset("calfFeed", dbOpenTable)
 tbfeed.Index = "primarykey"
 tbfeed.Seek "=", herdid$, oldid$
 If Not tbfeed.NoMatch Then
     If tbfeed!Status Then Chkfstatus.Value = vbChecked Else Chkfstatus.Value = vbUnchecked
     dtefinint.TEXT = Field2Date(tbfeed!int1date)
     txtfinintwt.TEXT = Field2Str(tbfeed!int1wt)
     txtfinintcond.TEXT = Field2Str(tbfeed!int1cond)
     txtfinintfat.TEXT = Field2Str(tbfeed!int1fat)
     dtefinint2.TEXT = Field2Date(tbfeed!int2date)
     txtfinint2wt.TEXT = Field2Str(tbfeed!int2wt)
     txtfinint2cond.TEXT = Field2Str(tbfeed!int2cond)
     txtfinint2fat.TEXT = Field2Str(tbfeed!int2fat)
     dtefinfin.TEXT = Field2Date(tbfeed!findate)
     txtfinfinwt.TEXT = Field2Str(tbfeed!finwt)
'     txtfinfincond.TEXT = Field2Str(tbfeed!fincond)
     txtfinfinfat.TEXT = Field2Str(tbfeed!finfat)
     txtfinfinrea.TEXT = Field2Str(tbfeed!finrea)
     txtfinfinmarbl.TEXT = Field2Str(tbfeed!finmarblING)
     txtfnotes.TEXT = Field2Str(tbfeed!notes)
     txtfeedmisc1.TEXT = Field2Str(tbfeed!misc1)
     txtfeedmisc2.TEXT = Field2Str(tbfeed!misc2)
     txtfeedmisc3.TEXT = Field2Str(tbfeed!misc3)
     txtrecscore.TEXT = Field2Str(tbfeed!recscore)
     txtrecrea.TEXT = Field2Str(tbfeed!recrea)
     txtrecmar.TEXT = Field2Str(tbfeed!recmar)
     txtflscore.TEXT = Field2Str(tbfeed!intscore)
     txtflrea.TEXT = Field2Str(tbfeed!intrea)
     txtflmar.TEXT = Field2Str(tbfeed!intmar)
 End If
 Set tbcarc = DB.OpenRecordset("calFcarcass", dbOpenTable)
 tbcarc.Index = "primarykey"
 tbcarc.Seek "=", herdid$, oldid$
 If Not tbcarc.NoMatch Then
     If tbcarc!Status Then chkcstatus.Value = vbChecked Else chkcstatus.Value = vbUnchecked
     txtcarcyield.TEXT = Field2Str(tbcarc!ygrade)
     txtcarccarcwt.TEXT = Field2Str(tbcarc!ywt)
     txtcarcfatthick.TEXT = Field2Str(tbcarc!yfat)
     txtcarckph.TEXT = Field2Str(tbcarc!ykidney)
     txtcarcrib.TEXT = Field2Str(tbcarc!yribeye)
     'txtcarcqual.TEXT = Field2Str(tbcarc!qgrade)
     Call set_combo(CBOCarcQual, Field2Str(tbcarc!qgrade))
     txtqscore.TEXT = Field2Str(tbcarc!qscore)
     txtqcolor.TEXT = Field2Str(tbcarc!qcolor)
     txtqlean.TEXT = Field2Str(tbcarc!qtexture)
     txtqmaturity.TEXT = Field2Str(tbcarc!qmaturity)
     dtecarc.TEXT = Field2Date(tbcarc!carcassdate)
     'txtcarcconf.TEXT = Field2Str(tbcarc!conformance)
     'Call set_combo(CBOConformance, Field2Str(tbcarc!conformance))
     If tbcarc!conformance = True Then Call set_combo(CBOConformance, "Yes") Else Call set_combo(CBOConformance, "No")
     txtcarcmuscle.TEXT = Field2Str(tbcarc!score)
     txtcarcnotes.TEXT = Field2Str(tbcarc!notes)
     txtcarcmisc1.TEXT = Field2Str(tbcarc!misc1)
     txtcarcmisc2.TEXT = Field2Str(tbcarc!misc2)
     txtcarcmisc3.TEXT = Field2Str(tbcarc!misc3)
 End If

 tbData.Close: Set tbData = Nothing
 tbwean.Close: Set tbwean = Nothing
 tbback.Close: Set tbback = Nothing
 tbrep.Close: Set tbrep = Nothing
 tbfeed.Close: Set tbfeed = Nothing
 tbcarc.Close: Set tbcarc = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault

End Sub

Private Sub save_information()
 Dim tbwean As Recordset
 Dim tbback As Recordset
 Dim tbrep As Recordset
 Dim tbfeed As Recordset
 Dim tbcarc As Recordset
 Dim Replace$, TheDate As String
 
  Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("calfBIRTH", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 save% = True
 If Not tbData.NoMatch Then
   If addedflag$ = "D" Then
     tbData.Delete
     Replace$ = ""
     GoTo donedel
    Else
     tbData.Edit
   End If
  Else
   tbData.AddNew
 End If
 With tbData
     !cowage = Val(TxtCowAge.TEXT)
     !herdid = herdid$
     !calfid = txtid.TEXT
     !sireid = Txtsireid.TEXT
     !CowID = txtcowid.TEXT
     !breed = txtcalfbreed.TEXT
     !registration = txtregistration.TEXT
     !regname = txtregname.TEXT
     !elecid = Cboeid.TEXT
     !birthwt = Val(txtbirthwt.TEXT)
     Call Date2Field(!birthdate, dtebirth.TEXT)
     '!birthdate = THEDATE
     !Sex = cbosex.TEXT
     !calvingease = cboease.TEXT
     !notes = txtnotes.TEXT
     !misc1 = txtpmisc1.TEXT
     !misc2 = txtpmisc2.TEXT
     !misc3 = txtpmisc3.TEXT
     .Update
     Replace$ = txtid.TEXT & vbTab & dtebirth
 End With
 Set tbwean = DB.OpenRecordset("calfwean", dbOpenTable)
 tbwean.Index = "primarykey"
 tbwean.Seek "=", herdid$, txtid.TEXT
 'If chkwstatus.Value = vbChecked Then
  If Not tbwean.NoMatch Then
     tbwean.Edit
  Else
   tbwean.AddNew
  End If
  With tbwean
     !herdid = herdid$
     !calfid = txtid.TEXT
     If chkwstatus.Value = vbChecked Then !Status = True Else !Status = False
     !actweight = Val(txtactwt)
     Call Date2Field(!dateweighed, Dtewt.TEXT)
     '!dateweighed = THEDATE
     !managecode = cbomancode.TEXT
     !chipheight = Val(txthipht.TEXT)
     Call Date2Field(!cdatemeas, dtemeas.TEXT)
     '!cdatemeas = THEDATE
     !group = txtgroup.TEXT
     !grade = Cbograde.TEXT
     !score = Val(txtframe.TEXT)
     !wt205 = Val(txtadj205.TEXT)
     !ratio = Val(txtratio.TEXT)
     !misc1 = txtmisc1.TEXT
     !misc2 = txtmisc2.TEXT
     !misc3 = txtmisc3.TEXT
     !misc4 = txtmisc4.TEXT
     !misc5 = txtmisc5.TEXT
     !misc6 = txtmisc6.TEXT
     !notes = txtwnotes.TEXT
     .Update
  End With
 'End If
 Set tbback = DB.OpenRecordset("calfback", dbOpenTable)
 tbback.Index = "primarykey"
 tbback.Seek "=", herdid$, txtid.TEXT
 'If Chkbackstatus.Value = vbChecked Then
  If Not tbback.NoMatch Then
    tbback.Edit
  Else
   tbback.AddNew
  End If
  With tbback
     !herdid = herdid$
     !calfid = txtid.TEXT
     If Chkbackstatus.Value = vbChecked Then tbback!Status = True Else tbback!Status = False
     Call Date2Field(!recdate, dtebackrec.TEXT)
'     !recdate = THEDATE
     !recweight = Val(txtbackrecwt.TEXT)
     !recheight = Val(txtbackhh.TEXT)
     !recscore = Val(txtbackrecframe.TEXT)
     Call Date2Field(!intdate, dtebackint.TEXT)
     '!intdate = THEDATE
     !intweight = Val(txtbackintwt.TEXT)
     !intheight = Val(txtbackinthh.TEXT)
     !intscore = Val(txtbackintframe.TEXT)
     Call Date2Field(!findate, dtebackfinal.TEXT)
     '!findate = Val(dtebackfinal.TEXT)
     !finweight = Val(txtbackfinwt.TEXT)
     !finheight = Val(txtbackfinhh.TEXT)
     !finscore = Val(txtbackfinframe.TEXT)
     !notes = txtbnotes.TEXT
     !misc1 = txtbackmisc1.TEXT
     !misc2 = txtbackmisc2.TEXT
     !misc3 = txtbackmisc3.TEXT
     !group = txtcontgrp.TEXT
     .Update
  End With
 'End If
 Set tbrep = DB.OpenRecordset("calfrep", dbOpenTable)
 tbrep.Index = "primarykey"
 tbrep.Seek "=", herdid$, txtid.TEXT
'If chkrstatus.Value = vbChecked Then
  If Not tbrep.NoMatch Then
    tbrep.Edit
  Else
    tbrep.AddNew
  End If
  With tbrep
     !herdid = herdid$
     !calfid = txtid.TEXT
     If chkrstatus.Value = vbChecked Then tbrep!Status = True Else tbrep!Status = False
     !scrotumcir = Val(txtscrotumcir.TEXT)
     Call Date2Field(!scrotumdate, dtescr.TEXT)
     '!scrotumdate = THEDATE
     !pelvic = Val(txtpelvicsz.TEXT)
     Call Date2Field(!pelvicdate, dtepelv.TEXT)
     '!pelvicdate = Val(dtepelv.TEXT)
     Call Date2Field(!recdate, Dterec.TEXT)
     !recwt = Val(txtrecwt.TEXT)
     !reccond = Val(txtreccond.TEXT)
     !recscore = Val(txtyframe.TEXT)
     !rechip = Val(txtyhipht.TEXT)
     !recfat = Val(txtrecfat.TEXT)
     !recribEYE = Val(txtrecrib.TEXT)
     !recmarbling = Val(txtrecmarb.TEXT)
     Call Date2Field(!intdate, dteint.TEXT)
'     !intdate = THEDATE
     !intwt = Val(txtintwt.TEXT)
     !intcond = Val(txtintcond.TEXT)
     !intscore = Val(txtintscore.TEXT)
     !inthip = Val(txtinthh.TEXT)
     !intfat = Val(txtintfat.TEXT)
     !intribEYE = Val(txtintrib.TEXT)
     !intmarbling = Val(txtintmarb.TEXT)
     Call Date2Field(!daydate, dte365.TEXT)
'     !daydate = THEDATE
     !daywt = Val(txt365wt.TEXT)
     !daycond = Val(txt365cond.TEXT)
     !dayscore = Val(txt365score.TEXT)
     !dayhip = Val(txt365hh.TEXT)
     !dayfat = Val(txt365fat.TEXT)
     !dayribEYE = Val(txt365rib.TEXT)
     !daymarbling = Val(txt365marb.TEXT)
     !notes = txtynotes.TEXT
     !misc1 = txtymisc1.TEXT
     !misc2 = txtymisc2.TEXT
     !misc3 = txtymisc3.TEXT
     !w365 = Val(txtweight365.TEXT)
     .Update
  End With
' End If
 Set tbfeed = DB.OpenRecordset("calFfeed", dbOpenTable)
 tbfeed.Index = "primarykey"
 tbfeed.Seek "=", herdid$, txtid.TEXT
' If Chkfstatus.Value = vbChecked Then
  If Not tbfeed.NoMatch Then
    tbfeed.Edit
  Else
    tbfeed.AddNew
  End If
  With tbfeed
     !herdid = herdid$
     !calfid = txtid.TEXT
     If Chkfstatus.Value = vbChecked Then tbfeed!Status = True Else tbfeed!Status = False
     Call Date2Field(!int1date, dtefinint.TEXT)
     '!int1date = THEDATE
     !int1wt = Val(txtfinintwt.TEXT)
     !int1cond = Val(txtfinintcond.TEXT)
     !int1fat = Val(txtfinintfat.TEXT)
     Call Date2Field(!int2date, dtefinint2.TEXT)
'     !int2date = THEDATE
     !int2wt = Val(txtfinint2wt.TEXT)
     !int2cond = Val(txtfinint2cond.TEXT)
     !int2fat = Val(txtfinint2fat.TEXT)
     Call Date2Field(!findate, dtefinfin.TEXT)
'     !findate = THEDATE
     !finwt = Val(txtfinfinwt.TEXT)
     '!fincond = Val(txtfinfincond.TEXT)
     !finfat = Val(txtfinfinfat.TEXT)
     !finrea = Val(txtfinfinrea.TEXT)
     !finmarblING = Val(txtfinfinmarbl.TEXT)
     !notes = txtfnotes.TEXT
     !misc1 = txtfeedmisc1.TEXT
     !misc2 = txtfeedmisc2.TEXT
     !misc3 = txtfeedmisc3.TEXT
     !recscore = Val(txtrecscore.TEXT)
     !recrea = Val(txtrecrea.TEXT)
     !recmar = Val(txtrecmar.TEXT)
     !intscore = Val(txtflscore.TEXT)
     !intrea = Val(txtflrea.TEXT)
     !intmar = Val(txtflmar.TEXT)
     .Update
  End With
'End If
 Set tbcarc = DB.OpenRecordset("calfcarcass", dbOpenTable)
 tbcarc.Index = "primarykey"
 tbcarc.Seek "=", herdid$, txtid.TEXT
'If chkcstatus.Value = vbChecked Then
  If Not tbcarc.NoMatch Then
    tbcarc.Edit
  Else
    tbcarc.AddNew
  End If
  With tbcarc
     !herdid = herdid$
     !calfid = txtid.TEXT
     If chkcstatus.Value = vbChecked Then tbcarc!Status = True Else tbcarc!Status = False
     !ygrade = Val(txtcarcyield.TEXT)
     !ywt = Val(txtcarccarcwt.TEXT)
     !yfat = Val(txtcarcfatthick.TEXT)
     !ykidney = Val(txtcarckph.TEXT)
     !yribeye = Val(txtcarcrib.TEXT)
     !qgrade = CBOCarcQual.TEXT
     !qscore = txtqscore.TEXT
     !qcolor = txtqcolor.TEXT
     !qtexture = txtqlean.TEXT
     !qmaturity = txtqmaturity.TEXT
     Call Date2Field(!carcassdate, dtecarc.TEXT)
'     !carcassdate = THEDATE
     If CBOConformance <> " " Then
      !conformance = IIf(CBOConformance = "Yes", True, False)
     Else
      !conformance = Null
     End If
     !score = Val(txtcarcmuscle.TEXT)
     !notes = txtcarcnotes.TEXT
     !misc1 = txtcarcmisc1.TEXT
     !misc2 = txtcarcmisc2.TEXT
     !misc3 = txtcarcmisc3.TEXT
   .Update
  End With
'End If
 tbwean.Close: Set tbwean = Nothing
 tbback.Close: Set tbback = Nothing
 tbrep.Close: Set tbrep = Nothing
 tbfeed.Close: Set tbfeed = Nothing
 tbcarc.Close: Set tbcarc = Nothing
donedel:
tbData.Close: Set tbData = Nothing
DB.Close: Set DB = Nothing
'Call Update_mh_ListBoxes("lstcalf", 0, oldid$, Replace$)
Call UpdateListProListBoxes("lstcalf", 0, oldid$, Replace$)
Screen.MousePointer = vbDefault: 'after this line update list boxes
dirtyflag% = False
dirtyflg0% = False
dirtyflg1% = False
dirtyflg2% = False
dirtyflg3% = False
dirtyflg4% = False
dirtyflg5% = False
End Sub


Private Sub valid_form(exitcode%)
  Dim sDate$
  Dim responce%
  exitcode% = 0
  If herdid = "" Then
        Beep
        MsgBox "Please Select A Herd", vbOKOnly + vbCritical, Me.Caption
        selherd_List.cmdcancel.Visible = False
        selherd_List.Show vbModal
    End If
  'calf birth textbox
    If txtid.TEXT = "" Then
        Beep
        MsgBox "Calf ID Must Be Filled Out", vbOKOnly
        txtid.SetFocus
        exitcode% = 1
        Exit Sub
    End If
    If UCase$(oldid$) <> UCase$(txtid.TEXT) Then
        Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
        Set tbData = DB.OpenRecordset("calfBIRTH", dbOpenTable)
        tbData.Index = "primarykey"
        tbData.Seek "=", herdid$, txtid.TEXT
        If Not tbData.NoMatch Then
            Beep
            MsgBox "Calf ID Can Not Be Duplicated", vbOKOnly
            exitcode% = 1
            tbData.Close: Set tbData = Nothing
            DB.Close: Set DB = Nothing
            txtid.SetFocus
            Exit Sub
        End If
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
    End If
    
    If Cboeid.TEXT <> "" Then
      If Len(Cboeid) <> 15 Then
            MsgBox "EID must be 15 characters.", vbOKOnly
            exitcode% = 1
            Cboeid.SetFocus
            Exit Sub
      End If
    End If
    
    
    'valid cow id
    
    Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
    Set tbData = DB.OpenRecordset("cowPROF", dbOpenTable)
    tbData.Index = "primarykey"
    tbData.Seek "=", herdid$, txtcowid.TEXT
    If tbData.NoMatch Then
       Beep
       MsgBox "Must have a valid Cow ID", vbOKOnly
       SSTab1.Tab = 0
       txtcowid.SetFocus
       exitcode% = 1
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
       txtcowid.SetFocus
       Exit Sub
    End If
    tbData.Close: Set tbData = Nothing
    DB.Close: Set DB = Nothing
    If TxtCowAge.TEXT = "" Then
        Beep
        MsgBox "Must have a valid cow age", vbOKOnly
        SSTab1.Tab = 0
        TxtCowAge.SetFocus
        exitcode% = 1
        Exit Sub
    End If
    
    
    'valid sire
    Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
    Set tbData = DB.OpenRecordset("sirePROF", dbOpenTable)
    tbData.Index = "primarykey"
    tbData.Seek "=", herdid$, Txtsireid.TEXT
    If tbData.NoMatch Then
       Beep
       MsgBox "Must have a valid Sire ID", vbOKOnly
       exitcode% = 1
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
       Txtsireid.SetFocus
       Exit Sub
    End If
    tbData.Close: Set tbData = Nothing
    DB.Close: Set DB = Nothing
    'valid sex
    If cbosex.TEXT = "" Then
        Beep
        MsgBox "Must have a valid sex", vbOKOnly
        SSTab1.Tab = 0
        cbosex.SetFocus
        exitcode% = 1
        Exit Sub
    End If
    'valid birth date
    'If cbomancode.TEXT <> "A" And cbomancode.TEXT <> "X" Then
       If Not IsDate(dtebirth.TEXT) Then
          Beep
          MsgBox "Must have a valid birth date if this cow is open please use the birthdate of the last calf born for this calving season.", vbOKOnly
          SSTab1.Tab = 0
          dtebirth.SetFocus
          exitcode% = 1
          Exit Sub
       End If
    'End If
    sDate = Right(dtebirth.TEXT, 4)
    
    
    'sh 9/15/99 as per kris ringwall
    'If cboease.TEXT = "" Then
    '    Beep
    '    MsgBox "Must have a valid calving ease", vbOKOnly
    '    SSTab1.Tab = 0
    '    cboease.SetFocus
    '    exitcode% = 1
    '    Exit Sub
    'End If
    
' VALIDITY CHECK FOR SECOND HALF OF CALF TAB
    If chkwstatus.Value = vbChecked Then
      If Val(txtactwt.TEXT) > 0 And Dtewt.TEXT = "--/--/----" Then 'if actual wt > 0 then must have a weigh date
         Beep
         MsgBox "Must have a valid Date Weighed", vbOKOnly
         SSTab1.Tab = 1
         Dtewt.SetFocus
         exitcode% = 1
         Exit Sub
      End If
      If Val(txtactwt.TEXT) = 0 Then 'if actual wt = 0 then must a have reason
         If cbomancode.TEXT <> "A" And cbomancode.TEXT <> "B" And cbomancode.TEXT <> "C" And cbomancode.TEXT <> "D" And cbomancode.TEXT <> "X" Then
               Beep
               MsgBox "Management code must be A, B, C, D or X if there is no or actual weight = 0", vbOKOnly
               SSTab1.Tab = 1
               cbomancode.SetFocus
               exitcode% = 1
               Exit Sub
         End If
      Else
         Select Case cbomancode.TEXT
            Case "A", "B", "C", "D", "X"
               Beep
               MsgBox "Management code can't be A, B, C, D or X if there is an actual weight > 0", vbOKOnly
               SSTab1.Tab = 1
               cbomancode.SetFocus
               exitcode% = 1
               Exit Sub
         End Select
      End If
      If dtemeas.TEXT <> "--/--/----" And Val(txthipht.TEXT) = 0 Then 'must have hip height if there is date measured
         Beep
         MsgBox "Must have a valid hip height if there is a date measured", vbOKOnly
         SSTab1.Tab = 1
         txthipht.SetFocus
         exitcode% = 1
         Exit Sub
      End If
   End If
    
If Chkbackstatus.Value = vbChecked Then
      'hip height and date validation
      If Val(txtbackhh) > 0 And dtebackrec = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackrec.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackrec.SetFocus
            dtebackrec.SelStart = 0
            dtebackrec.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtbackinthh.TEXT) > 0 And dtebackint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackint.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackint.SetFocus
            dtebackint.SelStart = 0
            dtebackint.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtbackfinhh.TEXT) > 0 And dtebackfinal.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have final date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackfinal.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackfinal.SetFocus
            dtebackfinal.SelStart = 0
            dtebackfinal.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      'weight and date validation
      If Val(txtbackrecwt) > 0 And dtebackrec = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackrec.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackrec.SetFocus
            dtebackrec.SelStart = 0
            dtebackrec.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtbackintwt.TEXT) > 0 And dtebackint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackint.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackint.SetFocus
            dtebackint.SelStart = 0
            dtebackint.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtbackfinwt.TEXT) > 0 And dtebackfinal.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have final date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 2
            dtebackfinal.TEXT = Format(Now, "mm/dd/yyyy")
            dtebackfinal.SetFocus
            dtebackfinal.SelStart = 0
            dtebackfinal.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If

End If
  
If chkrstatus.Value = vbChecked Then
   'hip height validation
   If Val(txtyhipht) > 0 And Dterec.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 3
            Dterec.TEXT = Format(Now, "mm/dd/yyyy")
            Dterec.SetFocus
            Dterec.SelStart = 0
            Dterec.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtinthh.TEXT) > 0 And dteint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 3
            dteint.TEXT = Format(Now, "mm/dd/yyyy")
            dteint.SetFocus
            dteint.SelStart = 0
            dteint.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txt365hh.TEXT) > 0 And dte365.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have final date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 3
            dte365.TEXT = Format(Now, "mm/dd/yyyy")
            dte365.SetFocus
            dte365.SelStart = 0
            dte365.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      'weight and date validation
         If Val(txtrecwt) > 0 And Dterec.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 3
            Dterec.TEXT = Format(Now, "mm/dd/yyyy")
            Dterec.SetFocus
            Dterec.SelStart = 0
            Dterec.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txtintwt.TEXT) > 0 And dteint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if hip weight is filled out", vbOKOnly
            SSTab1.Tab = 3
            dteint.TEXT = Format(Now, "mm/dd/yyyy")
            dteint.SetFocus
            dteint.SelStart = 0
            dteint.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
      
      If Val(txt365wt.TEXT) > 0 And dte365.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have final date filled out if hip weight is filled out", vbOKOnly
            SSTab1.Tab = 3
            dte365.TEXT = Format(Now, "mm/dd/yyyy")
            dte365.SetFocus
            dte365.SelStart = 0
            dte365.SelLength = 10
            exitcode% = 1
            Exit Sub
      End If
End If

If Chkfstatus.Value = vbChecked Then
   'hip height validation
   If Val(txtfinintcond) > 0 And dtefinint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 4
            dtefinint.TEXT = Format(Now, "mm/dd/yyyy")
            dtefinint.SetFocus
            dtefinint.SelStart = 0
            dtefinint.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If
      
   If Val(txtfinint2cond.TEXT) > 0 And dtefinint2.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if hip height is filled out", vbOKOnly
            SSTab1.Tab = 4
            dtefinint2.TEXT = Format(Now, "mm/dd/yyyy")
            dtefinint2.SetFocus
            dtefinint2.SelStart = 0
            dtefinint2.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If
      'weight and date validation
      If Val(txtfinintwt) > 0 And dtefinint.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have receiving date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 4
            dtefinint.TEXT = Format(Now, "mm/dd/yyyy")
            dtefinint.SetFocus
            dtefinint.SelStart = 0
            dtefinint.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If
      
   If Val(txtfinint2wt.TEXT) > 0 And dtefinint2.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have interim date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 4
            dtefinint2.TEXT = Format(Now, "mm/dd/yyyy")
            dtefinint2.SetFocus
            dtefinint2.SelStart = 0
            dtefinint2.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If

   If Val(txtfinfinwt) > 0 And dtefinfin.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have final date filled out if weight is filled out", vbOKOnly
            SSTab1.Tab = 4
            dtefinfin.TEXT = Format(Now, "mm/dd/yyyy")
            dtefinfin.SetFocus
            dtefinfin.SelStart = 0
            dtefinfin.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If
End If

If chkcstatus.Value = vbChecked Then
   If dtecarc.TEXT = "--/--/----" Then
            Beep
            MsgBox "Must have carcass date", vbOKOnly
            SSTab1.Tab = 5
            dtecarc.TEXT = Format(Now, "mm/dd/yyyy")
            dtecarc.SetFocus
            dtecarc.SelStart = 0
            dtecarc.SelLength = 10
            exitcode% = 1
            Exit Sub
   End If
End If

If cbosex.TEXT = 0 And Val(txtactwt.TEXT) <> 0 Then
  MsgBox "Zero Calf Sex and An Actual Weight, all records entered with an actual weight must be entered with an actual calf sex.  If the calf sex is undeterminable enter 3 (steer) for default calf sex.", vbOKOnly
  exitcode% = 1
  Exit Sub
End If

If cbosex.TEXT = 0 And cbomancode.TEXT = "X" Then
  MsgBox "Zero Calf Sex and An X Management Code, all records entered with ""X"" management code must be entered with an actual calf sex.  If the calf sex is undeterminable enter 3 (steer) for default calf sex.", vbOKOnly
  exitcode% = 1
  Exit Sub
End If

If Check_EID(Cboeid.TEXT, "Calf", oldid, herdid, "", "Calf") = False Then
   Beep
   MsgBox "EID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
   exitcode% = 1
End If

'If dirtyflg1% Then
'     responce% = MsgBox("Data on the Weaning tab was changed do you want to save?", vbYesNo + vbCritical, Me.Caption)
'     If responce% = vbYes Then chkwstatus.Value = vbChecked
'  End If
'  If dirtyflg2% Then
'     responce% = MsgBox("Data on the Background tab was changed do you want to save?", vbYesNo + vbCritical, Me.Caption)
'     If responce% = vbYes Then Chkbackstatus.Value = vbChecked
'  End If
'  If dirtyflg3% Then
'     responce% = MsgBox("Data on the Replacement tab was changed do you want to save?", vbYesNo + vbCritical, Me.Caption)
'     If responce% = vbYes Then chkrstatus.Value = vbChecked
'  End If
'  If dirtyflg4% Then
'     responce% = MsgBox("Data on the Feed Lot tab was changed do you want to save?", vbYesNo + vbCritical, Me.Caption)
'     If responce% = vbYes Then Chkfstatus.Value = vbChecked
'  End If
'  If dirtyflg5% Then
'     responce% = MsgBox("Data on the Carcass tab was changed do you want to save?", vbYesNo + vbCritical, Me.Caption)
'     If responce% = vbYes Then chkcstatus.Value = vbChecked
'  End If
End Sub

Private Sub Chkbackstatus_Click()
Dim CTL As Control
If Chkbackstatus.Value = vbUnchecked Then
   For Each CTL In Me.Controls
      If CTL.Tag = "Back" Then CTL.Enabled = False
   Next
Else
   For Each CTL In Me.Controls
      If CTL.Tag = "Back" Then CTL.Enabled = True
   Next
End If

End Sub

Private Sub Chkbackstatus_GotFocus()
  If SSTab1.Tab = 1 Then
    Cmdsave.SetFocus
  End If
End Sub


Private Sub chkcstatus_Click()
Dim CTL As Control
If chkcstatus.Value = vbUnchecked Then
   For Each CTL In Me.Controls
      If CTL.Tag = "Carc" Then CTL.Enabled = False
   Next
Else
   For Each CTL In Me.Controls
      If CTL.Tag = "Carc" Then CTL.Enabled = True
   Next
End If '
End Sub

Private Sub chkcstatus_GotFocus()
 If SSTab1.Tab = 4 Then
    Cmdsave.SetFocus
  End If

End Sub


Private Sub chkedit_Click()
   If chkedit.Value = vbChecked Then
     txtframe.Enabled = True
     txtadj205.Enabled = True
     txtratio.Enabled = True
     lblratio.Enabled = True
     lbl205.Enabled = True
     lblscore.Enabled = True
   Else
     txtframe.Enabled = False
     txtadj205.Enabled = False
     txtratio.Enabled = False
     lblratio.Enabled = False
     lbl205.Enabled = False
     lblscore.Enabled = False
   End If
End Sub

Private Sub Chkfstatus_Click()
Dim CTL As Control
If Chkfstatus.Value = vbUnchecked Then
   For Each CTL In Me.Controls
      If CTL.Tag = "F" Then CTL.Enabled = False
   Next
Else
   For Each CTL In Me.Controls
      If CTL.Tag = "F" Then CTL.Enabled = True
   Next
End If '
End Sub

Private Sub Chkfstatus_GotFocus()
  If SSTab1.Tab = 3 Then
    Cmdsave.SetFocus
  End If
End Sub


Private Sub chkrstatus_Click()
Dim CTL As Control
If chkrstatus.Value = vbUnchecked Then
   For Each CTL In Me.Controls
      If CTL.Tag = "REPL" Then CTL.Enabled = False
   Next
Else
   For Each CTL In Me.Controls
      If CTL.Tag = "REPL" Then CTL.Enabled = True
   Next
End If '
End Sub

Private Sub chkrstatus_GotFocus()
 If SSTab1.Tab = 2 Then
    Cmdsave.SetFocus
  End If
End Sub


Private Sub chkwstatus_Click()
Dim CTL As Control
If chkwstatus.Value = vbUnchecked Then
   For Each CTL In Me.Controls
      If CTL.Tag = "Weaning" Then CTL.Enabled = False
   Next
Else
   For Each CTL In Me.Controls
      If CTL.Tag = "Weaning" Then CTL.Enabled = True
   Next
End If
End Sub

Private Sub chkwstatus_GotFocus()
SSTab1.Tab = 1
End Sub


Private Sub cmdapply_Click()
 Call calc_205
 Dim numofday
 If IsDate(dtemeas.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtemeas.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txthipht.TEXT, framscor$, numofday)
    If framscor$ > 0 Then txtframe.TEXT = framscor$
 End If
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub



Private Sub cmddefault_Click()
  If SSTab1.Tab = 2 Then  'Background
    'get from weaning
    dtebackrec.TEXT = Dtewt.TEXT
    txtbackrecwt.TEXT = txtactwt.TEXT
    txtbackhh.TEXT = txthipht.TEXT
    txtbackrecframe.TEXT = txtframe.TEXT
  End If
  If SSTab1.Tab = 3 Then  'Replacement
    If Chkbackstatus.Value = vbChecked Then
      'background final
      Dterec.TEXT = dtebackfinal.TEXT
      txtrecwt.TEXT = txtbackfinwt.TEXT
      txtyhipht.TEXT = txtbackfinhh.TEXT
      txtyframe.TEXT = txtbackfinframe.TEXT
    Else
      'get from weaning
      Dterec.TEXT = Dtewt.TEXT
      txtrecwt.TEXT = txtactwt.TEXT
      txtyhipht.TEXT = txthipht.TEXT
      txtyframe.TEXT = txtframe.TEXT
    End If
  End If
  If SSTab1.Tab = 4 Then  'feedlot
   If Chkbackstatus.Value = vbChecked Then
      'background final
      dtefinint.TEXT = dtebackfinal.TEXT
      txtfinintwt.TEXT = txtbackfinwt.TEXT
      txtfinintcond.TEXT = txtbackfinhh.TEXT
      txtrecscore.TEXT = txtbackfinframe.TEXT
    Else
      'get from weaning
      dtefinint.TEXT = Dtewt.TEXT
      txtfinintwt.TEXT = txtactwt.TEXT
      txtfinintcond.TEXT = txthipht.TEXT
      txtrecscore.TEXT = txtframe.TEXT
    End If
  End If
      
  
  
End Sub

Private Sub cmdnext_Click()
 Dim RESPONSE%, exitcode%
 If dirtyflag% Then
   Beep
   RESPONSE% = MsgBox("Information Has Been Changed" & vbCrLf & " Do You Wish To Save?", vbYesNoCancel + vbQuestion, Me.Caption)
   Select Case RESPONSE%
    Case vbYes
     Call valid_form(exitcode%)
     If exitcode% <> 0 Then
       Exit Sub
     End If
     Call save_information
    Case vbCancel
   End Select
 End If
 Screen.MousePointer = vbHourglass

 'If frmcalf_list.lstcalf.ListCount > 0 Then
 'frmcalf_list.lstcalf.Col = 0
'if frmcalf_list.lstcalf.ListIndex = 0 Then
'   frmcalf_list.lstcalf.ListIndex = frmcalf_list.lstcalf.ListIndex + 1
'Else
'   If frmcalf_list.lstcalf.ListIndex < frmcalf_list.lstcalf.ListCount - 1 Then
'      frmcalf_list.lstcalf.ListIndex = frmcalf_list.lstcalf.ListIndex + 1
'   Else
'      frmcalf_list.lstcalf.ListIndex = 0
'   End If
'End If
' frmCalf_Data.Tag = "E/" & frmcalf_list.lstcalf.ColList(frmcalf_list.lstcalf.ListIndex)
'End If

  If listspot >= frmcalf_list.lstcalf.ListCount - 1 Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  listspot = listspot + 1
  frmcalf_list.lstcalf.Col = 0
  frmcalf_list.lstcalf.Row = listspot
  Me.Tag = "E/" & frmcalf_list.lstcalf.ColList

 Screen.MousePointer = vbDefault
 Call Form_Activate
End Sub
Private Sub cmdprev_Click()
 Dim RESPONSE%, exitcode%
 If dirtyflag% Then
   Beep
   RESPONSE% = MsgBox("Information Has Been Changed" & vbCrLf & " Do You Wish To Save?", vbYesNoCancel + vbQuestion, Me.Caption)
   Select Case RESPONSE%
    Case vbYes
     Call valid_form(exitcode%)
     If exitcode% <> 0 Then
       Exit Sub
     End If
     Call save_information
    Case vbCancel
   End Select
 End If
 Screen.MousePointer = vbHourglass

'If frmcalf_list.lstcalf.ListCount > 0 Then
' frmcalf_list.lstcalf.Col = 0
'If frmcalf_list.lstcalf.ListIndex = 0 Then
'   frmcalf_list.lstcalf.ListIndex = frmcalf_list.lstcalf.ListCount - 1
'Else
'   If frmcalf_list.lstcalf.ListIndex <= frmcalf_list.lstcalf.ListCount - 1 Then
'      frmcalf_list.lstcalf.ListIndex = frmcalf_list.lstcalf.ListIndex - 1
'   Else
'      frmcalf_list.lstcalf.ListIndex = 0
'   End If
'End If
' frmCalf_Data.Tag = "E/" & frmcalf_list.lstcalf.ColList(frmcalf_list.lstcalf.ListIndex)
'End If
  If listspot <= 0 Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  listspot = listspot - 1
  frmcalf_list.lstcalf.Col = 0
  frmcalf_list.lstcalf.Row = listspot
  Me.Tag = "E/" & frmcalf_list.lstcalf.ColList
 Screen.MousePointer = vbDefault
 Call Form_Activate
End Sub


Public Property Let ListBoxSelection(ByVal vNewValue As Variant)
  listspot = Val(vNewValue)
  If listspot = -1 Then listspot = 0
End Property
Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%, SQL$, theyear$
 Dim TableName$(100)
 Dim wdate As String, hhdate As String
 
If gIsDemo Then
   If IsValidCalfEntry = False Then MsgBox "The demo version of this software only allows twenty calf records to be entered.", vbOKOnly, "C.H.A.P.S. Demo": Exit Sub
End If
 
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
 If addedflag$ = "D" Then
   Call CheckID(dbfile$, "calf", oldid$, TableName$())
   RESPONSE% = vbYes
   If Val(TableName$(0)) > 0 Then
     Beep
     RESPONSE% = MsgBox("Warning This calf Id Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
   End If
   If RESPONSE% = vbYes Then
     Call save_information
   Else
     Exit Sub
   End If
 End If
 If addedflag$ <> "D" Then
    Call save_information
 End If
 If addedflag$ = "D" Then Unload Me
 If addedflag$ = "A" Then
   Me.Tag = "A"
   theyear$ = Right(dtebirth.TEXT, 4)
   wdate = Dtewt.TEXT
   hhdate = dtemeas.TEXT
   Call Form_Activate
   txtid.SetFocus
   dtebirth.TEXT = "--/--/" & theyear$
   Dtewt.TEXT = wdate
   dtemeas.TEXT = hhdate
  Else
   Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
   SQL$ = "select * from calfbirth where herdid='" & herdid$ & "'"
   If IsDate(frmcalf_list.Dtestart.TEXT) And IsDate(frmcalf_list.dteend.TEXT) Then SQL$ = SQL$ & " and birthdate between #" & frmcalf_list.Dtestart.TEXT & "# and #" & frmcalf_list.dteend.TEXT & "#"
   Set tbData = DB.OpenRecordset(SQL$, dbOpenDynaset)
   tbData.FindFirst "herdid = '" & herdid$ & "' And calfid = '" & txtid.TEXT & "'"
   tbData.MoveNext
   If tbData.EOF Then
     tbData.MoveFirst
     frmCalf_Data.Tag = "E/" & tbData!calfid
   Else
    'tbdata.MoveNext
    frmCalf_Data.Tag = "E/" & tbData!calfid
   End If
   tbData.Close: Set tbData = Nothing
   DB.Close: Set DB = Nothing
   Screen.MousePointer = vbDefault
   Call Form_Activate
 End If

End Sub

Private Sub Command1_Click()
 Dim numofday
 If IsDate(dtebackrec.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtebackrec.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtbackhh.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtbackrecframe.TEXT = framscor$
 End If
 
 If IsDate(dtebackint.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtebackint.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtbackinthh.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtbackintframe.TEXT = framscor$
 End If
 
 If IsDate(dtebackfinal.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtebackfinal.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtbackfinhh.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtbackfinframe.TEXT = framscor$
 End If
End Sub

Private Sub Command2_Click()
 Dim numofday, WT365#
 If IsDate(Dterec.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, Dterec.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtyhipht.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtyframe.TEXT = framscor$
 End If
 
 If IsDate(dteint.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dteint.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtinthh.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtintscore.TEXT = framscor$
 End If
 
 If IsDate(dte365.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dte365.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txt365hh.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txt365score.TEXT = framscor$
 End If
 
 If IsDate(dte365.TEXT) And IsDate(Dtewt.TEXT) Then
    numofday = DateDiff("d", Dtewt, dte365.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call Calc_365DWt(WT365)
    If WT365 > 0 Then txtweight365 = funround2(2, WT365)
 End If
End Sub

Private Sub Calc_365DWt(WT365#)
Dim NumDays#
NumDays = DateDiff("D", Dtewt, dte365)
WT365 = (((Val(txt365wt) - Val(txtactwt)) / NumDays) * 160 + Val(txtadj205))
End Sub

Private Sub Command3_Click()
txtcarcyield = Calc_Yield_Grade(Val(txtcarccarcwt), Val(txtcarcfatthick), Val(txtcarckph), Val(txtcarcrib))
End Sub

Private Sub Command4_Click()
Dim numofday
 If IsDate(dtefinint.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtefinint.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtfinintcond.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtrecscore.TEXT = framscor$
 End If
 
 If IsDate(dtefinint2.TEXT) And IsDate(dtebirth.TEXT) Then
    numofday = DateDiff("d", dtebirth.TEXT, dtefinint2.TEXT)
    If numofday < 0 Then
       Beep
       MsgBox ("Invalid Date Measured")
       Exit Sub
    End If
    Call calc_frame(txtfinint2cond.TEXT, framscor$, numofday)
    If Val(framscor$) > 0 Then txtflscore.TEXT = framscor$
 End If
 
 'If IsDate(dtebackfinal.TEXT) And IsDate(dtebirth.TEXT) Then
 '   numofday = DateDiff("d", dtebirth.TEXT, dtebackfinal.TEXT)
 '   If numofday < 0 Then
 '      Beep
 '      MsgBox ("Invalid Date Measured")
 '      Exit Sub
 '   End If
 '   Call calc_frame(txtbackfinhh.TEXT, framscor$, numofday)
 '   txtbackfinframe.TEXT = framscor$
 'End If
End Sub

Private Sub Form_Activate()
 If Me.Tag = "" Then Exit Sub
 addedflag$ = Left$(Me.Tag, 1)
 Me.Caption = Me.Caption & " for Herd " & herdid$
 Me.Caption = "Calf Information" & " for Herd " & herdid$
 Screen.MousePointer = vbHourglass
 Call Init_Information
 If addedflag$ = "A" Then
    cmdnext.Enabled = False
    CMDprev.Enabled = False
   'Me.caption = "Add"
    oldid$ = ""
    cboease.ListIndex = 1
    cbomancode.ListIndex = 0
    Cbograde.ListIndex = 0
 End If
 If addedflag$ = "E" Or addedflag$ = "D" Then
   oldid$ = Trim$(Mid$(Me.Tag, 3))
   'Me.caption = "Edit"
   Me.Caption = "Calf Information" & " for Herd " & herdid$ & " - Calf " & oldid$
   If addedflag$ = "D" Then
     'Me.caption = "Delete"
     Call disable_controls(Me)
     Cmdsave.Caption = "&Delete"
     Cmdsave.Enabled = True
     cmdcancel.Enabled = True
     SSTab1.Enabled = True
   End If
   Call Load_information
 End If
 SSTab1.Tab = 0
 Me.Tag = ""
 Screen.MousePointer = vbDefault
 Me.Enabled = True
 dirtyflag% = False
 dirtyflg0% = False
 dirtyflg1% = False
 dirtyflg2% = False
 dirtyflg3% = False
 dirtyflg4% = False
 dirtyflg5% = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
 If SSTab1.Tab = 0 Then dirtyflg0% = True
 If SSTab1.Tab = 1 Then dirtyflg1% = True
 If SSTab1.Tab = 2 Then dirtyflg2% = True
 If SSTab1.Tab = 3 Then dirtyflg3% = True
 If SSTab1.Tab = 4 Then dirtyflg4% = True
 If SSTab1.Tab = 5 Then dirtyflg5% = True

End Sub


Private Sub Form_Load()
  Dim i%
  Call centermdiform(Me, mdimain, 0, 0)
  'dtebirth.Format = "dd/MM/yyyy"
  For i% = 3 To 20
    lblmisc(i%).Caption = calfhead(i% - 2)
  Next i%
  For i% = 0 To 2
    lblmisc(i%).Caption = calfhead(i% + 1)
  Next i%
Call chkwstatus_Click
Call Chkbackstatus_Click
Call chkrstatus_Click
Call Chkfstatus_Click
Call chkcstatus_Click


Call AddCustomToolTip(cbosex, "1=Bull" & vbCrLf & "2=Heifer" & vbCrLf & "3=Steer", Me)

Call AddCustomToolTip(cbomancode, "A=Cow did not calve but retained in herd" & vbCrLf & "B=Cow aborted" & vbCrLf & "C=Calf died before 2 weeks of age" & vbCrLf & "D=Calf died after 2 weeks of age but before weaning" & vbCrLf & "E=Embryo" & vbCrLf & "F=Foster Calf" & vbCrLf & "K=Foster (twin) Calf" & vbCrLf & "N=Calf raised by onother cow" & vbCrLf & "T=Twin raised as twin" & vbCrLf & "S=Twin raised on own dam" & vbCrLf & "X=Incomplete record", Me)



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
       Exit Sub
     End If
     Call save_information
    Case vbCancel
     Cancel = True
   End Select
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Cancel Then Exit Sub
 Set frmCalf_Data = Nothing
End Sub











Private Sub SSTab1_GotFocus()
  If addedflag$ = "D" Then
    SSTab1.Tab = 0
    Exit Sub
  End If
  If SSTab1.Tab = 0 Then txtid.SetFocus
  If SSTab1.Tab = 1 Then
     If txtactwt.Enabled = True Then txtactwt.SetFocus
  End If
  If SSTab1.Tab = 2 Then
     If dtebackrec.Enabled = True Then dtebackrec.SetFocus
  End If
  If SSTab1.Tab = 3 Then
   If txtscrotumcir.Enabled = True Then txtscrotumcir.SetFocus
  End If
  If SSTab1.Tab = 4 Then
     If dtefinint.Enabled = True Then dtefinint.SetFocus
  End If
  If SSTab1.Tab = 5 Then
     If txtcarcyield.Enabled = True Then txtcarcyield.SetFocus
  End If
  
  Me.Caption = "Calf Information" & " for Herd " & herdid$ & " - Calf " & txtid.TEXT

End Sub



Private Sub txt365cond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365cond_LostFocus()
If Val(txt365cond.TEXT) < 0 Or Val(txt365cond.TEXT) >= 10 Then
   MsgBox "Please enter a number between 0 and 9", vbOKOnly
   SSTab1.Tab = 3
   txt365cond.SetFocus
   txt365cond.SelStart = 0
   txt365cond.SelLength = Len(txt365cond.TEXT)
End If
End Sub


Private Sub txt365fat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365fat_LostFocus()
If Val(txt365fat.TEXT) < 0 Or Val(txt365fat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 3
   txt365fat.SetFocus
   txt365fat.SelStart = 0
   txt365fat.SelLength = Len(txt365fat.TEXT)
End If
End Sub


Private Sub txt365hh_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365hh_LostFocus()
If Val(txt365hh.TEXT) < 0 Or Val(txt365hh.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 3
   txt365hh.SetFocus
   txt365hh.SelStart = 0
   txt365hh.SelLength = Len(txt365hh.TEXT)
End If
End Sub


Private Sub txt365marb_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 9 Then Cmdsave.SetFocus
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365marb_LostFocus()
If Val(txt365marb.TEXT) < 0 Or Val(txt365marb.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 3
   txt365marb.SetFocus
   txt365marb.SelStart = 0
   txt365marb.SelLength = Len(txt365marb.TEXT)
End If
End Sub


Private Sub txt365rib_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365rib_LostFocus()
If Val(txt365rib.TEXT) < 0 Or Val(txt365rib.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 3
   txt365rib.SetFocus
   txt365rib.SelStart = 0
   txt365rib.SelLength = Len(txt365rib.TEXT)
End If
End Sub


Private Sub txt365score_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365wt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txt365wt_LostFocus()
If Val(txt365wt.TEXT) < 0 Or Val(txt365wt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 3
   txt365wt.SetFocus
   txt365wt.SelStart = 0
   txt365wt.SelLength = Len(txt365wt.TEXT)
End If
End Sub


Private Sub txtactwt_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 8 And txtactwt.SelStart = 0 Then
'     SSTab1.Tab = 0
'     txtid.SetFocus
'   End If
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtactwt_LostFocus()
If Val(txtactwt.TEXT) < 0 Or Val(txtactwt.TEXT) > 1000 Then
   MsgBox "Please enter a number between 0 and 1000", vbOKOnly
   SSTab1.Tab = 1
   txtactwt.SetFocus
   txtactwt.SelStart = 0
   txtactwt.SelLength = Len(txtactwt.TEXT)
End If
End Sub


Private Sub txtadj205_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackfinframe_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackfinhh_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackfinhh_LostFocus()
If Val(txtbackfinhh.TEXT) < 0 Or Val(txtbackfinhh.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 2
   txtbackfinhh.SetFocus
   txtbackfinhh.SelStart = 0
   txtbackfinhh.SelLength = Len(txtbackfinhh.TEXT)
End If
End Sub


Private Sub txtbackfinwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackfinwt_LostFocus()
If Val(txtbackfinwt.TEXT) < 0 Or Val(txtbackfinwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 2
   txtbackfinwt.SetFocus
   txtbackfinwt.SelStart = 0
   txtbackfinwt.SelLength = Len(txtbackfinwt.TEXT)
End If
End Sub


Private Sub txtbackhh_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackhh_LostFocus()
If Val(txtbackhh.TEXT) < 0 Or Val(txtbackhh.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 2
   txtbackhh.SetFocus
   txtbackhh.SelStart = 0
   txtbackhh.SelLength = Len(txtbackhh.TEXT)
End If
End Sub


Private Sub txtbackintframe_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackinthh_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackinthh_LostFocus()
If Val(txtbackinthh.TEXT) < 0 Or Val(txtbackinthh.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 2
   txtbackinthh.SetFocus
   txtbackinthh.SelStart = 0
   txtbackinthh.SelLength = Len(txtbackinthh.TEXT)
End If
End Sub


Private Sub txtbackintwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackintwt_LostFocus()
If Val(txtbackintwt.TEXT) < 0 Or Val(txtbackintwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 2
   txtbackintwt.SetFocus
   txtbackintwt.SelStart = 0
   txtbackintwt.SelLength = Len(txtbackintwt.TEXT)
End If
End Sub


Private Sub txtbackrecframe_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackrecwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbackrecwt_LostFocus()
If Val(txtbackrecwt.TEXT) < 0 Or Val(txtbackrecwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 2
   txtbackrecwt.SetFocus
   txtbackrecwt.SelStart = 0
   txtbackrecwt.SelLength = Len(txtbackrecwt.TEXT)
End If
End Sub


Private Sub txtbirthwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtbirthwt_LostFocus()
If Val(txtbirthwt.TEXT) < 0 Or Val(txtbirthwt.TEXT) > 300 Then
   MsgBox "Please enter a number between 0 and 300", vbOKOnly
   txtbirthwt.SetFocus
   txtbirthwt.SelStart = 0
   txtbirthwt.SelLength = Len(txtbirthwt.TEXT)
Else
   txtbirthwt.TEXT = funround2(0, txtbirthwt.TEXT)
End If
End Sub


Private Sub txtbnotes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 9 Then Cmdsave.SetFocus
End Sub


Private Sub txtcalfbreed_GotFocus()
Dim cowbrd$, sirebrd$
Dim DB As database
Dim RS As Recordset
 If Trim(txtcalfbreed.TEXT) <> "" Then Exit Sub
Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
Set RS = DB.OpenRecordset("COWprof", dbOpenTable)
RS.Index = "primarykey"
RS.Seek "=", herdid$, txtcowid.TEXT
If Not RS.NoMatch Then cowbrd$ = Field2Str(RS!breed)
RS.Close
Set RS = DB.OpenRecordset("sireprof", dbOpenTable)
RS.Index = "primarykey"
RS.Seek "=", herdid$, Txtsireid.TEXT
If Not RS.NoMatch Then sirebrd$ = Field2Str(RS!breed)
If cowbrd$ = sirebrd$ Then
   txtcalfbreed.TEXT = sirebrd$
Else
   txtcalfbreed.TEXT = sirebrd$ & cowbrd$
End If
DB.Close: Set DB = Nothing
End Sub

Private Sub txtcarccarcwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcarccarcwt_LostFocus()
If Val(txtcarccarcwt.TEXT) < 0 Or Val(txtcarccarcwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 5
   txtcarccarcwt.SetFocus
   txtcarccarcwt.SelStart = 0
   txtcarccarcwt.SelLength = Len(txtcarccarcwt.TEXT)
End If
End Sub


Private Sub txtcarcfatthick_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcarcfatthick_LostFocus()
If Val(txtcarcfatthick.TEXT) < 0 Or Val(txtcarcfatthick.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 5
   txtcarcfatthick.SetFocus
   txtcarcfatthick.SelStart = 0
   txtcarcfatthick.SelLength = Len(txtcarcfatthick.TEXT)
End If
End Sub


Private Sub txtcarckph_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcarckph_LostFocus()
If Val(txtcarckph.TEXT) < 0 Or Val(txtcarckph.TEXT) > 10 Then
   MsgBox "Please enter a number between 0 and 10", vbOKOnly
   SSTab1.Tab = 5
   txtcarckph.SetFocus
   txtcarckph.SelStart = 0
   txtcarckph.SelLength = Len(txtcarckph.TEXT)
End If
End Sub


Private Sub txtcarcmuscle_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcarcnotes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 9 Then Cmdsave.SetFocus
End Sub


Private Sub txtcarcrib_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcarcrib_LostFocus()
If Val(txtcarcrib.TEXT) < 0 Or Val(txtcarcrib.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 5
   txtcarcrib.SetFocus
   txtcarcrib.SelStart = 0
   txtcarcrib.SelLength = Len(txtcarcrib.TEXT)
End If
End Sub


Private Sub txtcarcyield_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtcarcyield_LostFocus()
txtcarcyield = funround2(1, txtcarcyield)
End Sub

Private Sub txtcowage_GotFocus()
Dim RS As Recordset
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select max(cowage) as maxcowage from calfbirth where herdid = '" & herdid & "' and cowid = '" & txtcowid.TEXT & "' ", dbOpenSnapshot)
If Not RS.EOF Then TxtCowAge.TEXT = Field2Num(RS!maxcowage) + 1
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End Sub

Private Sub TxtCowAge_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtcowid_DblClick()
 Dim FrmSelCow As New selcow_list
 FrmSelCow.SetMode = 0
 FrmSelCow.Show vbModal
 If FrmSelCow.Tag = "CANCEL" Then Exit Sub
 txtcowid.TEXT = FrmSelCow.Tag
'FrmSelect_Multi_Cows.Show vbModal
'If FrmSelect_Multi_Cows.Tag <> "Cancel" Then txtcowid.TEXT = FrmSelect_Multi_Cows.Tag
Unload FrmSelCow: Set FrmSelCow = Nothing
End Sub


Private Sub txtfinfincond_Change()

End Sub

Private Sub txtcowid_LostFocus()
' Dim cowbrd$, sirebrd$
' Dim DB As database
' Dim TBCOWprof As Recordset
' Dim tbsireprof As Recordset
' Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
'
' 'Set TBCOWprof = DB.OpenRecordset("select max(cowage) as maxcowage from calfbirth where herdid = '" & herdid & "' and cowid = '" & txtcowid.TEXT & "' ", dbOpenSnapshot)
'
' 'If Not TBCOWprof.EOF Then txtcowage.TEXT = Field2Num(TBCOWprof!maxcowage) + 1 Else txtcowage.TEXT = "2"
'
' If txtcalfbreed.TEXT <> "" Then GoTo Exit_Sub
' If txtcowid.TEXT = "" Then GoTo Exit_Sub
' Set TBCOWprof = DB.OpenRecordset("COWprof", dbOpenTable)
' TBCOWprof.Index = "primarykey"
' TBCOWprof.Seek "=", herdid$, txtcowid.TEXT
' If Not TBCOWprof.NoMatch Then cowbrd$ = TBCOWprof!breed
' Set tbsireprof = DB.OpenRecordset("sireprof", dbOpenTable)
' tbsireprof.Index = "primarykey"
' tbsireprof.Seek "=", herdid$, Txtsireid.TEXT
' If Not tbsireprof.NoMatch Then sirebrd$ = tbsireprof!breed
' If cowbrd$ = sirebrd$ Then
'    txtcalfbreed.TEXT = sirebrd$
' Else
'    txtcalfbreed.TEXT = sirebrd$ & cowbrd$
' End If
'tbsireprof.Close: Set tbsireprof = Nothing
'Exit_Sub:
'TBCOWprof.Close: Set TBCOWprof = Nothing
'
' DB.Close: Set DB = Nothing
End Sub

Private Sub txtfinfinfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinfinfat_LostFocus()
If Val(txtfinfinfat.TEXT) < 0 Or Val(txtfinfinfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 4
   txtfinfinfat.SetFocus
   txtfinfinfat.SelStart = 0
   txtfinfinfat.SelLength = Len(txtfinfinfat.TEXT)
End If
End Sub


Private Sub txtfinfinmarbl_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinfinmarbl_LostFocus()
If Val(txtfinfinmarbl.TEXT) < 0 Or Val(txtfinfinmarbl.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 4
   txtfinfinmarbl.SetFocus
   txtfinfinmarbl.SelStart = 0
   txtfinfinmarbl.SelLength = Len(txtfinfinmarbl.TEXT)
End If
End Sub


Private Sub txtfinfinrea_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinfinrea_LostFocus()
If Val(txtfinfinrea.TEXT) < 0 Or Val(txtfinfinrea.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 4
   txtfinfinrea.SetFocus
   txtfinfinrea.SelStart = 0
   txtfinfinrea.SelLength = Len(txtfinfinrea.TEXT)
End If
End Sub


Private Sub txtfinfinwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinfinwt_LostFocus()
If Val(txtfinfinwt.TEXT) < 0 Or Val(txtfinfinwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 4
   txtfinfinwt.SetFocus
   txtfinfinwt.SelStart = 0
   txtfinfinwt.SelLength = Len(txtfinfinwt.TEXT)
End If
End Sub


Private Sub txtfinint2cond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinint2cond_LostFocus()
If Val(txtfinint2cond.TEXT) < 0 Or Val(txtfinint2cond.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 4
   txtfinint2cond.SetFocus
   txtfinint2cond.SelStart = 0
   txtfinint2cond.SelLength = Len(txtfinint2cond.TEXT)
End If
End Sub


Private Sub txtfinint2fat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinint2fat_LostFocus()
If Val(txtfinint2fat.TEXT) < 0 Or Val(txtfinint2fat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 4
   txtfinint2fat.SetFocus
   txtfinint2fat.SelStart = 0
   txtfinint2fat.SelLength = Len(txtfinint2fat.TEXT)
End If
End Sub


Private Sub txtfinint2wt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinint2wt_LostFocus()
If Val(txtfinint2wt.TEXT) < 0 Or Val(txtfinint2wt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 4
   txtfinint2wt.SetFocus
   txtfinint2wt.SelStart = 0
   txtfinint2wt.SelLength = Len(txtfinint2wt.TEXT)
End If
End Sub


Private Sub txtfinintcond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinintcond_LostFocus()
If Val(txtfinintcond.TEXT) < 0 Or Val(txtfinintcond.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 4
   txtfinintcond.SetFocus
   txtfinintcond.SelStart = 0
   txtfinintcond.SelLength = Len(txtfinintcond.TEXT)
End If
End Sub


Private Sub txtfinintfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinintfat_LostFocus()
If Val(txtfinintfat.TEXT) < 0 Or Val(txtfinintfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 4
   txtfinintfat.SetFocus
   txtfinintfat.SelStart = 0
   txtfinintfat.SelLength = Len(txtfinintfat.TEXT)
End If
End Sub


Private Sub txtfinintwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfinintwt_LostFocus()
If Val(txtfinintwt.TEXT) < 0 Or Val(txtfinintwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 4
   txtfinintwt.SetFocus
   txtfinintwt.SelStart = 0
   txtfinintwt.SelLength = Len(txtfinintwt.TEXT)
End If
End Sub


Private Sub txtflmar_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtflmar_LostFocus()
If Val(txtflmar.TEXT) < 0 Or Val(txtflmar.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 4
   txtflmar.SetFocus
   txtflmar.SelStart = 0
   txtflmar.SelLength = Len(txtflmar.TEXT)
End If
End Sub


Private Sub txtflrea_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtflrea_LostFocus()
If Val(txtflrea.TEXT) < 0 Or Val(txtflrea.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 4
   txtflrea.SetFocus
   txtflrea.SelStart = 0
   txtflrea.SelLength = Len(txtflrea.TEXT)
End If
End Sub


Private Sub txtflscore_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtfnotes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 9 Then Cmdsave.SetFocus
End Sub


Private Sub txtframe_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txthipht_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txthipht_LostFocus()
If dtemeas.TEXT = "--/--/----" Then Exit Sub
If Val(txthipht.TEXT) < 0 Or Val(txthipht.TEXT) > 75 Then
   MsgBox "Please enter a number between 0 and 75", vbOKOnly
   SSTab1.Tab = 1
   txthipht.SetFocus
   txthipht.SelStart = 0
   txthipht.SelLength = Len(txthipht.TEXT)
End If
End Sub


Private Sub txtid_GotFocus()
  SSTab1.Tab = 0

End Sub


Private Sub txtid_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then KeyAscii = 0
End Sub


Private Sub txtintcond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintcond_LostFocus()
If Val(txtintcond.TEXT) < 0 Or Val(txtintcond.TEXT) >= 10 Then
   MsgBox "Please enter a number between 0 and 9", vbOKOnly
   SSTab1.Tab = 3
   txtintcond.SetFocus
   txtintcond.SelStart = 0
   txtintcond.SelLength = Len(txtintcond.TEXT)
End If
End Sub


Private Sub txtintfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintfat_LostFocus()
If Val(txtintfat.TEXT) < 0 Or Val(txtintfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 3
   txtintfat.SetFocus
   txtintfat.SelStart = 0
   txtintfat.SelLength = Len(txtintfat.TEXT)
End If
End Sub


Private Sub txtinthh_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtinthh_LostFocus()
If Val(txtinthh.TEXT) < 0 Or Val(txtinthh.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 3
   txtinthh.SetFocus
   txtinthh.SelStart = 0
   txtinthh.SelLength = Len(txtinthh.TEXT)
End If
End Sub


Private Sub txtintmarb_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintmarb_LostFocus()
If Val(txtintmarb.TEXT) < 0 Or Val(txtintmarb.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 3
   txtintmarb.SetFocus
   txtintmarb.SelStart = 0
   txtintmarb.SelLength = Len(txtintmarb.TEXT)
End If
End Sub


Private Sub txtintrib_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintrib_LostFocus()
If Val(txtintrib.TEXT) < 0 Or Val(txtintrib.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 3
   txtintrib.SetFocus
   txtintrib.SelStart = 0
   txtintrib.SelLength = Len(txtintrib.TEXT)
End If
End Sub


Private Sub txtintscore_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtintwt_LostFocus()
If Val(txtintwt.TEXT) < 0 Or Val(txtintwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 3
   txtintwt.SetFocus
   txtintwt.SelStart = 0
   txtintwt.SelLength = Len(txtintwt.TEXT)
End If
End Sub


Private Sub txtpelvicsz_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtpelvicsz_LostFocus()
If Val(txtpelvicsz.TEXT) < 0 Or Val(txtpelvicsz.TEXT) > 300 Then
   MsgBox "Please enter a number between 0 and 300", vbOKOnly
   SSTab1.Tab = 3
   txtpelvicsz.SetFocus
   txtpelvicsz.SelStart = 0
   txtpelvicsz.SelLength = Len(txtpelvicsz.TEXT)
End If
End Sub


Private Sub txtqscore_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtqscore_LostFocus()
If Val(txtqscore.TEXT) < 0 Or Val(txtqscore.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 5
   txtqscore.SetFocus
   txtqscore.SelStart = 0
   txtqscore.SelLength = Len(txtqscore.TEXT)
End If
End Sub


Private Sub txtratio_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtreccond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtreccond_LostFocus()
If Val(txtreccond.TEXT) < 0 Or Val(txtreccond.TEXT) >= 10 Then
   MsgBox "Please enter a number between 0 and 9", vbOKOnly
   SSTab1.Tab = 3
   txtreccond.SetFocus
   txtreccond.SelStart = 0
   txtreccond.SelLength = Len(txtreccond.TEXT)
End If
End Sub


Private Sub txtrecfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecfat_LostFocus()
If Val(txtrecfat.TEXT) < 0 Or Val(txtrecfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 3
   txtrecfat.SetFocus
   txtrecfat.SelStart = 0
   txtrecfat.SelLength = Len(txtrecfat.TEXT)
End If
End Sub


Private Sub txtrecmar_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecmar_LostFocus()
If Val(txtrecmar.TEXT) < 0 Or Val(txtrecmar.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 4
   txtrecmar.SetFocus
   txtrecmar.SelStart = 0
   txtrecmar.SelLength = Len(txtrecmar.TEXT)
End If
End Sub


Private Sub txtrecmarb_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecmarb_LostFocus()
If Val(txtrecmarb.TEXT) < 0 Or Val(txtrecmarb.TEXT) > 2000 Then
   MsgBox "Please enter a number between 0 and 2000", vbOKOnly
   SSTab1.Tab = 3
   txtrecmarb.SetFocus
   txtrecmarb.SelStart = 0
   txtrecmarb.SelLength = Len(txtrecmarb.TEXT)
End If
End Sub


Private Sub txtrecrea_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecrea_LostFocus()
If Val(txtrecrea.TEXT) < 0 Or Val(txtrecrea.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 4
   txtrecrea.SetFocus
   txtrecrea.SelStart = 0
   txtrecrea.SelLength = Len(txtrecrea.TEXT)
End If
End Sub


Private Sub txtrecrib_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecrib_LostFocus()
If Val(txtrecrib.TEXT) < 0 Or Val(txtrecrib.TEXT) > 50 Then
   MsgBox "Please enter a number between 0 and 50", vbOKOnly
   SSTab1.Tab = 3
   txtrecrib.SetFocus
   txtrecrib.SelStart = 0
   txtrecrib.SelLength = Len(txtrecrib.TEXT)
End If
End Sub


Private Sub txtrecscore_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtrecwt_LostFocus()
If Val(txtrecwt.TEXT) < 0 Or Val(txtrecwt.TEXT) > 3000 Then
   MsgBox "Please enter a number between 0 and 3000", vbOKOnly
   SSTab1.Tab = 3
   txtrecwt.SetFocus
   txtrecwt.SelStart = 0
   txtrecwt.SelLength = Len(txtrecwt.TEXT)
End If
End Sub


Private Sub txtscrotumcir_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtscrotumcir_LostFocus()
If Val(txtscrotumcir.TEXT) < 0 Or Val(txtscrotumcir.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 3
   txtscrotumcir.SetFocus
   txtscrotumcir.SelStart = 0
   txtscrotumcir.SelLength = Len(txtscrotumcir.TEXT)
End If
End Sub


Private Sub Txtsireid_DblClick()
   selsire_list.Show vbModal
   If selsire_list.Tag = "CANCEL" Then Exit Sub
   Txtsireid.TEXT = selsire_list.Tag
End Sub


Private Sub txtweight365_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtwnotes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 9 Then Cmdsave.SetFocus
End Sub


Private Sub txtyframe_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtyhipht_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtyhipht_LostFocus()
If Val(txtyhipht.TEXT) < 0 Or Val(txtyhipht.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 3
   txtyhipht.SetFocus
   txtyhipht.SelStart = 0
   txtyhipht.SelLength = Len(txtyhipht.TEXT)
End If
End Sub


