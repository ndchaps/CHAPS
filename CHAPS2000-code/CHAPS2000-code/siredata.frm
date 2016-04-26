VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmsire_data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sire Information"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Prev"
      Height          =   375
      Left            =   4170
      TabIndex        =   90
      Top             =   5280
      Width           =   1050
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5280
      TabIndex        =   91
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Height          =   385
      Left            =   6360
      TabIndex        =   92
      Top             =   5265
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   7440
      TabIndex        =   93
      Top             =   5265
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Profile"
      TabPicture(0)   =   "siredata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EPD"
      TabPicture(1)   =   "siredata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   37
         Top             =   360
         Width           =   8175
         Begin VB.Frame fraepd 
            Caption         =   "EPD'S"
            Height          =   4335
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   7935
            Begin VB.TextBox txtmisc6 
               Height          =   285
               Left            =   5265
               MaxLength       =   10
               TabIndex        =   76
               Top             =   2400
               Width           =   1000
            End
            Begin VB.TextBox txtmisc7 
               Height          =   285
               Left            =   5265
               MaxLength       =   10
               TabIndex        =   79
               Top             =   2730
               Width           =   1000
            End
            Begin VB.TextBox txtmisc8 
               Height          =   285
               Left            =   5265
               MaxLength       =   10
               TabIndex        =   82
               Top             =   3060
               Width           =   1000
            End
            Begin VB.TextBox txtmisc9 
               Height          =   285
               Left            =   5265
               MaxLength       =   10
               TabIndex        =   85
               Top             =   3360
               Width           =   1000
            End
            Begin VB.TextBox txtmisc10 
               Height          =   285
               Left            =   5265
               MaxLength       =   10
               TabIndex        =   88
               Top             =   3700
               Width           =   1000
            End
            Begin VB.TextBox txtacc6 
               Height          =   285
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   77
               Top             =   2400
               Width           =   800
            End
            Begin VB.TextBox txtacc7 
               Height          =   285
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   80
               Top             =   2730
               Width           =   800
            End
            Begin VB.TextBox txtacc8 
               Height          =   285
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   83
               Top             =   3060
               Width           =   800
            End
            Begin VB.TextBox txtacc9 
               Height          =   285
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   86
               Top             =   3360
               Width           =   800
            End
            Begin VB.TextBox txtacc10 
               Height          =   285
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   89
               Top             =   3705
               Width           =   800
            End
            Begin VB.TextBox txtacc1 
               Height          =   285
               Left            =   3120
               MaxLength       =   10
               TabIndex        =   60
               Top             =   2400
               Width           =   800
            End
            Begin VB.TextBox txtacc2 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3120
               MaxLength       =   10
               TabIndex        =   63
               Top             =   2730
               Width           =   800
            End
            Begin VB.TextBox txtacc3 
               Height          =   285
               Left            =   3120
               MaxLength       =   10
               TabIndex        =   66
               Top             =   3060
               Width           =   800
            End
            Begin VB.TextBox txtacc4 
               Height          =   285
               Left            =   3120
               MaxLength       =   10
               TabIndex        =   69
               Top             =   3360
               Width           =   800
            End
            Begin VB.TextBox txtacc5 
               Height          =   285
               Left            =   3120
               MaxLength       =   10
               TabIndex        =   72
               Top             =   3705
               Width           =   800
            End
            Begin VB.TextBox txtepdaccbwt 
               Height          =   285
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   43
               Top             =   510
               Width           =   800
            End
            Begin VB.TextBox txtepdaccwwt 
               Height          =   285
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   46
               Top             =   840
               Width           =   800
            End
            Begin VB.TextBox txtepdaccywt 
               Height          =   285
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   49
               Top             =   1170
               Width           =   800
            End
            Begin VB.TextBox txtepdaccmatww 
               Height          =   285
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   52
               Top             =   1500
               Width           =   800
            End
            Begin VB.TextBox txtepdaccmatmilk 
               Height          =   285
               Left            =   4800
               MaxLength       =   10
               TabIndex        =   55
               Top             =   1830
               Width           =   800
            End
            Begin VB.TextBox txtbirthwt 
               Height          =   285
               Left            =   3240
               MaxLength       =   10
               TabIndex        =   42
               Top             =   510
               Width           =   1290
            End
            Begin VB.TextBox txtmatmilk 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3240
               MaxLength       =   10
               TabIndex        =   54
               Top             =   1830
               Width           =   1290
            End
            Begin VB.TextBox txtmisc1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   59
               Top             =   2400
               Width           =   1290
            End
            Begin VB.TextBox txtmisc2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   62
               Top             =   2730
               Width           =   1290
            End
            Begin VB.TextBox txtmisc3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   65
               Top             =   3060
               Width           =   1290
            End
            Begin VB.TextBox txtmisc4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   68
               Top             =   3390
               Width           =   1290
            End
            Begin VB.TextBox txtmisc5 
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   71
               Top             =   3705
               Width           =   1290
            End
            Begin VB.TextBox txtweanwt 
               Height          =   285
               Left            =   3240
               MaxLength       =   10
               TabIndex        =   45
               Top             =   840
               Width           =   1290
            End
            Begin VB.TextBox txtyearwt 
               Height          =   285
               Left            =   3240
               MaxLength       =   10
               TabIndex        =   48
               Top             =   1170
               Width           =   1290
            End
            Begin VB.TextBox txtmatww 
               Height          =   285
               Left            =   3240
               MaxLength       =   10
               TabIndex        =   51
               Top             =   1500
               Width           =   1290
            End
            Begin VB.Label lblmisc6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 6"
               Height          =   255
               Left            =   4100
               TabIndex        =   75
               Top             =   2400
               Width           =   1100
            End
            Begin VB.Label lblmisc7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 7"
               Height          =   255
               Left            =   4100
               TabIndex        =   78
               Top             =   2730
               Width           =   1100
            End
            Begin VB.Label lblmisc8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 8"
               Height          =   255
               Left            =   4100
               TabIndex        =   81
               Top             =   3060
               Width           =   1100
            End
            Begin VB.Label lblmisc9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 9"
               Height          =   255
               Left            =   4100
               TabIndex        =   84
               Top             =   3360
               Width           =   1100
            End
            Begin VB.Label lblmisc10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 10"
               Height          =   255
               Left            =   4100
               TabIndex        =   87
               Top             =   3700
               Width           =   1100
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   5340
               TabIndex        =   73
               Top             =   2160
               Width           =   1005
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   6375
               TabIndex        =   74
               Top             =   2160
               Width           =   1005
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   1680
               TabIndex        =   56
               Top             =   2160
               Width           =   1005
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   3000
               TabIndex        =   57
               Top             =   2160
               Width           =   1005
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   4560
               TabIndex        =   40
               Top             =   220
               Width           =   975
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   3360
               TabIndex        =   39
               Top             =   225
               Width           =   1005
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Wt"
               Height          =   270
               Left            =   2145
               TabIndex        =   41
               Top             =   510
               Width           =   1005
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Maternal Milk"
               Height          =   240
               Left            =   2145
               TabIndex        =   53
               Top             =   1830
               Width           =   1005
            End
            Begin VB.Label lblmisc1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " Misc 1"
               Height          =   240
               Left            =   390
               TabIndex        =   58
               Top             =   2400
               Width           =   1095
            End
            Begin VB.Label lblmisc2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 2"
               Height          =   255
               Left            =   390
               TabIndex        =   61
               Top             =   2715
               Width           =   1100
            End
            Begin VB.Label lblmisc3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 3"
               Height          =   255
               Left            =   390
               TabIndex        =   64
               Top             =   3060
               Width           =   1100
            End
            Begin VB.Label lblmisc4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 4"
               Height          =   240
               Left            =   390
               TabIndex        =   67
               Top             =   3390
               Width           =   1100
            End
            Begin VB.Label lblmisc5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 5"
               Height          =   255
               Left            =   390
               TabIndex        =   70
               Top             =   3705
               Width           =   1100
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Wean Wt"
               Height          =   255
               Left            =   2145
               TabIndex        =   44
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Yearling Wt"
               Height          =   255
               Left            =   2145
               TabIndex        =   47
               Top             =   1170
               Width           =   1005
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Maternal"
               Height          =   255
               Left            =   2055
               TabIndex        =   50
               Top             =   1500
               Width           =   1100
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         Begin VB.ComboBox CBOReasonCode 
            Height          =   315
            Left            =   4365
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   3660
            Width           =   615
         End
         Begin VB.TextBox txtregname 
            Height          =   285
            Left            =   1695
            MaxLength       =   40
            TabIndex        =   9
            Top             =   1440
            Width           =   2325
         End
         Begin VB.TextBox txtprofnotes 
            Height          =   855
            Left            =   5700
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   2415
            Width           =   2235
         End
         Begin VB.TextBox txtbirthid 
            Height          =   285
            Left            =   3090
            TabIndex        =   21
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            Height          =   1095
            Left            =   1680
            TabIndex        =   26
            Top             =   3240
            Width           =   1335
            Begin VB.OptionButton optactive 
               Caption         =   "Active"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   225
               Width           =   975
            End
            Begin VB.OptionButton optculled 
               Caption         =   "Culled"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   495
               Width           =   975
            End
            Begin VB.OptionButton optpedigree 
               Caption         =   "Pedigree"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   750
               Width           =   1095
            End
         End
         Begin VB.Frame frasource 
            Caption         =   "Source"
            Height          =   1095
            Left            =   1680
            TabIndex        =   22
            Top             =   2160
            Width           =   1335
            Begin VB.OptionButton optai 
               Caption         =   "A.I."
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   770
               Width           =   975
            End
            Begin VB.OptionButton optraised 
               Caption         =   "Raised"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optpurchased 
               Caption         =   "Purchased"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.TextBox txtcomments 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4365
            MaxLength       =   25
            TabIndex        =   34
            Top             =   4080
            Width           =   1650
         End
         Begin VB.TextBox txtsireofcow 
            Height          =   285
            Left            =   5700
            MaxLength       =   8
            TabIndex        =   15
            Top             =   795
            Width           =   1215
         End
         Begin VB.TextBox txtdamofcow 
            Height          =   285
            Left            =   5685
            TabIndex        =   17
            Top             =   1290
            Width           =   1215
         End
         Begin VB.TextBox txtbreed 
            Height          =   285
            Left            =   1695
            MaxLength       =   8
            TabIndex        =   5
            Top             =   705
            Width           =   1170
         End
         Begin VB.TextBox TXTID 
            Height          =   285
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   3
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txtregnum 
            Height          =   285
            Left            =   1695
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1065
            Width           =   2325
         End
         Begin VB.TextBox txteleid 
            Height          =   285
            Left            =   1695
            MaxLength       =   15
            TabIndex        =   11
            Top             =   1815
            Width           =   2325
         End
         Begin MSMask.MaskEdBox Dteculled 
            Height          =   285
            Left            =   4365
            TabIndex        =   31
            Top             =   3345
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox dteentered 
            Height          =   285
            Left            =   5685
            TabIndex        =   19
            Top             =   1755
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSMask.MaskEdBox MHBIRTHDATE 
            Height          =   285
            Left            =   5700
            TabIndex        =   13
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
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
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Premise ID#"
            Height          =   255
            Left            =   615
            TabIndex        =   8
            Top             =   1485
            Width           =   1005
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Profile Notes"
            Height          =   255
            Left            =   4560
            TabIndex        =   35
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblbirthid 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Calf ID at Birth"
            Height          =   255
            Left            =   3120
            TabIndex        =   20
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Entered Herd"
            Height          =   255
            Left            =   4200
            TabIndex        =   18
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lbldateculled 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Culled"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            Top             =   3360
            Width           =   1005
         End
         Begin VB.Label lblcomments 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   33
            Top             =   4080
            Width           =   1005
         End
         Begin VB.Label lblreasoncode 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Reason Code"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   32
            Top             =   3720
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sire of Sire"
            Height          =   255
            Left            =   4680
            TabIndex        =   14
            Top             =   825
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dam of Sire"
            Height          =   255
            Left            =   4560
            TabIndex        =   16
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sire ID"
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   270
            Left            =   4545
            TabIndex        =   12
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sire Breed"
            Height          =   270
            Left            =   600
            TabIndex        =   4
            Top             =   735
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registration #"
            Height          =   255
            Left            =   600
            TabIndex        =   6
            Top             =   1110
            Width           =   1005
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Electronic Id"
            Height          =   270
            Left            =   600
            TabIndex        =   10
            Top             =   1815
            Width           =   1005
         End
      End
   End
End
Attribute VB_Name = "frmsire_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addedflag$
Dim dirtyflag%
Dim oldid$
Dim tbData As Recordset
Dim save%

Private Sub Init_Information()
 Call init_form(Me) ' Clear Text Boxes
 'CBOReasonCode.AddItem " "
 CBOReasonCode.AddItem "G"
 CBOReasonCode.AddItem "H"
 CBOReasonCode.AddItem "J"
 CBOReasonCode.AddItem "L"
 CBOReasonCode.AddItem "R"
 CBOReasonCode.AddItem "Y"
'Me.Refresh
End Sub

Private Sub Load_information()
 Screen.MousePointer = vbHourglass
 Dim tbsireepd As Recordset
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbsireepd = DB.OpenRecordset("sireepd", dbOpenTable)
 tbsireepd.Index = "primarykey"
 tbsireepd.Seek "=", herdid$, oldid$
 Set tbData = DB.OpenRecordset("sireprof", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 If Not tbData.NoMatch Then
   txtid.TEXT = Field2Str(tbData!sireid)
   Mhbirthdate.TEXT = Field2Date(tbData!birthdate)
   txtbreed.TEXT = Field2Str(tbData!breed)
   txtregnum.TEXT = Field2Str(tbData!regnum)
   TXTREGNAME.TEXT = Field2Str(tbData!regname)
   txteleid.TEXT = Field2Str(tbData!elecid)
   txtsireofcow.TEXT = Field2Str(tbData!sire)
   txtdamofcow.TEXT = Field2Str(tbData!dam)
   txtproFnotes.TEXT = Field2Str(tbData!notes)
   dteentered.TEXT = Field2Date(tbData!enteredherd)
   txtbirthid.TEXT = Field2Str(tbData!calfid)
   If Field2Str(tbData!Source) = "R" Then optraised.Value = True
   If Field2Str(tbData!Source) = "P" Then optpurchased.Value = True
   If Field2Str(tbData!Source) = "A" Then optai.Value = True
   If Field2Str(tbData!active) = "A" Then optactive.Value = True
   If Field2Str(tbData!active) = "C" Then optculled.Value = True
   If Field2Str(tbData!active) = "P" Then optpedigree.Value = True
   Dteculled.TEXT = Field2Date(tbData!dateculled)
   Call set_combo(CBOReasonCode, Field2Str(tbData!reasonculled))
   txtcomments.TEXT = Field2Str(tbData!commentsculled)
 End If
 If Not tbsireepd.NoMatch Then
   txtbirthwt.TEXT = Field2Str(tbsireepd!epdbirthwt)
   txtweanwt.TEXT = Field2Str(tbsireepd!epdweanwt)
   txtyearwt.TEXT = Field2Str(tbsireepd!epdyearwt)
   txtmatww.TEXT = Field2Str(tbsireepd!epdmatww)
   txtmatmilk.TEXT = Field2Str(tbsireepd!epdmatmilk)
   txtepdaccbwt.TEXT = Field2Str(tbsireepd!accbirthwt)
   txtepdaccwwt.TEXT = Field2Str(tbsireepd!accweanwt)
   txtepdaccywt.TEXT = Field2Str(tbsireepd!accyearwt)
   txtepdaccmatww.TEXT = Field2Str(tbsireepd!accmatww)
   txtepdaccmatmilk.TEXT = Field2Str(tbsireepd!accmatmilk)
   txtmisc1.TEXT = Field2Str(tbsireepd!misc1)
   txtmisc2.TEXT = Field2Str(tbsireepd!misc2)
   Txtmisc3.TEXT = Field2Str(tbsireepd!misc3)
   txtmisc4.TEXT = Field2Str(tbsireepd!misc4)
   txtmisc5.TEXT = Field2Str(tbsireepd!misc5)
   txtmisc6.TEXT = Field2Str(tbsireepd!misc6)
   txtmisc7.TEXT = Field2Str(tbsireepd!misc7)
   txtmisc8.TEXT = Field2Str(tbsireepd!misc8)
   txtmisc9.TEXT = Field2Str(tbsireepd!misc9)
   txtmisc10.TEXT = Field2Str(tbsireepd!misc10)
   txtacc1.TEXT = Field2Str(tbsireepd!acc1)
   txtacc2.TEXT = Field2Str(tbsireepd!acc2)
   txtacc3.TEXT = Field2Str(tbsireepd!acc3)
   txtacc4.TEXT = Field2Str(tbsireepd!acc4)
   txtacc5.TEXT = Field2Str(tbsireepd!acc5)
   txtacc6.TEXT = Field2Str(tbsireepd!acc6)
   txtacc7.TEXT = Field2Str(tbsireepd!acc7)
   txtacc8.TEXT = Field2Str(tbsireepd!acc8)
   txtacc9.TEXT = Field2Str(tbsireepd!acc9)
   txtacc10.TEXT = Field2Str(tbsireepd!acc10)
End If
 tbData.Close: Set tbData = Nothing
 tbsireepd.Close: Set tbsireepd = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub save_information()
 Dim Replace$
 Dim tbsireepd As Recordset
 
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("sireprof", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 save% = True
 If Not tbData.NoMatch Then
   If addedflag$ = "D" Then
     tbData.Delete
     Replace$ = ""
     save% = False
     GoSub donedel
    Else
     tbData.Edit
   End If
  Else
   tbData.AddNew
 End If
 If save% Then ' if in add or edit mode save the information on the form
   With tbData
       !herdid = herdid$
       !sireid = txtid.TEXT
       Call Date2Field(!birthdate, Mhbirthdate.TEXT)
'       !birthdate = THEDATE
       !breed = txtbreed.TEXT
       !regnum = txtregnum.TEXT
       !regname = TXTREGNAME.TEXT
       !elecid = txteleid.TEXT
       !sire = txtsireofcow.TEXT
       !dam = txtdamofcow.TEXT
       !notes = txtproFnotes.TEXT
       Call Date2Field(!enteredherd, dteentered.TEXT)
'       !enteredherd = THEDATE
       !calfid = txtbirthid.TEXT
       If optraised Then !Source = "R"
       If optpurchased Then !Source = "P"
       If optai Then !Source = "A"
       If optactive Then !active = "A"
       If optculled Then !active = "C"
       If optpedigree Then !active = "P"
       Call Date2Field(!dateculled, Dteculled.TEXT)
       '!dateculled = THEDATE
       !reasonculled = CBOReasonCode.TEXT
       !commentsculled = txtcomments.TEXT
       .Update
       Replace$ = txtid.TEXT
   End With
 Set tbsireepd = DB.OpenRecordset("sireepd", dbOpenTable)
 tbsireepd.Index = "primarykey"
 tbsireepd.Seek "=", herdid$, oldid$
   If Not tbsireepd.NoMatch Then
   If addedflag$ = "D" Then
     tbsireepd.Delete
     Replace$ = ""
     save% = False
    Else
     tbsireepd.Edit
   End If
  Else
   tbsireepd.AddNew
 End If
   With tbsireepd
       !herdid = herdid$
       !sireid = txtid.TEXT
       !epdbirthwt = Val(txtbirthwt.TEXT)
       !epdweanwt = Val(txtweanwt.TEXT)
       !epdyearwt = Val(txtyearwt.TEXT)
       !epdmatww = Val(txtmatww.TEXT)
       !epdmatmilk = Val(txtmatmilk.TEXT)
       !accbirthwt = Val(txtepdaccbwt.TEXT)
       !accweanwt = Val(txtepdaccwwt.TEXT)
       !accyearwt = Val(txtepdaccywt.TEXT)
       !accmatww = Val(txtepdaccmatww.TEXT)
       !accmatmilk = Val(txtepdaccmatmilk.TEXT)
       !misc1 = txtmisc1.TEXT
       !misc2 = txtmisc2.TEXT
       !misc3 = Txtmisc3.TEXT
       !misc4 = txtmisc4.TEXT
       !misc5 = txtmisc5.TEXT
       !misc5 = txtmisc5.TEXT
       !misc6 = txtmisc6.TEXT
       !misc7 = txtmisc7.TEXT
       !misc8 = txtmisc8.TEXT
       !misc9 = txtmisc9.TEXT
       !misc10 = txtmisc10.TEXT
       !acc1 = txtacc1.TEXT
       !acc2 = txtacc2.TEXT
       !acc3 = txtacc3.TEXT
       !acc4 = txtacc4.TEXT
       !acc5 = txtacc5.TEXT
       !acc5 = txtacc5.TEXT
       !acc6 = txtacc6.TEXT
       !acc7 = txtacc7.TEXT
       !acc8 = txtacc8.TEXT
       !acc9 = txtacc9.TEXT
       !acc10 = txtacc10.TEXT
       .Update
   End With
 End If
 tbsireepd.Close: Set tbsireepd = Nothing
donedel:
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 dirtyflag% = False
 Call Update_mh_ListBoxes("lstsire", 0, oldid$, Replace$)
 Screen.MousePointer = vbDefault
End Sub


Private Sub valid_form(exitcode%)
    exitcode% = 0
    If herdid = "" Then
        Beep
        MsgBox "Please Select A Herd", vbOKOnly + vbCritical, Me.Caption
        selherd_List.cmdcancel.Visible = False
        selherd_List.Show vbModal
    End If
    If txtid.TEXT = "" Then
        Beep
        MsgBox "Sire ID Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
        SSTab1.Tab = 0
        txtid.SetFocus
        exitcode% = 1
        Exit Sub
    End If
    If txteleid.TEXT <> "" Then
      If Len(txteleid) <> 15 Then
            MsgBox "EID must be 15 characters.", vbOKOnly
            exitcode% = 1
            txteleid.SetFocus
            Exit Sub
      End If
    End If
    
    If UCase$(oldid$) <> UCase$(txtid.TEXT) Then
        Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
        Set tbData = DB.OpenRecordset("sirePROF", dbOpenTable)
        tbData.Index = "primarykey"
        tbData.Seek "=", herdid$, txtid.TEXT
        If Not tbData.NoMatch Then
            Beep
            MsgBox "Sire ID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
            exitcode% = 1
            tbData.Close: Set tbData = Nothing
            DB.Close: Set DB = Nothing
            txtid.SetFocus
            Exit Sub
        End If
        tbData.Close: Set tbData = Nothing
        DB.Close: Set DB = Nothing
    End If
    
    If Check_EID(txteleid.TEXT, "Sire", oldid, herdid, txtbirthid.TEXT, "Sire") = False Then
       Beep
       MsgBox "EID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
       exitcode% = 1
    End If
    
End Sub




Private Sub CMDCancel_Click()
 Unload Me
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
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("select * from sireprof where herdid='" & herdid$ & "'", dbOpenDynaset)
 'tbdata.Index = "primarykey"
 tbData.FindFirst "herdid = '" & herdid$ & "' And sireid = '" & txtid.TEXT & "'"
 tbData.MoveNext
 If tbData.EOF Then
  tbData.MoveFirst
  frmsire_data.Tag = "E/" & tbData!sireid
 Else
  'tbdata.MoveNext
  frmsire_data.Tag = "E/" & tbData!sireid
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
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
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("select * from sireprof where herdid='" & herdid$ & "'", dbOpenDynaset)
 'tbdata.Index = "primarykey"
 tbData.FindFirst "herdid = '" & herdid$ & "' And sireid = '" & txtid.TEXT & "'"
 tbData.MovePrevious
 If tbData.BOF Then
  tbData.MoveLast
  frmsire_data.Tag = "E/" & tbData!sireid
 Else
  'tbdata.MoveNext
  frmsire_data.Tag = "E/" & tbData!sireid
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
 Call Form_Activate
End Sub


Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim TableName$(100)
 Dim iRet%
 
 If gIsDemo Then
   If IsValidSireEntry = False Then MsgBox "The demo version of this software only allows ten sire records.", vbOKOnly, "C.H.A.P.S. Demo": Exit Sub
 End If
 
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
 If addedflag$ = "D" Then
   Call CheckID(dbfile$, "sireprof", oldid$, TableName$())
   RESPONSE% = vbYes
   If Val(TableName$(0)) > 1 Then
     Beep
     RESPONSE% = MsgBox("Warning This Sire Id Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
   End If
   If RESPONSE% = vbYes Then Call save_information
 End If
 If addedflag$ <> "D" Then
   If optpedigree.Value = True Then
      iRet = MsgBox("Pedigree option is to record parental performance data for sires or dams" & vbCrLf & "that were not raised within this herd.  Do you wish to continue?", vbYesNo, Me.Caption)
      If iRet = vbNo Then Exit Sub
   End If
      Call save_information
 End If
 If addedflag$ = "A" Then
   Me.Tag = "A"
   Call Form_Activate
   txtid.SetFocus
  Else
   Unload Me
 End If
End Sub

Private Sub Form_Activate()
 If Me.Tag = "" Then Exit Sub
 addedflag$ = Left$(Me.Tag, 1)
 Me.Caption = "Sire Information" & " for Herd " & herdid$
 Screen.MousePointer = vbHourglass
 Call Init_Information
 lblmisc1 = epdhead1$
 lblmisc2 = epdhead2$
 lblmisc3 = epdhead3$
 lblmisc4 = epdhead4$
 lblmisc5 = epdhead5$
 lblmisc6 = epdhead6$
 lblmisc7 = epdhead7$
 lblmisc8 = epdhead8$
 lblmisc9 = epdhead9$
 lblmisc10 = epdhead10$
 If addedflag$ = "A" Then
    cmdnext.Enabled = False
    cmdprev.Enabled = False
   'Me.caption = "Add"
    oldid$ = ""
    If optactive.Value = False And optculled.Value = False Then optactive.Value = True
    If optraised.Value = False And optpurchased.Value = False And optai.Value = False Then optraised.Value = True
    If optraised.Value = True Then txtbirthid.Enabled = True: lblbirthid.Enabled = True Else txtbirthid.Enabled = False: lblbirthid.Enabled = False
    If optculled.Value = False Then
       CBOReasonCode.Enabled = False
       txtcomments.Enabled = False
       Dteculled.Enabled = False
      Else
       CBOReasonCode.Enabled = True
       txtcomments.Enabled = True
       Dteculled.Enabled = True
    End If

    
 End If
 If addedflag$ = "E" Or addedflag$ = "D" Then
   oldid$ = Trim$(Mid$(Me.Tag, 3))
   'Me.caption = "Edit"
   Me.Caption = "Sire Information" & " for Herd " & herdid$ & " - Sire " & oldid$
   Call Load_information
   If addedflag$ = "D" Then
     'Me.caption = "Delete"
     Call disable_controls(Me)
     Cmdsave.Caption = "&Delete"
     Cmdsave.Enabled = True
     cmdcancel.Enabled = True
     SSTab1.Enabled = True
   End If
 End If
 Me.Tag = ""
 SSTab1.Tab = 0
 Screen.MousePointer = vbDefault
 Me.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
End Sub


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
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
 Set frmsire_data = Nothing
End Sub




Private Sub optactive_Click()
 If optculled.Value = False Then
   CBOReasonCode.ListIndex = -1
   Dteculled.TEXT = "--/--/----"
   Dteculled.Enabled = False
   CBOReasonCode.Enabled = False
   txtcomments.Enabled = False
   lblreasoncode.Enabled = False
   lbldateculled.Enabled = False
   lblcomments.Enabled = False
  Else
   Dteculled.Enabled = True
   CBOReasonCode.Enabled = True
   txtcomments.Enabled = True
   lblreasoncode.Enabled = True
   lbldateculled.Enabled = True
   lblcomments.Enabled = True
 End If

End Sub


Private Sub optai_Click()
If optraised.Value = True Then lblbirthid.Enabled = True
End Sub

Private Sub optculled_Click()
 If optculled.Value = False Then
   Dteculled.TEXT = "--/--/----"
   CBOReasonCode.ListIndex = -1
   Dteculled.Enabled = False
   CBOReasonCode.Enabled = False
   txtcomments.Enabled = False
   lblreasoncode.Enabled = False
   lbldateculled.Enabled = False
   lblcomments.Enabled = False
  Else
   Dteculled.Enabled = True
   CBOReasonCode.Enabled = True
   txtcomments.Enabled = True
   lblreasoncode.Enabled = True
   lbldateculled.Enabled = True
   lblcomments.Enabled = True
 End If

End Sub


Private Sub optpedigree_Click()
 If optculled.Value = False Then
   Dteculled.TEXT = "--/--/----"
   CBOReasonCode.ListIndex = -1
   Dteculled.Enabled = False
   CBOReasonCode.Enabled = False
   txtcomments.Enabled = False
   lblreasoncode.Enabled = False
   lbldateculled.Enabled = False
   lblcomments.Enabled = False
  'MsgBox "Pedigree option is to record parental performance for sire or dams that was not raised in this operation", vbOKOnly + vbInformation, Me.Caption
  Else
   Dteculled.Enabled = True
   CBOReasonCode.Enabled = True
   txtcomments.Enabled = True
   lblreasoncode.Enabled = True
   lbldateculled.Enabled = True
   lblcomments.Enabled = True
 End If

End Sub


Private Sub optpurchased_Click()
  If optpurchased.Value = True Then lblbirthid.Enabled = False
End Sub

Private Sub optraised_Click()
  If optraised.Value = True Then lblbirthid.Enabled = True
End Sub

Private Sub SSTab1_GotFocus()
  If addedflag$ = "D" Then Exit Sub
  If SSTab1.Tab = 0 Then
     txtid.SetFocus
  End If
  If SSTab1.Tab = 1 Then
     txtbirthwt.SetFocus
     lblmisc1 = epdhead1$
     lblmisc2 = epdhead2$
     lblmisc3 = epdhead3$
     lblmisc4 = epdhead4$
     lblmisc5 = epdhead5$
     lblmisc6 = epdhead6$
     lblmisc7 = epdhead7$
     lblmisc8 = epdhead8$
     lblmisc9 = epdhead9$
     lblmisc10 = epdhead10$
  End If
  Me.Caption = "Sire Information" & " for Herd " & herdid$ & " - Sire " & txtid.TEXT

End Sub


Private Sub txtbirthid_DblClick()
Dim tbCalfData As Recordset, pDB As database, pSQL$
On Local Error GoTo ErrHandler
 selcalf_list.SetSex = 1
 selcalf_list.Show vbModal
 If selcalf_list.Tag = "CANCEL" Then Exit Sub
 txtbirthid.TEXT = selcalf_list.Tag
Set pDB = DBEngine(0).OpenDatabase(dbfile$, False, False)
pSQL = "SELECT DISTINCTROW calfbirth.breed, calfbirth.registration, calfbirth.regname, calfbirth.elecid, calfbirth.birthdate, " & _
    "calfbirth.sireID, sireprof.dam, calfbirth.notes FROM sireprof INNER JOIN calfbirth ON (calfbirth.HerdID = sireprof.HerdID) " & _
    "AND (sireprof.SireID = calfbirth.sireID) where calfbirth.herdid = '" & herdid & "' and calfbirth.calfid = '" & txtbirthid.TEXT & "'"
    
Set tbCalfData = pDB.OpenRecordset(pSQL, dbOpenSnapshot)
If tbCalfData.EOF Then GoTo Close_DB
With tbCalfData
10    txtbreed = Field2Str(!breed)
20    txtregnum = Field2Str(!registration)
30    TXTREGNAME = Field2Str(!regname)
40    txteleid = Field2Str(!elecid)
50    Mhbirthdate = Field2Date(!birthdate)
60    txtsireofcow = Field2Str(!sireid)
70    txtdamofcow = Field2Str(!dam)
    'txtprofnotes = Field2Str(!notes)
End With
optraised.Value = True
Close_DB:
tbCalfData.Close: Set tbCalfData = Nothing
pDB.Close: Set pDB = Nothing
Exit Sub
ErrHandler:
TEXT(2) = Erl
GMODNAME$ = Me.Name & " - txtbirthid_DblClick"
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())
End Sub


Private Sub txtbirthwt_GotFocus()
  SSTab1.Tab = 1
End Sub


Private Sub txtbirthwt_KeyPress(KeyAscii As Integer)
   'If KeyAscii = 8 And txtbirthwt.SelStart = 0 Then
   '  SSTab1.Tab = 0
   '  txtid.SetFocus
   'End If
KeyAscii = NumericOnly(KeyAscii)
End Sub


Private Sub txtdamofcow_DblClick()
 Dim FrmSelCow As New selcow_list
 FrmSelCow.SetMode = 1
 FrmSelCow.Show vbModal
 If FrmSelCow.Tag = "CANCEL" Then Exit Sub
 txtdamofcow.TEXT = FrmSelCow.Tag
 Unload FrmSelCow: Set FrmSelCow = Nothing
End Sub


Private Sub txtid_GotFocus()
   SSTab1.Tab = 0
End Sub


Private Sub txtid_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then KeyAscii = 0
End Sub


Private Sub txtmatmilk_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtmatww_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtsireofcow_DblClick()
 selsire_list.Show vbModal
 If selsire_list.Tag = "CANCEL" Then Exit Sub
 txtsireofcow.TEXT = selsire_list.Tag
End Sub


Private Sub txtweanwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtyearwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub
