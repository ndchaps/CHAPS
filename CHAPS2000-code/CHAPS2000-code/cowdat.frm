VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmcow_data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cow Information"
   ClientHeight    =   5760
   ClientLeft      =   2085
   ClientTop       =   1785
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   8160
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Prev"
      Height          =   375
      Left            =   2850
      TabIndex        =   147
      Top             =   5310
      Width           =   1095
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4185
      TabIndex        =   148
      Top             =   5310
      Width           =   1095
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "&Save"
      Height          =   385
      Left            =   5505
      TabIndex        =   149
      Top             =   5310
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   6720
      TabIndex        =   150
      Top             =   5310
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5250
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9260
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Profile"
      TabPicture(0)   =   "cowdat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EPD"
      TabPicture(1)   =   "cowdat.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Breeding/Conception Info"
      TabPicture(2)   =   "cowdat.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fracowweight"
      Tab(2).ControlCount=   1
      Begin VB.Frame fracowweight 
         Height          =   4785
         Left            =   -74925
         TabIndex        =   112
         Top             =   345
         Width           =   7770
         Begin VB.TextBox txtBCComments 
            Height          =   495
            Left            =   1620
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   168
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Frame Frame5 
            Caption         =   "Weaning"
            Height          =   1695
            Left            =   5340
            TabIndex        =   163
            Top             =   3000
            Width           =   2355
            Begin VB.TextBox txtbrdweanwt 
               Height          =   285
               Left            =   1000
               TabIndex        =   65
               Top             =   650
               Width           =   1250
            End
            Begin VB.TextBox txtbrdweancond 
               Height          =   285
               Left            =   1000
               TabIndex        =   66
               Top             =   1000
               Width           =   1250
            End
            Begin VB.TextBox txtbrdweanfat 
               Height          =   285
               Left            =   1000
               TabIndex        =   67
               Top             =   1350
               Width           =   1250
            End
            Begin MSMask.MaskEdBox Dtewean 
               Height          =   285
               Left            =   1005
               TabIndex        =   64
               Top             =   285
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
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   200
               TabIndex        =   167
               Top             =   300
               Width           =   700
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   200
               TabIndex        =   166
               Top             =   650
               Width           =   700
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   200
               TabIndex        =   165
               Top             =   1000
               Width           =   700
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   200
               TabIndex        =   164
               Top             =   1350
               Width           =   700
            End
         End
         Begin VB.CommandButton cmdNewExpDate 
            Caption         =   "&New"
            Height          =   330
            Left            =   2700
            TabIndex        =   161
            Top             =   240
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Frame Frame10 
            Caption         =   "Breeding"
            Height          =   1680
            Left            =   120
            TabIndex        =   125
            Top             =   3015
            Width           =   2535
            Begin VB.TextBox txtbreedfat 
               Height          =   285
               Left            =   1185
               TabIndex        =   59
               Top             =   1335
               Width           =   1230
            End
            Begin VB.TextBox txtbreedcond 
               Height          =   285
               Left            =   1185
               TabIndex        =   58
               Top             =   990
               Width           =   1230
            End
            Begin VB.TextBox txtbreedwt 
               Height          =   285
               Left            =   1185
               TabIndex        =   57
               Top             =   660
               Width           =   1230
            End
            Begin MSMask.MaskEdBox dtebreed 
               Height          =   285
               Left            =   1185
               TabIndex        =   56
               Top             =   330
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
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   180
               TabIndex        =   129
               Top             =   1335
               Width           =   900
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   180
               TabIndex        =   127
               Top             =   990
               Width           =   900
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   180
               TabIndex        =   128
               Top             =   675
               Width           =   900
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   200
               TabIndex        =   126
               Top             =   360
               Width           =   900
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Status"
            Height          =   870
            Left            =   150
            TabIndex        =   122
            Top             =   645
            Width           =   1395
            Begin VB.OptionButton Optopen 
               Caption         =   "Open"
               Height          =   255
               Left            =   180
               TabIndex        =   124
               Top             =   465
               Width           =   840
            End
            Begin VB.OptionButton optpreg 
               Caption         =   "Pregnant"
               Height          =   255
               Left            =   180
               TabIndex        =   123
               Top             =   210
               Width           =   1020
            End
         End
         Begin VB.ComboBox cboyear 
            Height          =   315
            Left            =   1365
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   270
            Width           =   1275
         End
         Begin VB.Frame Frame8 
            Caption         =   "Ultra_Sound"
            Height          =   2025
            Left            =   5115
            TabIndex        =   130
            Top             =   960
            Width           =   2550
            Begin VB.TextBox txtcw 
               Height          =   285
               Left            =   1170
               TabIndex        =   53
               Top             =   570
               Width           =   1250
            End
            Begin VB.TextBox txtbl 
               Height          =   285
               Left            =   1170
               TabIndex        =   54
               Top             =   915
               Width           =   1250
            End
            Begin VB.TextBox txteage 
               Height          =   285
               Left            =   1170
               TabIndex        =   55
               Top             =   1275
               Width           =   1250
            End
            Begin VB.ComboBox cbosex 
               Height          =   315
               Left            =   1170
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   151
               Top             =   1620
               Width           =   615
            End
            Begin MSMask.MaskEdBox Dteconc 
               Height          =   285
               Left            =   1170
               TabIndex        =   52
               Top             =   225
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
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Crown Width"
               Height          =   255
               Left            =   90
               TabIndex        =   160
               Top             =   585
               Width           =   1035
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Body Length"
               Height          =   255
               Left            =   90
               TabIndex        =   159
               Top             =   930
               Width           =   1035
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estimated Age"
               Height          =   255
               Left            =   90
               TabIndex        =   158
               Top             =   1290
               Width           =   1035
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Sex of Calf"
               Height          =   255
               Left            =   165
               TabIndex        =   132
               Top             =   1650
               Width           =   855
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   420
               TabIndex        =   131
               Top             =   255
               Width           =   690
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Calving"
            Height          =   1695
            Left            =   4680
            TabIndex        =   138
            Top             =   195
            Visible         =   0   'False
            Width           =   2400
            Begin VB.TextBox txtbrdcalfat 
               Height          =   285
               Left            =   1035
               TabIndex        =   146
               Top             =   1350
               Width           =   1250
            End
            Begin VB.TextBox txtbrdcalcond 
               Height          =   285
               Left            =   1035
               TabIndex        =   144
               Top             =   1000
               Width           =   1250
            End
            Begin VB.TextBox txtbrdcalwt 
               Height          =   285
               Left            =   1020
               TabIndex        =   142
               Top             =   650
               Width           =   1250
            End
            Begin MSMask.MaskEdBox Dtebrdcal 
               Height          =   285
               Left            =   1020
               TabIndex        =   140
               Top             =   270
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
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   240
               TabIndex        =   145
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   240
               TabIndex        =   143
               Top             =   1000
               Width           =   705
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   240
               TabIndex        =   141
               Top             =   650
               Width           =   705
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   240
               TabIndex        =   139
               Top             =   300
               Width           =   705
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Calving"
            Height          =   1695
            Left            =   2805
            TabIndex        =   133
            Top             =   3000
            Width           =   2355
            Begin VB.TextBox txtbrdtrifat 
               Height          =   285
               Left            =   1035
               TabIndex        =   63
               Top             =   1350
               Width           =   1250
            End
            Begin VB.TextBox txtbrdtricond 
               Height          =   285
               Left            =   1035
               TabIndex        =   62
               Top             =   1000
               Width           =   1250
            End
            Begin VB.TextBox txtbrdtriwt 
               Height          =   285
               Left            =   1035
               TabIndex        =   61
               Top             =   650
               Width           =   1250
            End
            Begin MSMask.MaskEdBox Dtetri 
               Height          =   285
               Left            =   1035
               TabIndex        =   60
               Top             =   285
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
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Backfat"
               Height          =   255
               Left            =   225
               TabIndex        =   137
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condition"
               Height          =   255
               Left            =   240
               TabIndex        =   136
               Top             =   1000
               Width           =   705
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weight"
               Height          =   255
               Left            =   240
               TabIndex        =   135
               Top             =   650
               Width           =   705
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   240
               TabIndex        =   134
               Top             =   300
               Width           =   705
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Breeding"
            Height          =   1290
            Left            =   135
            TabIndex        =   115
            Top             =   1680
            Width           =   4870
            Begin VB.TextBox txtbull3 
               Height          =   285
               Left            =   3300
               TabIndex        =   51
               Top             =   900
               Width           =   1250
            End
            Begin VB.TextBox txtbull2 
               Height          =   285
               Left            =   3300
               TabIndex        =   49
               Top             =   550
               Width           =   1250
            End
            Begin VB.TextBox txtbull1 
               Height          =   285
               Left            =   3300
               TabIndex        =   47
               Top             =   200
               Width           =   1250
            End
            Begin MSMask.MaskEdBox dtebreed1 
               Height          =   285
               Left            =   1170
               TabIndex        =   46
               Top             =   180
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
            Begin MSMask.MaskEdBox dtebreed2 
               Height          =   285
               Left            =   1170
               TabIndex        =   48
               Top             =   525
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
            Begin MSMask.MaskEdBox dtebreed3 
               Height          =   285
               Left            =   1170
               TabIndex        =   50
               Top             =   870
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
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bull ID 3"
               Height          =   255
               Left            =   2600
               TabIndex        =   121
               Top             =   900
               Width           =   605
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bull ID 2"
               Height          =   255
               Left            =   2600
               TabIndex        =   119
               Top             =   550
               Width           =   605
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bull ID 1"
               Height          =   255
               Left            =   2600
               TabIndex        =   117
               Top             =   200
               Width           =   605
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date 3"
               Height          =   255
               Left            =   400
               TabIndex        =   120
               Top             =   900
               Width           =   700
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date 2"
               Height          =   255
               Left            =   400
               TabIndex        =   118
               Top             =   550
               Width           =   700
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date 1"
               Height          =   255
               Left            =   400
               TabIndex        =   116
               Top             =   200
               Width           =   700
            End
         End
         Begin MSMask.MaskEdBox Dteexposed 
            Height          =   285
            Left            =   4620
            TabIndex        =   45
            Top             =   915
            Visible         =   0   'False
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
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Comment"
            Height          =   195
            Left            =   1620
            TabIndex        =   169
            Top             =   960
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lblWarning 
            Alignment       =   2  'Center
            Caption         =   "An asterisk (*) by the date denotes that no breeding/conception records exist with this date."
            Height          =   795
            Left            =   3480
            TabIndex        =   162
            Top             =   300
            Width           =   2475
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Exposed Date"
            Height          =   255
            Left            =   225
            TabIndex        =   157
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Exposure Period"
            Height          =   195
            Left            =   1650
            TabIndex        =   113
            Top             =   750
            Visible         =   0   'False
            Width           =   1170
         End
      End
      Begin VB.Frame Frame3 
         ClipControls    =   0   'False
         Height          =   4725
         Left            =   -74880
         TabIndex        =   89
         Top             =   375
         Width           =   7695
         Begin VB.Frame fraepd 
            Caption         =   "EPD'S"
            Height          =   4245
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   7455
            Begin VB.TextBox txtacc10 
               Height          =   285
               Left            =   6180
               MaxLength       =   10
               TabIndex        =   43
               Top             =   3840
               Width           =   800
            End
            Begin VB.TextBox txtacc9 
               Height          =   285
               Left            =   6180
               MaxLength       =   10
               TabIndex        =   41
               Top             =   3480
               Width           =   800
            End
            Begin VB.TextBox txtacc8 
               Height          =   285
               Left            =   6180
               MaxLength       =   10
               TabIndex        =   39
               Top             =   3180
               Width           =   800
            End
            Begin VB.TextBox txtacc7 
               Height          =   285
               Left            =   6180
               MaxLength       =   10
               TabIndex        =   37
               Top             =   2850
               Width           =   800
            End
            Begin VB.TextBox txtacc6 
               Height          =   285
               Left            =   6180
               MaxLength       =   10
               TabIndex        =   35
               Top             =   2520
               Width           =   800
            End
            Begin VB.TextBox txtacc5 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   33
               Top             =   3840
               Width           =   800
            End
            Begin VB.TextBox txtacc4 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   31
               Top             =   3480
               Width           =   800
            End
            Begin VB.TextBox txtacc3 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   29
               Top             =   3180
               Width           =   800
            End
            Begin VB.TextBox txtacc2 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   27
               Top             =   2850
               Width           =   800
            End
            Begin VB.TextBox txtacc1 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   25
               Top             =   2520
               Width           =   800
            End
            Begin VB.TextBox txtaccmatmilk 
               Height          =   285
               Left            =   4245
               MaxLength       =   10
               TabIndex        =   23
               Top             =   1830
               Width           =   800
            End
            Begin VB.TextBox txtaccmatww 
               Height          =   285
               Left            =   4245
               MaxLength       =   10
               TabIndex        =   21
               Top             =   1500
               Width           =   800
            End
            Begin VB.TextBox txtaccywt 
               Height          =   285
               Left            =   4245
               MaxLength       =   10
               TabIndex        =   19
               Top             =   1170
               Width           =   800
            End
            Begin VB.TextBox txtaccww 
               Height          =   285
               Left            =   4245
               MaxLength       =   10
               TabIndex        =   17
               Top             =   840
               Width           =   800
            End
            Begin VB.TextBox txtaccbw 
               Height          =   285
               Left            =   4245
               MaxLength       =   10
               TabIndex        =   15
               Top             =   480
               Width           =   800
            End
            Begin VB.TextBox txtmisc10 
               Height          =   285
               Left            =   4725
               MaxLength       =   10
               TabIndex        =   42
               Top             =   3840
               Width           =   1000
            End
            Begin VB.TextBox txtmisc9 
               Height          =   285
               Left            =   4725
               MaxLength       =   10
               TabIndex        =   40
               Top             =   3510
               Width           =   1000
            End
            Begin VB.TextBox txtmisc8 
               Height          =   285
               Left            =   4725
               MaxLength       =   10
               TabIndex        =   38
               Top             =   3180
               Width           =   1000
            End
            Begin VB.TextBox txtmisc7 
               Height          =   285
               Left            =   4725
               MaxLength       =   10
               TabIndex        =   36
               Top             =   2850
               Width           =   1000
            End
            Begin VB.TextBox txtmisc6 
               Height          =   285
               Left            =   4725
               MaxLength       =   10
               TabIndex        =   34
               Top             =   2520
               Width           =   1000
            End
            Begin VB.TextBox txtmatww 
               Height          =   285
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   20
               Top             =   1500
               Width           =   1000
            End
            Begin VB.TextBox txtyearwt 
               Height          =   285
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   18
               Top             =   1170
               Width           =   1000
            End
            Begin VB.TextBox txtweanwt 
               Height          =   285
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   16
               Top             =   840
               Width           =   1000
            End
            Begin VB.TextBox txtmisc5 
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   32
               Top             =   3840
               Width           =   1000
            End
            Begin VB.TextBox txtmisc4 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   30
               Top             =   3510
               Width           =   1000
            End
            Begin VB.TextBox Txtmisc3 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   28
               Top             =   3180
               Width           =   1000
            End
            Begin VB.TextBox txtmisc2 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   26
               Top             =   2850
               Width           =   1000
            End
            Begin VB.TextBox txtmisc1 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   24
               Top             =   2520
               Width           =   1000
            End
            Begin VB.TextBox txtmatmilk 
               Height          =   285
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   22
               Top             =   1830
               Width           =   1000
            End
            Begin VB.TextBox txtbirthwt 
               Height          =   285
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   14
               Top             =   480
               Width           =   1000
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   6075
               TabIndex        =   106
               Top             =   2280
               Width           =   1005
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   4800
               TabIndex        =   105
               Top             =   2280
               Width           =   1005
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   2520
               TabIndex        =   99
               Top             =   2280
               Width           =   1000
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   1320
               TabIndex        =   98
               Top             =   2280
               Width           =   1000
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EPD"
               Height          =   255
               Left            =   3000
               TabIndex        =   91
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Accuracy"
               Height          =   255
               Left            =   4035
               TabIndex        =   92
               Top             =   195
               Width           =   975
            End
            Begin VB.Label lblepd10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 10"
               Height          =   255
               Left            =   3780
               TabIndex        =   111
               Top             =   3840
               Width           =   855
            End
            Begin VB.Label lblepd9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 9"
               Height          =   255
               Left            =   3780
               TabIndex        =   110
               Top             =   3510
               Width           =   855
            End
            Begin VB.Label lblepd8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 8"
               Height          =   255
               Left            =   3780
               TabIndex        =   109
               Top             =   3180
               Width           =   855
            End
            Begin VB.Label lblepd7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 7"
               Height          =   255
               Left            =   3780
               TabIndex        =   108
               Top             =   2850
               Width           =   855
            End
            Begin VB.Label lblepd6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 6"
               Height          =   255
               Left            =   3780
               TabIndex        =   107
               Top             =   2520
               Width           =   855
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Maternal"
               Height          =   255
               Left            =   1560
               TabIndex        =   96
               Top             =   1500
               Width           =   1230
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Yearling Wt"
               Height          =   255
               Left            =   1785
               TabIndex        =   95
               Top             =   1170
               Width           =   1005
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Wean Wt"
               Height          =   255
               Left            =   1785
               TabIndex        =   94
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label lblepdmisc5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 5"
               Height          =   255
               Left            =   105
               TabIndex        =   104
               Top             =   3840
               Width           =   1005
            End
            Begin VB.Label lblepdmisc4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 4"
               Height          =   240
               Left            =   105
               TabIndex        =   103
               Top             =   3510
               Width           =   1005
            End
            Begin VB.Label lblepdmisc3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 3"
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   3180
               Width           =   1005
            End
            Begin VB.Label lblepdmisc2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 2"
               Height          =   255
               Left            =   105
               TabIndex        =   101
               Top             =   2850
               Width           =   1005
            End
            Begin VB.Label lblepdmisc1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Misc 1"
               Height          =   240
               Left            =   120
               TabIndex        =   100
               Top             =   2520
               Width           =   1005
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Maternal Milk"
               Height          =   240
               Left            =   1785
               TabIndex        =   97
               Top             =   1830
               Width           =   1005
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Wt"
               Height          =   270
               Left            =   1785
               TabIndex        =   93
               Top             =   480
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame2 
         ClipControls    =   0   'False
         Height          =   4770
         Left            =   105
         TabIndex        =   44
         Top             =   375
         Width           =   7695
         Begin VB.ComboBox CBOReasonCode 
            Height          =   315
            Left            =   3885
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   3690
            Width           =   615
         End
         Begin VB.Frame frampda 
            Caption         =   "Calculated Data"
            Height          =   1005
            Left            =   5280
            TabIndex        =   154
            Top             =   3315
            Width           =   1410
            Begin VB.TextBox txtmpda 
               Enabled         =   0   'False
               Height          =   285
               Left            =   480
               TabIndex        =   155
               Top             =   480
               Width           =   795
            End
            Begin VB.Label lblmpda 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "MPPA"
               Enabled         =   0   'False
               Height          =   225
               Left            =   195
               TabIndex        =   156
               Top             =   270
               Width           =   495
            End
         End
         Begin VB.CheckBox chkedit 
            Caption         =   "Edit"
            Height          =   195
            Left            =   6870
            TabIndex        =   153
            Top             =   4080
            Width           =   765
         End
         Begin VB.CommandButton cmdapply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   6855
            TabIndex        =   152
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox TXTREGNAME 
            Height          =   285
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   3
            Top             =   1605
            Width           =   2325
         End
         Begin VB.TextBox txtproFnotes 
            Height          =   855
            Left            =   5400
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtbirthid 
            Height          =   285
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   5
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtdam 
            Height          =   285
            Left            =   5400
            MaxLength       =   8
            TabIndex        =   10
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtsire 
            Height          =   285
            Left            =   5400
            MaxLength       =   8
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtbreed 
            Height          =   285
            Left            =   1275
            MaxLength       =   8
            TabIndex        =   1
            Top             =   750
            Width           =   1035
         End
         Begin VB.TextBox txtcomments 
            Height          =   285
            Left            =   3885
            MaxLength       =   25
            TabIndex        =   7
            Top             =   4080
            Width           =   1275
         End
         Begin VB.Frame frasource 
            Caption         =   "Source"
            Enabled         =   0   'False
            Height          =   855
            Left            =   1275
            TabIndex        =   78
            Top             =   2535
            Width           =   1335
            Begin VB.OptionButton optpurchased 
               Caption         =   "Purchased"
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton optraised 
               Caption         =   "Raised"
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame frame1 
            Height          =   1095
            Left            =   1260
            TabIndex        =   81
            Top             =   3495
            Width           =   1335
            Begin VB.OptionButton optpedigree 
               Caption         =   "Pedigree"
               Height          =   195
               Left            =   120
               TabIndex        =   84
               Top             =   750
               Width           =   1095
            End
            Begin VB.OptionButton optculled 
               Caption         =   "Culled"
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   495
               Width           =   975
            End
            Begin VB.OptionButton optactive 
               Caption         =   "Active"
               Height          =   195
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txteleid 
            Height          =   285
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   4
            Top             =   2055
            Width           =   2325
         End
         Begin VB.TextBox TxtREGID 
            Height          =   285
            Left            =   1275
            MaxLength       =   20
            TabIndex        =   2
            Top             =   1170
            Width           =   2325
         End
         Begin VB.TextBox txtid 
            Height          =   285
            Left            =   1275
            MaxLength       =   8
            TabIndex        =   0
            Top             =   315
            Width           =   1050
         End
         Begin MSMask.MaskEdBox Mhbirthdate 
            Height          =   285
            Left            =   5400
            TabIndex        =   8
            Top             =   285
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
         Begin MSMask.MaskEdBox dteentered 
            Height          =   285
            Left            =   5400
            TabIndex        =   11
            Top             =   1785
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
         Begin MSMask.MaskEdBox Dteculled 
            Height          =   285
            Left            =   3885
            TabIndex        =   6
            Top             =   3330
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
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Premise ID#"
            Height          =   255
            Left            =   105
            TabIndex        =   71
            Top             =   1650
            Width           =   1140
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Profile Notes"
            Height          =   255
            Left            =   4200
            TabIndex        =   88
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Entered Herd"
            Height          =   255
            Left            =   3840
            TabIndex        =   76
            Top             =   1815
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calf ID at Birth"
            Height          =   255
            Left            =   2880
            TabIndex        =   77
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dam of Cow"
            Height          =   255
            Left            =   4320
            TabIndex        =   75
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sire of Cow"
            Height          =   255
            Left            =   4440
            TabIndex        =   74
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblreasoncode 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Reason Code"
            Height          =   255
            Left            =   2760
            TabIndex        =   86
            Top             =   3720
            Width           =   1005
         End
         Begin VB.Label lblcomments 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Height          =   255
            Left            =   2760
            TabIndex        =   87
            Top             =   4080
            Width           =   1005
         End
         Begin VB.Label lbldateculled 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Culled"
            Height          =   255
            Left            =   2760
            TabIndex        =   85
            Top             =   3360
            Width           =   1005
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Electronic ID"
            Height          =   270
            Left            =   330
            TabIndex        =   72
            Top             =   2070
            Width           =   1050
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registration #"
            Height          =   255
            Left            =   90
            TabIndex        =   70
            Top             =   1215
            Width           =   1140
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow Breed"
            Height          =   270
            Left            =   405
            TabIndex        =   69
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow Birth Date"
            Height          =   270
            Left            =   4080
            TabIndex        =   73
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow ID"
            Height          =   255
            Left            =   405
            TabIndex        =   68
            Top             =   315
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmcow_data"
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
Dim iRet%

Private Sub Init_Information()
 Call init_form(Me) ' Clear Text Boxes
 frasource.Enabled = True
 
 ' load all combo boxes
 cbosex.AddItem "0"
 cbosex.AddItem "1"
 cbosex.AddItem "2"
 cbosex.AddItem "3"
 'cbosex.AddItem " "
 'CBOReasonCode.AddItem " "
 CBOReasonCode.AddItem "G"
 CBOReasonCode.AddItem "H"
 CBOReasonCode.AddItem "J"
 CBOReasonCode.AddItem "K"
 CBOReasonCode.AddItem "L"
 CBOReasonCode.AddItem "R"
 CBOReasonCode.AddItem "Y"
  'set cboyear to current default bull turn out date
 'Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 'Set tbData = DB.OpenRecordset("misc", dbOpenTable)
 '  tbData.Index = "primarykey"
 '
 '  tbData.Seek "=", "TurnDate" & herdid$
 'lblWarning.Visible = False
 'If Not tbData.NoMatch Then
 '  cboyear.AddItem "*" & tbData!thetext, 0
 '  TurnDate = tbData!thetext
lblWarning.Visible = True
'End If
' tbData.Close: Set tbData = Nothing
' DB.Close: Set DB = Nothing
End Sub

Private Sub Load_information()
 Screen.MousePointer = vbHourglass
 Dim SQL$
 Dim DB As database
 Dim TBCOWEPD As Recordset
 Dim tbCowBrd As Recordset
 Call load_year
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set TBCOWEPD = DB.OpenRecordset("COWepd", dbOpenTable)
 TBCOWEPD.Index = "primarykey"
 TBCOWEPD.Seek "=", herdid$, oldid$
 Set tbData = DB.OpenRecordset("COWprof", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 If Not tbData.NoMatch Then
   txtid.TEXT = Field2Str(tbData!CowID)
   Mhbirthdate.TEXT = Field2Date(tbData!birthdate)
   txtbreed.TEXT = Field2Str(tbData!breed)
   TxtREGID.TEXT = Field2Str(tbData!regnum)
   TXTREGNAME.TEXT = Field2Str(tbData!regname)
   txteleid.TEXT = Field2Str(tbData!elecid)
   txtsire.TEXT = Field2Str(tbData!sire)
   txtdam.TEXT = Field2Str(tbData!dam)
   txtproFnotes.TEXT = Field2Str(tbData!notes)
   dteentered.TEXT = Field2Date(tbData!enteredherd)
   txtbirthid.TEXT = Field2Str(tbData!calfid)
   If Field2Str(tbData!Source) = "R" Then optraised.Value = True
   If Field2Str(tbData!Source) = "P" Then optpurchased.Value = True
   If Field2Str(tbData!active) = "A" Then optactive.Value = True
   If Field2Str(tbData!active) = "C" Then optculled.Value = True
   If Field2Str(tbData!active) = "P" Then optpedigree.Value = True
   Dteculled.TEXT = Field2Date(tbData!dateculled)
   Call set_combo(CBOReasonCode, Field2Str(tbData!reasonculled))
   txtcomments.TEXT = Field2Str(tbData!commentsculled)
   txtmpda.TEXT = Field2Str(tbData!mpda)
 End If
 TBCOWEPD.Seek "=", herdid$, oldid$
 If Not TBCOWEPD.NoMatch Then
   txtbirthwt.TEXT = Field2Str(TBCOWEPD!epdbirthwt)
   txtweanwt.TEXT = Field2Num(TBCOWEPD!epdweanwt)
   txtyearwt.TEXT = Field2Str(TBCOWEPD!epdyearwt)
   txtmatww.TEXT = Field2Str(TBCOWEPD!epdmatww)
   txtmatmilk.TEXT = Field2Str(TBCOWEPD!epdmatmilk)
   txtaccbw.TEXT = Field2Str(TBCOWEPD!accbirthwt)
   txtaccww.TEXT = Field2Str(TBCOWEPD!accweanwt)
   txtaccywt.TEXT = Field2Str(TBCOWEPD!accyearwt)
   txtaccmatww.TEXT = Field2Str(TBCOWEPD!accmatww)
   txtaccmatmilk.TEXT = Field2Str(TBCOWEPD!accmatmilk)
   txtmisc1.TEXT = Field2Str(TBCOWEPD!misc1)
   txtmisc2.TEXT = Field2Str(TBCOWEPD!misc2)
   Txtmisc3.TEXT = Field2Str(TBCOWEPD!misc3)
   txtmisc4.TEXT = Field2Str(TBCOWEPD!misc4)
   txtmisc5.TEXT = Field2Str(TBCOWEPD!misc5)
   txtmisc6.TEXT = Field2Str(TBCOWEPD!misc6)
   txtmisc7.TEXT = Field2Str(TBCOWEPD!misc7)
   txtmisc8.TEXT = Field2Str(TBCOWEPD!misc8)
   txtmisc9.TEXT = Field2Str(TBCOWEPD!misc9)
   txtmisc10.TEXT = Field2Str(TBCOWEPD!misc10)
   txtacc1.TEXT = Field2Str(TBCOWEPD!acc1)
   txtacc2.TEXT = Field2Str(TBCOWEPD!acc2)
   txtacc3.TEXT = Field2Str(TBCOWEPD!acc3)
   txtacc4.TEXT = Field2Str(TBCOWEPD!acc4)
   txtacc5.TEXT = Field2Str(TBCOWEPD!acc5)
   txtacc6.TEXT = Field2Str(TBCOWEPD!acc6)
   txtacc7.TEXT = Field2Str(TBCOWEPD!acc7)
   txtacc8.TEXT = Field2Str(TBCOWEPD!acc8)
   txtacc9.TEXT = Field2Str(TBCOWEPD!acc9)
   txtacc10.TEXT = Field2Str(TBCOWEPD!acc10)
 End If
 SQL$ = "select * from cowbrd where herdid = '" & herdid$ & "' and cowid = '" & txtid.TEXT & "'"
 Set tbCowBrd = DB.OpenRecordset(SQL$, dbOpenDynaset)
 If tbCowBrd.RecordCount > 0 Then
   tbCowBrd.MoveLast
   dtebreed1.TEXT = Field2Date(tbCowBrd!breeddate1)
   dtebreed2.TEXT = Field2Date(tbCowBrd!breeddate2)
   dtebreed3.TEXT = Field2Date(tbCowBrd!BREEDDATE3)
   txtbull1.TEXT = Field2Str(tbCowBrd!breedbull1)
   txtbull2.TEXT = Field2Str(tbCowBrd!breedbull2)
   txtbull3.TEXT = Field2Str(tbCowBrd!breedbull3)
   Dteexposed.TEXT = Field2Date(tbCowBrd!exposed)
   dtebreed.TEXT = Field2Date(tbCowBrd!breeddate)
   txtbreedwt.TEXT = Field2Str(tbCowBrd!breedwt)
   txtbreedcond.TEXT = Field2Str(tbCowBrd!breedcond)
   txtbreedfat.TEXT = Field2Str(tbCowBrd!breedbaCKfat)
   Dtewean.TEXT = Field2Date(tbCowBrd!weandate)
   txtbrdweanwt.TEXT = Field2Str(tbCowBrd!weanwt)
   txtbrdweancond.TEXT = Field2Str(tbCowBrd!weancond)
   txtbrdweanfat.TEXT = Field2Str(tbCowBrd!weanbackfat)
   Dtetri.TEXT = Field2Date(tbCowBrd!tridate)
   txtbrdtriwt.TEXT = Field2Str(tbCowBrd!triwt)
   txtbrdtricond.TEXT = Field2Str(tbCowBrd!tricond)
   txtbrdtrifat.TEXT = Field2Str(tbCowBrd!tribackfat)
   Call set_combo(Me!cboyear, Field2Date(tbCowBrd!calfdate))
   If tbCowBrd!stat = "P" Then
     optpreg.Value = True
   Else
     Optopen.Value = True
   End If
   Dteconc.TEXT = Field2Date(tbCowBrd!conceptdate)
   'txtconcbull.TEXT = Field2Str(TBCOWbrd!conceptbull)
   Call set_combo(Me!cbosex, Trim$(Field2Str(tbCowBrd!sexofcalf)))
   Dtebrdcal.TEXT = Field2Date(tbCowBrd!calfdate)
   txtbrdcalwt.TEXT = Field2Str(tbCowBrd!calfwt)
   txtbrdcalcond.TEXT = Field2Str(tbCowBrd!calfcond)
   txtbrdcalfat.TEXT = Field2Str(tbCowBrd!calfbackfat)
   txtcw = Field2Str(tbCowBrd!crownwidth)
   txtbl = Field2Str(tbCowBrd!bodylength)
   txteage = Field2Str(tbCowBrd!age)
   If tbCowBrd!calfdate <> TurnDate Then
      Call set_combo(Me!cboyear, Field2Date(tbCowBrd!calfdate))
   Else
      cboyear.RemoveItem (0)
      Call set_combo(Me!cboyear, Field2Date(tbCowBrd!calfdate))
   End If
   txtBCComments = Field2Str(tbCowBrd!Comments)
 End If
 tbData.Close: Set tbData = Nothing
 TBCOWEPD.Close: Set TBCOWEPD = Nothing
 tbCowBrd.Close: Set tbCowBrd = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub load_year()
 Screen.MousePointer = vbHourglass
 Dim SQL$, TheDate$
 Dim dbtest As database
 Dim tbCowBrd As Recordset
 Set dbtest = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select calfdate, breeddate1, weanwt from cowbrd where herdid = '" & herdid$ & "' and cowid = '" & oldid$ & "' order by  breeddate1 desc "
 Set tbCowBrd = dbtest.OpenRecordset(SQL$, dbOpenDynaset)
 Do Until tbCowBrd.EOF
   TheDate = tbCowBrd!calfdate & IIf(IsNull(tbCowBrd!breeddate1) And Field2Num(tbCowBrd!weanwt) = 0, "*", "")
   cboyear.AddItem TheDate
   tbCowBrd.MoveNext
 Loop
 tbCowBrd.Close: Set tbCowBrd = Nothing
 dbtest.Close: Set dbtest = Nothing
End Sub

Private Sub save_information()
 Dim Replace$, RESPONSE%
 Dim TableName$(100)
 Dim TBCOWEPD As Recordset
 Dim tbCowBrd As Recordset, TheDate As Date
 
On Local Error GoTo ErrHandler
 
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbData = DB.OpenRecordset("COWprof", dbOpenTable)
 tbData.Index = "primarykey"
 tbData.Seek "=", herdid$, oldid$
 save% = True
  If Not tbData.NoMatch Then
   If addedflag$ = "D" Then
     Call CheckID(dbfile$, "cowprof", oldid$, TableName$())
     RESPONSE% = vbYes
     If Val(TableName$(0)) > 2 Then
       Beep
       RESPONSE% = MsgBox("Warning This Cow Id Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
     End If
     If RESPONSE% = vbYes Then
       tbData.Delete
       GoSub donedel
     End If
     Replace$ = ""
     save% = False
    Else
     tbData.Edit
   End If
  Else
   tbData.AddNew
 End If
 If save% Then ' if in add or edit mode save the information on the form
   With tbData
10       !herdid = herdid$
11       !CowID = txtid.TEXT
12       Call Date2Field(!birthdate, Mhbirthdate.TEXT)
'       !birthdate = THEDATE
13       !breed = txtbreed.TEXT
14       !regnum = TxtREGID.TEXT
15       !regname = TXTREGNAME.TEXT
16       !elecid = txteleid.TEXT
17       !sire = txtsire.TEXT
18       !dam = txtdam.TEXT
19       !notes = txtproFnotes.TEXT
20       Call Date2Field(!enteredherd, dteentered.TEXT)
'       !enteredherd = THEDATE
21       !calfid = txtbirthid.TEXT
22       If optraised Then !Source = "R"
23       If optpurchased Then !Source = "P"
24       If optactive Then !active = "A"
25       If optculled Then !active = "C"
26       If optpedigree Then !active = "P"
27       Call Date2Field(!dateculled, Dteculled.TEXT)
       '!dateculled = THEDATE
28       !reasonculled = CBOReasonCode.TEXT
29       !commentsculled = txtcomments.TEXT
30       !mpda = Val(txtmpda.TEXT)
       .Update
       Replace$ = txtid.TEXT & vbTab & txtbreed.TEXT & vbTab & Mhbirthdate.TEXT
   End With
 Set TBCOWEPD = DB.OpenRecordset("COWepd", dbOpenTable)
 TBCOWEPD.Index = "primarykey"
 TBCOWEPD.Seek "=", herdid$, txtid.TEXT
   If Not TBCOWEPD.NoMatch Then
   If addedflag$ = "D" Then
     TBCOWEPD.Delete
     Replace$ = ""
     save% = False
    Else
     TBCOWEPD.Edit
   End If
  Else
   TBCOWEPD.AddNew
 End If
   With TBCOWEPD
31       !herdid = herdid$
32       !CowID = txtid.TEXT
33       !epdbirthwt = Val(txtbirthwt.TEXT)
34       !epdweanwt = Val(txtweanwt.TEXT)
35       !epdyearwt = Val(txtyearwt.TEXT)
36       !epdmatww = Val(txtmatww.TEXT)
37       !epdmatmilk = Val(txtmatmilk.TEXT)
38       !accbirthwt = Val(txtaccbw.TEXT)
39       !accweanwt = Val(txtaccww.TEXT)
40       !accyearwt = Val(txtaccywt.TEXT)
41       !accmatww = Val(txtaccmatww.TEXT)
42       !accmatmilk = Val(txtaccmatmilk.TEXT)
43       !misc1 = txtmisc1.TEXT
44       !misc2 = txtmisc2.TEXT
45       !misc3 = Txtmisc3.TEXT
46       !misc4 = txtmisc4.TEXT
47       !misc5 = txtmisc5.TEXT
48       !misc5 = txtmisc5.TEXT
49       !misc6 = txtmisc6.TEXT
50       !misc7 = txtmisc7.TEXT
51       !misc8 = txtmisc8.TEXT
52       !misc9 = txtmisc9.TEXT
53       !misc10 = txtmisc10.TEXT
54       !acc1 = txtacc1.TEXT
55       !acc2 = txtacc2.TEXT
56       !acc3 = txtacc3.TEXT
57       !acc4 = txtacc4.TEXT
58       !acc5 = txtacc5.TEXT
59       !acc5 = txtacc5.TEXT
60       !acc6 = txtacc6.TEXT
61       !acc7 = txtacc7.TEXT
62       !acc8 = txtacc8.TEXT
63       !acc9 = txtacc9.TEXT
64       !acc10 = txtacc10.TEXT
       .Update
   End With
 End If
 If addedflag$ = "D" Then GoSub donedel
 If iRet = 1 Then GoTo Skip_Cow_Breed_Tab 'not a valid year, skip this section
Set tbCowBrd = DB.OpenRecordset("COWbrd", dbOpenTable)
tbCowBrd.Index = "primarykey"
 If cboyear.TEXT <> "" Then
 If Right(cboyear.TEXT, 1) = "*" Then TheDate = Left(cboyear.TEXT, 10) Else TheDate = cboyear.TEXT
 If IsDate(TheDate) Then
   tbCowBrd.Seek "=", herdid$, txtid.TEXT, TheDate
   If Not tbCowBrd.NoMatch Then
      tbCowBrd.Edit
   Else
      tbCowBrd.AddNew
   End If
   With tbCowBrd
65      !herdid = herdid$
66      !CowID = txtid.TEXT
      '!calfdate = THEDATE
67      Call Date2Field(!breeddate1, dtebreed1.TEXT)
68      '!breeddate1 = THEDATE
69      Call Date2Field(!breeddate2, dtebreed2.TEXT)
70      '!breeddate2 = THEDATE
71      Call Date2Field(!BREEDDATE3, dtebreed3.TEXT)
72      '!breeddate3 = THEDATE
73      !breedbull1 = txtbull1.TEXT
74      !breedbull2 = txtbull2.TEXT
75      !breedbull3 = txtbull3.TEXT
76      Call Date2Field(!exposed, Dteexposed.TEXT)
'      !exposed = THEDATE
77      Call Date2Field(!breeddate, dtebreed.TEXT)
'      !breeddate = THEDATE
78      !breedwt = Val(txtbreedwt.TEXT)
79      !breedcond = Val(txtbreedcond.TEXT)
80      !breedbaCKfat = Val(txtbreedfat.TEXT)
81      Call Date2Field(!weandate, Dtewean.TEXT)
'      !weandate = THEDATE
82      !weanwt = Val(txtbrdweanwt.TEXT)
83      !weancond = Val(txtbrdweancond.TEXT)
84      !weanbackfat = Val(txtbrdweanfat.TEXT)
85      Call Date2Field(!tridate, Dtetri.TEXT)
      '!tridate = THEDATE
86      !triwt = Val(txtbrdtriwt.TEXT)
87      !tricond = Val(txtbrdtricond.TEXT)
88      !tribackfat = Val(txtbrdtrifat.TEXT)
'      !Year = cboyear.TEXT
      If optpreg.Value = True Then
89        tbCowBrd!stat = "P"
      Else
90        tbCowBrd!stat = "O"
      End If
91      Call Date2Field(!conceptdate, Dteconc.TEXT)
      '!conceptdate = THEDATE
      '!conceptbull = txtconcbull.TEXT
92      !sexofcalf = cbosex.TEXT
93      !calfwt = Val(txtbrdcalwt.TEXT)
94      !calfcond = Val(txtbrdcalcond.TEXT)
95      !calfbackfat = Val(txtbrdcalfat.TEXT)
96      !crownwidth = Val(txtcw.TEXT)
97      !bodylength = Val(txtbl.TEXT)
98      !age = Val(txteage.TEXT)
99      !Comments = txtBCComments
      .Update
   End With
 End If
 End If
 tbCowBrd.Close: Set tbCowBrd = Nothing
Skip_Cow_Breed_Tab:
 TBCOWEPD.Close: Set TBCOWEPD = Nothing
 
donedel:
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 dirtyflag% = False
 Call Update_mh_ListBoxes("lstCOW", 0, oldid$, Replace$)
 Screen.MousePointer = vbDefault
Exit Sub
ErrHandler:
TEXT(2) = Erl
GMODNAME$ = Me.Name & " - save_information"
GERRNUM$ = Str$(Err.Number)
GERRSOURCE$ = Err.Source
Call POP_ERROR(TEXT$())
End Sub



Private Sub valid_form(exitcode%)
    Dim RESPONSE%
    exitcode% = 0
    iRet = 0 'holds status of cboyear
    If herdid = "" Then
        Beep
        MsgBox "Please Select A Herd", vbOKOnly + vbCritical, Me.Caption
        selherd_List.cmdcancel.Visible = False
        selherd_List.Show vbModal
    End If
    'If Not IsDate(cboyear.TEXT) Then
    '    Beep
    '    RESPONSE% = MsgBox("Warning: Without an Exposed Date the Breeding tab can not be saved." & vbCrLf & "Do you wish to save anyway?", vbYesNo + vbQuestion, Me.Caption)
    '    iRet = 1
    '    If RESPONSE% = vbNo Then
    '      SSTab1.Tab = 2
    '      cboyear.SetFocus
    '      exitcode% = 1
    '      Exit Sub
    '    End If
    'End If
    If txtid.TEXT = "" Then
        Beep
        MsgBox "Cow ID Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
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
        Set tbData = DB.OpenRecordset("cowPROF", dbOpenTable)
        tbData.Index = "primarykey"
        tbData.Seek "=", herdid$, txtid.TEXT
        If Not tbData.NoMatch Then
            Beep
            MsgBox "Cow ID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
            exitcode% = 1
            tbData.Close: Set tbData = Nothing
            DB.Close: Set DB = Nothing
            txtid.SetFocus
            Exit Sub
        End If
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
    End If
    Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
    Set tbData = DB.OpenRecordset("sirePROF", dbOpenTable)
    tbData.Index = "primarykey"
    tbData.Seek "=", herdid$, txtsire.TEXT
    If tbData.NoMatch Then
       Beep
       MsgBox "Must have a valid Sire ID", vbOKOnly + vbCritical, Me.Caption
       SSTab1.Tab = 0
       txtsire.SetFocus
       exitcode% = 1
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
       txtsire.SetFocus
       Exit Sub
    End If
    tbData.Close: Set tbData = Nothing
    DB.Close: Set DB = Nothing
   If optculled = True Then
      If Dteculled.TEXT = "--/--/----" Then
        MsgBox "Date Culled Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
        SSTab1.Tab = 0
        Dteculled.SetFocus
        exitcode% = 1
        Exit Sub
      End If
      If CBOReasonCode.TEXT = "" Then
        MsgBox "Reason Culled Must Be Filled Out", vbOKOnly + vbCritical, Me.Caption
        SSTab1.Tab = 0
        CBOReasonCode.SetFocus
        exitcode% = 1
        Exit Sub
    End If
   End If
   
   If Check_EID(txteleid.TEXT, "Cow", oldid, herdid, txtbirthid.TEXT, "Cow") = False Then
       Beep
       MsgBox "EID Can Not Be Duplicated", vbOKOnly + vbCritical, Me.Caption
       exitcode% = 1
    End If
End Sub

Private Sub cboyear_Click()
 On Error GoTo ehandle
 Dim SQL$, iResponse As Integer
 Dim tbCowBrd As Recordset, TheDate As String
 If Right(cboyear.TEXT, 1) <> "*" Then TheDate = cboyear.TEXT Else TheDate = Left(cboyear.TEXT, 10)
 If Left$(cboyear.TEXT, 1) = "*" Then lblWarning.Visible = True Else lblWarning.Visible = False
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL$ = "select * from cowbrd where herdid = '" & herdid$ & "' and cowid = '" & txtid.TEXT & "' and calfdate = #" & TheDate & "#"
' MsgBox SQL$
 Set tbCowBrd = DB.OpenRecordset(SQL$, dbOpenDynaset)
 If tbCowBrd.RecordCount > 0 Then
   dtebreed1.TEXT = Field2Date(tbCowBrd!breeddate1)
   dtebreed2.TEXT = Field2Date(tbCowBrd!breeddate2)
   dtebreed3.TEXT = Field2Date(tbCowBrd!BREEDDATE3)
   txtbull1.TEXT = Field2Str(tbCowBrd!breedbull1)
   txtbull2.TEXT = Field2Str(tbCowBrd!breedbull2)
   txtbull3.TEXT = Field2Str(tbCowBrd!breedbull3)
   Dteexposed.TEXT = Field2Date(tbCowBrd!exposed)
   dtebreed.TEXT = Field2Date(tbCowBrd!breeddate)
   txtbreedwt.TEXT = Field2Str(tbCowBrd!breedwt)
   txtbreedcond.TEXT = Field2Str(tbCowBrd!breedcond)
   txtbreedfat.TEXT = Field2Str(tbCowBrd!breedbaCKfat)
   Dtewean.TEXT = Field2Date(tbCowBrd!weandate)
   txtbrdweanwt.TEXT = Field2Str(tbCowBrd!weanwt)
   txtbrdweancond.TEXT = Field2Str(tbCowBrd!weancond)
   txtbrdweanfat.TEXT = Field2Str(tbCowBrd!weanbackfat)
   Dtetri.TEXT = Field2Date(tbCowBrd!tridate)
   txtbrdtriwt.TEXT = Field2Str(tbCowBrd!triwt)
   txtbrdtricond.TEXT = Field2Str(tbCowBrd!tricond)
   txtbrdtrifat.TEXT = Field2Str(tbCowBrd!tribackfat)
   Call set_combo(Me!cboyear, Field2Str(tbCowBrd!Year))
   If tbCowBrd!stat = "P" Then
     optpreg.Value = True
   Else
     Optopen.Value = True
   End If
   Dteconc.TEXT = Field2Date(tbCowBrd!conceptdate)
   'txtconcbull.TEXT = Field2Str(TBCOWbrd!conceptbull)
   'cbosex.TEXT = Field2str(TBCOWbrd!sexofcalf)
   Call set_combo(Me!cbosex, Field2Str(tbCowBrd!sexofcalf))
   Dtebrdcal.TEXT = Field2Date(tbCowBrd!calfdate)
   txtbrdcalwt.TEXT = Field2Str(tbCowBrd!calfwt)
   txtbrdcalcond.TEXT = Field2Str(tbCowBrd!calfcond)
   txtbrdcalfat.TEXT = Field2Str(tbCowBrd!calfbackfat)
   txtBCComments.TEXT = Field2Str(tbCowBrd!Comments)
 Else
   If optactive.Value Then
      iResponse = MsgBox("A New Breeding/Conception Record Is About To Be Created.  Do You Wish To Continue?", vbYesNo + vbCritical, Me.Caption)
      If iResponse = vbYes Then
         tbCowBrd.AddNew
         tbCowBrd!herdid = herdid$
         If txtid.TEXT = "" Then
            MsgBox "Please Enter A Valid Cow ID", vbOKOnly, Me.Caption
            SSTab1.Tab = 0
            txtid.SetFocus
            Exit Sub
         End If
         tbCowBrd!CowID = txtid.TEXT
         tbCowBrd!calfdate = TurnDate
         tbCowBrd.Update
         cboyear.AddItem CStr(TurnDate)
         cboyear.RemoveItem (0)
      End If
   End If
 End If
 tbCowBrd.Close: Set tbCowBrd = Nothing
 DB.Close: Set DB = Nothing
Exit Sub
ehandle:
If Err.Number = 3021 Then
   tbCowBrd.Close: Set tbCowBrd = Nothing
   DB.Close: Set DB = Nothing
   Exit Sub
End If
End Sub

Private Sub cboyear_GotFocus()
SSTab1.Tab = 2
End Sub

Private Sub chkedit_Click()
 If chkedit.Value = vbChecked Then
   lblmpda.Enabled = True
   txtmpda.Enabled = True
 Else
   lblmpda.Enabled = False
   txtmpda.Enabled = False
 End If
End Sub

Private Sub cmdapply_Click()
  Dim DB As database
  Dim tbMPPA As Recordset
  Dim SQL$
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  'SQL$ = "SELECT DISTINCTROW calfwean.HerdID, Count(calfwean.CalfID) AS CountOfCalfID, calfwean.ratio, Sum(calfwean.ratio) AS SumOfratio From calfwean GROUP BY calfwean.HerdID, calfwean.ratio HAVING (((calfwean.HerdID)='" & herdid$ & "') AND ((calfwean.ratio)>0))"
  SQL$ = "SELECT DISTINCTROW calfwean.HerdID, Count(calfwean.CalfID) AS CountOfCalfID, Sum(calfwean.ratio) AS Sum, calfbirth.CowID"
  SQL$ = SQL$ & " FROM calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID)"
  SQL$ = SQL$ & " where (((calfwean.actweight) > 0) And ((calfwean.Ratio) > 0)) and calfwean.herdid = '" & herdid$ & "' and calfbirth.cowid = '" & txtid.TEXT & "'"
  SQL$ = SQL$ & " GROUP BY calfwean.HerdID, calfbirth.CowID"
  Set tbMPPA = DB.OpenRecordset(SQL$, dbOpenDynaset)
  'txtmpda.TEXT = 100 + ((N * 0.4) / (1 + ((N - 1) * 0.4))) * (c - 100)
  If tbMPPA.RecordCount > 0 Then
  With tbMPPA
      txtmpda.TEXT = 100 + ((!countofcalfid * 0.4) / (1 + (!countofcalfid - 1) * 0.4)) * ((!Sum / !countofcalfid) - 100)
  End With
  End If
  tbMPPA.Close: Set tbMPPA = Nothing
  DB.Close: Set DB = Nothing
End Sub

Private Sub CMDCancel_Click()
 Unload Me
End Sub



Private Sub cmdNewExpDate_Click()
frmPopUpDate.Show vbModal
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
 Set tbData = DB.OpenRecordset("select * from cowprof where herdid='" & herdid$ & "'", dbOpenDynaset)
 'tbdata.Index = "primarykey"
 tbData.FindFirst "herdid = '" & herdid$ & "' And cowid = '" & txtid.TEXT & "'"
 tbData.MoveNext
 If tbData.EOF Then
  tbData.MoveFirst
  frmcow_data.Tag = "E/" & tbData!CowID
 Else
  'tbdata.MoveNext
  frmcow_data.Tag = "E/" & tbData!CowID
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
 Set tbData = DB.OpenRecordset("select * from cowprof where herdid='" & herdid$ & "'", dbOpenDynaset)
 'tbdata.Index = "primarykey"
 tbData.FindFirst "herdid = '" & herdid$ & "' And cowid = '" & txtid.TEXT & "'"
 tbData.MovePrevious
 If tbData.BOF Then
  tbData.MoveLast
  frmcow_data.Tag = "E/" & tbData!CowID
 Else
  'tbdata.MoveNext
  frmcow_data.Tag = "E/" & tbData!CowID
 End If
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
 Call Form_Activate
End Sub


Private Sub CmdSave_Click()
 Dim exitcode%, RESPONSE%
 Dim iRet%
 Dim TableName$(100)
 
 If gIsDemo Then
   If IsValidCowEntry = False Then MsgBox "The demo version of this software only allows ten cow records.", vbOKOnly, "C.H.A.P.S. Demo": Exit Sub
 End If
  
 If addedflag$ <> "D" Then
   Call valid_form(exitcode%)
   If exitcode% = 1 Then Exit Sub
 End If
 If addedflag$ = "D" Then
   Call CheckID(dbfile$, "cowprof", oldid$, TableName$())
   RESPONSE% = vbYes
   If Val(TableName$(0)) > 3 Then
     Beep
     RESPONSE% = MsgBox("Warning This Cow Id Is Referenced By Other Files. Deleting Would Also Delete That Data also." & vbCrLf & " Do You Wish To Delete Anyway?", vbYesNo + vbQuestion, Me.Caption)
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



Private Sub dtebreed1_GotFocus()
   SSTab1.Tab = 2
End Sub

Private Sub Form_Activate()
 If Me.Tag = "" Then Exit Sub
 addedflag$ = Left$(Me.Tag, 1)
 Me.Caption = "Cow Information" & " for Herd " & herdid$
 Screen.MousePointer = vbHourglass
 Call Init_Information
 optraised.Value = True
 optactive.Value = True
 lblepdmisc1 = epdhead1$
 lblepdmisc2 = epdhead2$
 lblepdmisc3 = epdhead3$
 lblepdmisc4 = epdhead4$
 lblepdmisc5 = epdhead5$
 lblepd6 = epdhead6$
 lblepd7 = epdhead7$
 lblepd8 = epdhead8$
 lblepd9 = epdhead9$
 lblepd10 = epdhead10$
 If addedflag$ = "A" Then
    cmdnext.Enabled = False
    cmdprev.Enabled = False
   'Me.caption = "Add"
    oldid$ = ""
    'If optactive.Value = False And optculled.Value = False Then optactive.Value = True
    'If optculled.Value = False Then
    '   CBOReasonCode.Enabled = False
    '   txtcomments.Enabled = False
    '   Dteculled.Enabled = False
    '  Else
    '   CBOReasonCode.Enabled = True
    '   txtcomments.Enabled = True
    '   Dteculled.Enabled = True
    End If
 'End If
 If addedflag$ = "E" Or addedflag$ = "D" Then
   oldid$ = Trim$(Mid$(Me.Tag, 3))
   'Me.caption = "Edit"
   Me.Caption = "Cow Information" & " for Herd " & herdid$ & " - Cow " & oldid$

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
 'Me.Enabled = True
  If addedflag$ <> "D" Then txtid.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 dirtyflag% = True
End Sub


Private Sub Form_Load()
 Call centermdiform(Me, mdimain, 0, 0)
 Call AddCustomToolTip(CBOReasonCode, "G=Died" & vbCrLf & "H=Age" & vbCrLf & "J=Physical Defect" & vbCrLf & "K=Open" & vbCrLf & "L=Inferior Calves" & vbCrLf & "R=Replacement" & vbCrLf & "Y=Unknown reason", Me)
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
 Set frmcow_data = Nothing
End Sub
















Private Sub optactive_Click()
 If optactive.Value = True Then
  Dteculled.TEXT = "--/--/----"
  CBOReasonCode.ListIndex = -1
  Dteculled.Enabled = False
  CBOReasonCode.Enabled = False
  txtcomments.Enabled = False
  lblreasoncode.Enabled = False
  lbldateculled.Enabled = False
  lblcomments.Enabled = False
End If
 End Sub

Private Sub optculled_Click()
If optculled.Value = False Then
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
If optculled.Value = True Then
   Dteculled.Enabled = True
   CBOReasonCode.Enabled = True
   txtcomments.Enabled = True
   lblreasoncode.Enabled = True
   lbldateculled.Enabled = True
   lblcomments.Enabled = True
   MsgBox "Pedigree option is to record parental performance for sire or dams that was not raised in this operation", vbOKOnly + vbInformation, Me.Caption
Else
   Dteculled.TEXT = "--/--/----"
   CBOReasonCode.ListIndex = -1
   Dteculled.Enabled = False
   CBOReasonCode.Enabled = False
   txtcomments.Enabled = False
   lblreasoncode.Enabled = False
   lbldateculled.Enabled = False
   lblcomments.Enabled = False
 End If
End Sub

Private Sub optpurchased_Click()
Label10.Enabled = False
txtbirthid.Enabled = False
End Sub

Private Sub optraised_Click()
Label10.Enabled = True
txtbirthid.Enabled = True
End Sub


Private Sub SSTab1_GotFocus()
  If addedflag$ = "D" Then Exit Sub
  If SSTab1.Tab = 0 Then txtid.SetFocus
  If SSTab1.Tab = 1 Then
    txtbirthwt.SetFocus
    lblepdmisc1 = epdhead1$
    lblepdmisc2 = epdhead2$
    lblepdmisc3 = epdhead3$
    lblepdmisc4 = epdhead4$
    lblepdmisc5 = epdhead5$
    lblepd6 = epdhead6$
    lblepd7 = epdhead7$
    lblepd8 = epdhead8$
    lblepd9 = epdhead9$
    lblepd10 = epdhead10$
  End If
  If SSTab1.Tab = 2 Then dtebreed1.SetFocus
  Me.Caption = "Cow Information" & " for Herd " & herdid$ & " - Cow " & txtid.TEXT

End Sub


Private Sub txtbirthid_DblClick()
Dim tbCalfData As Recordset, pDB As database, pSQL$
On Local Error GoTo ErrHandler
 selcalf_list.SetSex = 2
 selcalf_list.Show vbModal
 If selcalf_list.Tag = "CANCEL" Then Exit Sub
 txtbirthid.TEXT = selcalf_list.Tag
Set pDB = DBEngine(0).OpenDatabase(dbfile$, False, False)
pSQL = "SELECT DISTINCTROW calfbirth.breed, calfbirth.registration, calfbirth.regname, calfbirth.elecid, calfbirth.birthdate, " & _
    "calfbirth.sireID, calfbirth.cowid, calfbirth.notes FROM sireprof INNER JOIN calfbirth ON (calfbirth.HerdID = sireprof.HerdID) " & _
    "AND (sireprof.SireID = calfbirth.sireID) where calfbirth.herdid = '" & herdid & "' and calfbirth.calfid = '" & txtbirthid.TEXT & "'"
    
Set tbCalfData = pDB.OpenRecordset(pSQL, dbOpenSnapshot)
If tbCalfData.EOF Then GoTo Close_DB
With tbCalfData
10    txtbreed = Field2Str(!breed)
20    TxtREGID = Field2Str(!registration)
30    TXTREGNAME = Field2Str(!regname)
40    txteleid = Field2Str(!elecid)
50    Mhbirthdate = Field2Date(!birthdate)
60    txtsire = Field2Str(!sireid)
70    txtdam = Field2Str(!CowID)
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


Private Sub txtbl_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbl_LostFocus()
If Val(txtbl.TEXT) < 0 Or Val(txtbl.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 2
   txtbl.SetFocus
   txtbl.SelStart = 0
   txtbl.SelLength = Len(txtbl.TEXT)
End If
End Sub


Private Sub txtbrdtricond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdtricond_LostFocus()
If Val(txtbrdtricond.TEXT) < 0 Or Val(txtbrdtricond.TEXT) > 10 Then
   MsgBox "Please enter a number between 0 and 10", vbOKOnly
   SSTab1.Tab = 2
   txtbrdtricond.SetFocus
   txtbrdtricond.SelStart = 0
   txtbrdtricond.SelLength = Len(txtbrdtricond.TEXT)
End If
End Sub


Private Sub txtbrdtrifat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdtrifat_LostFocus()
If Val(txtbrdtrifat.TEXT) < 0 Or Val(txtbrdtrifat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 2
   txtbrdtrifat.SetFocus
   txtbrdtrifat.SelStart = 0
   txtbrdtrifat.SelLength = Len(txtbrdtrifat.TEXT)
End If
End Sub


Private Sub txtbrdtriwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdtriwt_LostFocus()
If Val(txtbrdtriwt.TEXT) < 0 Or Val(txtbrdtriwt.TEXT) > 4000 Then
   MsgBox "Please enter a number between 0 and 4000", vbOKOnly
   SSTab1.Tab = 2
   txtbrdtriwt.SetFocus
   txtbrdtriwt.SelStart = 0
   txtbrdtriwt.SelLength = Len(txtbrdtriwt.TEXT)
End If
End Sub


Private Sub txtbrdweancond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdweancond_LostFocus()
If Val(txtbrdweancond.TEXT) < 0 Or Val(txtbrdweancond.TEXT) > 10 Then
   MsgBox "Please enter a number between 0 and 10", vbOKOnly
   SSTab1.Tab = 2
   txtbrdweancond.SetFocus
   txtbrdweancond.SelStart = 0
   txtbrdweancond.SelLength = Len(txtbrdweancond.TEXT)
End If
End Sub


Private Sub txtbrdweanfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdweanfat_LostFocus()
If Val(txtbrdweanfat.TEXT) < 0 Or Val(txtbrdweanfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 2
   txtbrdweanfat.SetFocus
   txtbrdweanfat.SelStart = 0
   txtbrdweanfat.SelLength = Len(txtbrdweanfat.TEXT)
End If
End Sub


Private Sub txtbrdweanwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbrdweanwt_LostFocus()
If Val(txtbrdweanwt.TEXT) < 0 Or Val(txtbrdweanwt.TEXT) > 4000 Then
   MsgBox "Please enter a number between 0 and 4000", vbOKOnly
   SSTab1.Tab = 2
   txtbrdweanwt.SetFocus
   txtbrdweanwt.SelStart = 0
   txtbrdweanwt.SelLength = Len(txtbrdweanwt.TEXT)
End If
End Sub


Private Sub txtbreedcond_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbreedcond_LostFocus()
If Val(txtbreedcond.TEXT) < 0 Or Val(txtbreedcond.TEXT) > 10 Then
   MsgBox "Please enter a number between 0 and 10", vbOKOnly
   SSTab1.Tab = 2
   txtbreedcond.SetFocus
   txtbreedcond.SelStart = 0
   txtbreedcond.SelLength = Len(txtbreedcond.TEXT)
End If
End Sub


Private Sub txtbreedfat_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbreedfat_LostFocus()
If Val(txtbreedfat.TEXT) < 0 Or Val(txtbreedfat.TEXT) > 5 Then
   MsgBox "Please enter a number between 0 and 5", vbOKOnly
   SSTab1.Tab = 2
   txtbreedfat.SetFocus
   txtbreedfat.SelStart = 0
   txtbreedfat.SelLength = Len(txtbreedfat.TEXT)
End If
End Sub


Private Sub txtbreedwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtbreedwt_LostFocus()
If Val(txtbreedwt.TEXT) < 0 Or Val(txtbreedwt.TEXT) > 4000 Then
   MsgBox "Please enter a number between 0 and 4000", vbOKOnly
   SSTab1.Tab = 2
   txtbreedwt.SetFocus
   txtbreedwt.SelStart = 0
   txtbreedwt.SelLength = Len(txtbreedwt.TEXT)
End If
End Sub


Private Sub txtbull1_DblClick()
 selsire_list.Show vbModal
 If selsire_list.Tag = "CANCEL" Then Exit Sub
 txtbull1.TEXT = selsire_list.Tag
End Sub


Private Sub txtbull2_DblClick()
 selsire_list.Show vbModal
 If selsire_list.Tag = "CANCEL" Then Exit Sub
 txtbull2.TEXT = selsire_list.Tag
End Sub


Private Sub txtbull3_DblClick()
 selsire_list.Show vbModal
 If selsire_list.Tag = "CANCEL" Then Exit Sub
 txtbull3.TEXT = selsire_list.Tag
End Sub


Private Sub txtcw_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtcw_LostFocus()
If Val(txtcw.TEXT) < 0 Or Val(txtcw.TEXT) > 100 Then
   MsgBox "Please enter a number between 0 and 100", vbOKOnly
   SSTab1.Tab = 2
   txtcw.SetFocus
   txtcw.SelStart = 0
   txtcw.SelLength = Len(txtcw.TEXT)
End If
End Sub


Private Sub txtdam_DblClick()
 selcow_list.SetMode = 1
 selcow_list.Show vbModal
 If selcow_list.Tag = "CANCEL" Then Exit Sub
 txtdam.TEXT = selcow_list.Tag
End Sub


Private Sub txteage_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txteage_LostFocus()
If Val(txteage.TEXT) < 0 Or Val(txteage.TEXT) > 286 Then
   MsgBox "Please enter a number between 0 and 286", vbOKOnly
   SSTab1.Tab = 2
   txteage.SetFocus
   txteage.SelStart = 0
   txteage.SelLength = Len(txteage.TEXT)
End If
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

Private Sub txtmpda_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtsire_DblClick()
 selsire_list.Show vbModal
 If selsire_list.Tag = "CANCEL" Then Exit Sub
 txtsire.TEXT = selsire_list.Tag
End Sub


Private Sub txtweanwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub txtyearwt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericOnly(KeyAscii)
End Sub
