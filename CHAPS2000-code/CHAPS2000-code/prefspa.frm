VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmprefspa 
   Caption         =   "SPA Defaults"
   ClientHeight    =   5700
   ClientLeft      =   645
   ClientTop       =   1650
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   7860
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6735
      TabIndex        =   82
      Top             =   5295
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   81
      Top             =   5280
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   80
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "SPA"
      TabPicture(0)   =   "prefspa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CSF"
      TabPicture(1)   =   "prefspa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "prefspa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame6 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   7455
         Begin VB.TextBox txtadg 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   41
            Top             =   250
            Width           =   1000
         End
         Begin VB.TextBox txtbwt 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   45
            Top             =   950
            Width           =   1000
         End
         Begin VB.TextBox txt205dwt 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1300
            Width           =   1000
         End
         Begin VB.TextBox txtfrsc 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   49
            Top             =   1650
            Width           =   1000
         End
         Begin VB.TextBox txthfrearly 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   51
            Top             =   2000
            Width           =   1000
         End
         Begin VB.TextBox txthfr21 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   53
            Top             =   2350
            Width           =   1000
         End
         Begin VB.TextBox txthfr42 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   55
            Top             =   2700
            Width           =   1000
         End
         Begin VB.TextBox txtcow21 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   57
            Top             =   3050
            Width           =   1000
         End
         Begin VB.TextBox txtcow42 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   59
            Top             =   3400
            Width           =   1000
         End
         Begin VB.TextBox txtcowage 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   61
            Top             =   3750
            Width           =   1000
         End
         Begin VB.TextBox txtwda 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   43
            Top             =   600
            Width           =   1000
         End
         Begin VB.Frame fracowcond 
            Caption         =   "Cow Condition"
            Height          =   1395
            Left            =   3360
            TabIndex        =   71
            Top             =   2280
            Width           =   2535
            Begin VB.TextBox txtcalfwt 
               Height          =   285
               Left            =   1200
               TabIndex        =   75
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txttri 
               Height          =   285
               Left            =   1200
               TabIndex        =   79
               Top             =   1320
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtbreedwt 
               Height          =   285
               Left            =   1200
               TabIndex        =   73
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtwean 
               Height          =   285
               Index           =   1
               Left            =   1200
               TabIndex        =   77
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calving"
               Height          =   255
               Left            =   105
               TabIndex        =   74
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weaning"
               Height          =   255
               Left            =   105
               TabIndex        =   76
               Top             =   960
               Width           =   1005
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Breeding"
               Height          =   255
               Left            =   105
               TabIndex        =   72
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "3rd Trimester"
               Height          =   255
               Left            =   105
               TabIndex        =   78
               Top             =   1320
               Visible         =   0   'False
               Width           =   1005
            End
         End
         Begin VB.Frame fracowwt 
            Caption         =   "Cow Weight"
            Height          =   1455
            Left            =   3360
            TabIndex        =   62
            Top             =   240
            Width           =   2535
            Begin VB.TextBox txtcalwt 
               Height          =   285
               Left            =   1200
               TabIndex        =   66
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txttriwt 
               Height          =   285
               Left            =   1200
               TabIndex        =   70
               Top             =   1320
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtbreed 
               Height          =   285
               Left            =   1200
               TabIndex        =   64
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtwean 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   68
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calving"
               Height          =   255
               Left            =   105
               TabIndex        =   65
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "3rd Trimester"
               Height          =   255
               Left            =   105
               TabIndex        =   69
               Top             =   1320
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Breeding"
               Height          =   255
               Left            =   105
               TabIndex        =   63
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weaning"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   960
               Width           =   1005
            End
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ADG"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   250
            Width           =   1000
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "WDA"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   1000
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Weight"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   950
            Width           =   1000
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "205 Day Wt."
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1300
            Width           =   1000
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Frame Score"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1650
            Width           =   1000
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Hfrs Early"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   2000
            Width           =   1000
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Hfrs 21"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2350
            Width           =   1000
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Hfrs 42"
            Height          =   255
            Left            =   105
            TabIndex        =   54
            Top             =   2700
            Width           =   1005
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Cows 21"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   3050
            Width           =   1000
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Cows 42"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   3400
            Width           =   1000
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cow Age"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   3750
            Width           =   1000
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   7455
         Begin VB.Frame Frame5 
            Caption         =   "Reproduction Perf. Measures Based on Exposed Females"
            Height          =   2175
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   7215
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   6
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   15
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   5
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   13
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   4
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   11
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   3
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   9
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   2
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   7
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   1
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   5
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox perf 
               Height          =   285
               Index           =   0
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   3
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Pregnancy Percentage"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Pregnancy Loss Percentage"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   4
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label1 
               Caption         =   "Calving Percentage"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   6
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Calf Death Loss"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   8
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Calf Crop or Weaning Percentage"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   10
               Top             =   1200
               Width           =   2415
            End
            Begin VB.Label Label1 
               Caption         =   "Calf Death Loss Based On Number Of Calves Born"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   14
               Top             =   1680
               Width           =   3615
            End
            Begin VB.Label Label1 
               Caption         =   "Female Replacement Rate Percentage"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   12
               Top             =   1440
               Width           =   2775
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Production Performance Measures"
            Height          =   2100
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   7215
            Begin VB.TextBox Aveage 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   18
               Top             =   375
               Width           =   615
            End
            Begin VB.TextBox born 
               Height          =   285
               Index           =   0
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   20
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox born 
               Height          =   285
               Index           =   1
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   22
               Top             =   1050
               Width           =   615
            End
            Begin VB.TextBox born 
               Height          =   285
               Index           =   2
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   24
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox born 
               Height          =   285
               Index           =   3
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   26
               Top             =   1590
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   5
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   38
               Top             =   1575
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   4
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   36
               Top             =   1335
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   3
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   34
               Top             =   1095
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   2
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   32
               Top             =   855
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   1
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   30
               Top             =   615
               Width           =   615
            End
            Begin VB.TextBox weights 
               Height          =   285
               Index           =   0
               Left            =   6450
               MaxLength       =   10
               TabIndex        =   28
               Top             =   375
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "Average Age At Weaning"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label2 
               Caption         =   "Calves Born During First 42 Days"
               Height          =   255
               Index           =   4
               Left            =   105
               TabIndex        =   21
               Top             =   1050
               Width           =   2775
            End
            Begin VB.Label Label2 
               Caption         =   "Calves Born During First 63 Days"
               Height          =   255
               Index           =   3
               Left            =   105
               TabIndex        =   23
               Top             =   1365
               Width           =   2775
            End
            Begin VB.Label Label2 
               Caption         =   "Calves Born During After 63 Days"
               Height          =   255
               Index           =   2
               Left            =   105
               TabIndex        =   25
               Top             =   1635
               Width           =   2775
            End
            Begin VB.Label Label2 
               Caption         =   "Calves Born During First 21 Days"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   780
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "Steers Weaning Weight"
               Height          =   255
               Index           =   0
               Left            =   3690
               TabIndex        =   27
               Top             =   375
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Bulls Weaning Weight"
               Height          =   255
               Index           =   2
               Left            =   3690
               TabIndex        =   31
               Top             =   855
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Heifers Weaning Weight"
               Height          =   255
               Index           =   3
               Left            =   3690
               TabIndex        =   29
               Top             =   615
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Average Weaning Weight"
               Height          =   255
               Index           =   4
               Left            =   3690
               TabIndex        =   33
               Top             =   1095
               Width           =   1935
            End
            Begin VB.Label Label3 
               Caption         =   "Pounds Weaned per Exposed Female"
               Height          =   255
               Index           =   5
               Left            =   3690
               TabIndex        =   35
               Top             =   1335
               Width           =   2775
            End
            Begin VB.Label Label3 
               Caption         =   "SPA Source (ex. NORTH DAKOTA)"
               Height          =   255
               Index           =   6
               Left            =   3690
               TabIndex        =   37
               Top             =   1575
               Width           =   2655
            End
         End
      End
   End
End
Attribute VB_Name = "frmprefspa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Load_information()
 Dim tbpref As Recordset
 Dim tbprefcsf As Recordset
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbpref = DB.OpenRecordset("prefspa", dbOpenTable)
 Set tbprefcsf = DB.OpenRecordset("prefcsf", dbOpenTable)
 If Not tbpref.EOF And Not tbpref.BOF Then
   perf(0) = tbpref!PregPer
   perf(1) = tbpref!pregperlos
   perf(2) = tbpref!calfper
   perf(3) = tbpref!CalfDeath
   perf(4) = tbpref!WeanPer
   perf(5) = tbpref!femreplc
   perf(6) = tbpref!deathbased
   born(0) = tbpref!born21
   born(1) = tbpref!born42
   born(2) = tbpref!born63
   born(3) = tbpref!born63pls
   Aveage = tbpref!weanage
   weights(0) = tbpref!weanwtst
   weights(1) = tbpref!weanwthf
   weights(2) = tbpref!weanwtbl
   weights(3) = tbpref!weanwtave
   weights(4) = tbpref!lbsweaned
   weights(5) = tbpref!spasource
   tbpref.Close: Set tbpref = Nothing
 End If
 If Not tbprefcsf.EOF And Not tbprefcsf.BOF Then
   txtadg.TEXT = Field2Str(tbprefcsf!adg)
   txtwda.TEXT = Field2Str(tbprefcsf!wda)
   txtbwt.TEXT = Field2Str(tbprefcsf!birthweight)
   txt205dwt.TEXT = Field2Str(tbprefcsf!day205)
   txtfrsc.TEXT = Field2Str(tbprefcsf!score)
   txthfrearly.TEXT = Field2Str(tbprefcsf!hefearly)
   txthfr21.TEXT = Field2Str(tbprefcsf!hef21)
   txthfr42.TEXT = Field2Str(tbprefcsf!hef42)
   txtcow21.TEXT = Field2Str(tbprefcsf!cow21)
   txtcow42.TEXT = Field2Str(tbprefcsf!cow42)
   txtcowage.TEXT = Field2Str(tbprefcsf!cowage)
   txtwean(0).TEXT = Field2Str(tbprefcsf!weanwt)
   txtbreed.TEXT = Field2Str(tbprefcsf!breedwt)
   txttriwt.TEXT = Field2Str(tbprefcsf!triwt)
   txtcalwt.TEXT = Field2Str(tbprefcsf!calwt)
   txtwean(1).TEXT = Field2Str(tbprefcsf!wean)
   txtbreedwt.TEXT = Field2Str(tbprefcsf!breed)
   txttri.TEXT = Field2Str(tbprefcsf!tri)
   txtcalfwt.TEXT = Field2Str(tbprefcsf!calf)
 End If
 tbprefcsf.Close: Set tbprefcsf = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub
Private Sub save_information()
 Dim tbpref As Recordset
 Dim tbprefcsf As Recordset
 Screen.MousePointer = vbHourglass
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 Set tbpref = DB.OpenRecordset("prefspa", dbOpenTable)
 Set tbprefcsf = DB.OpenRecordset("prefcsf", dbOpenTable)
 If Not tbpref.EOF And Not tbpref.BOF Then
     tbpref.Edit
 Else
   tbpref.AddNew
 End If
 With tbpref
   !PregPer = perf(0)
   !pregperlos = perf(1)
   !calfper = perf(2)
   !CalfDeath = perf(3)
   !WeanPer = perf(4)
   !femreplc = perf(5)
   !deathbased = perf(6)
   !born21 = born(0)
   !born42 = born(1)
   !born63 = born(2)
   !born63pls = born(3)
   !weanage = Aveage
   !weanwtst = weights(0)
   !weanwthf = weights(1)
   !weanwtbl = weights(2)
   !weanwtave = weights(3)
   !lbsweaned = weights(4)
   !spasource = weights(5)
 End With
 tbpref.Update
 If Not tbprefcsf.EOF And Not tbprefcsf.BOF Then
     tbprefcsf.Edit
 Else
   tbprefcsf.AddNew
 End If
 With tbprefcsf
   !adg = txtadg.TEXT
   !wda = txtwda.TEXT
   !birthweight = txtbwt.TEXT
   !day205 = txt205dwt.TEXT
   !score = txtfrsc.TEXT
   !hefearly = txthfrearly.TEXT
   !hef21 = txthfr21.TEXT
   !hef42 = txthfr42.TEXT
   !cow21 = txtcow21.TEXT
   !cow42 = txtcow42.TEXT
   !cowage = txtcowage.TEXT
   !weanwt = txtwean(0).TEXT
   !breedwt = txtbreed.TEXT
   !triwt = txttriwt.TEXT
   !calwt = txtcalwt.TEXT
   !wean = txtwean(1).TEXT
   !breed = txtbreedwt.TEXT
   !tri = txttri.TEXT
   !calf = txtcalfwt.TEXT
 End With
 tbprefcsf.Update
 tbpref.Close: Set tbpref = Nothing
 tbprefcsf.Close: Set tbprefcsf = Nothing
 DB.Close: Set DB = Nothing
 Screen.MousePointer = vbDefault
End Sub

Private Sub CMDCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Call save_information
    Unload Me
End Sub


Private Sub Form_Activate()
  SSTab1.Tab = 0
  perf(0).SetFocus
End Sub

Private Sub Form_Load()
  Call centermdiform(Me, mdimain, 0, 0)
  Call Load_information
End Sub

Private Sub SSTab1_GotFocus()
   If SSTab1.Tab = 0 Then perf(0).SetFocus
   If SSTab1.Tab = 1 Then txtadg.SetFocus
End Sub

Private Sub txtadg_GotFocus()
  SSTab1.Tab = 1
End Sub
