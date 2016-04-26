VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmImportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Chaps Data"
   ClientHeight    =   5745
   ClientLeft      =   705
   ClientTop       =   2700
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcframe 
      Caption         =   "&Calc Cframe"
      Height          =   385
      Left            =   1290
      TabIndex        =   49
      Top             =   5265
      Width           =   1000
   End
   Begin VB.CommandButton CMDforms 
      Caption         =   "&Forms"
      Height          =   385
      Left            =   165
      TabIndex        =   17
      Top             =   5250
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5040
      Left            =   150
      TabIndex        =   2
      Top             =   105
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   8890
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Select Spreadsheet"
      TabPicture(0)   =   "Impdata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Column Setup"
      TabPicture(1)   =   "Impdata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Data Preview"
      TabPicture(2)   =   "Impdata.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4635
         Left            =   -74865
         TabIndex        =   5
         Top             =   330
         Width           =   7620
         Begin VB.ComboBox CboMarblingScore 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2190
            Width           =   1800
         End
         Begin VB.ComboBox Cbocolor 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   2535
            Width           =   1785
         End
         Begin VB.ComboBox Cbomaturity 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   3210
            Width           =   1785
         End
         Begin VB.ComboBox Cbotextureoflean 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   2865
            Width           =   1800
         End
         Begin VB.ComboBox CboCarcassDate 
            Height          =   315
            Left            =   4845
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   3525
            Width           =   1800
         End
         Begin VB.ComboBox cboyieldgrade 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   195
            Width           =   1800
         End
         Begin VB.ComboBox CBOhotCarcassweight 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   540
            Width           =   1785
         End
         Begin VB.ComboBox CboKidney 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1215
            Width           =   1785
         End
         Begin VB.ComboBox cbofatthickness 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   870
            Width           =   1800
         End
         Begin VB.ComboBox CboRibEye 
            Height          =   315
            Left            =   4845
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1530
            Width           =   1800
         End
         Begin VB.ComboBox CboQualityGrade 
            Height          =   315
            Left            =   4845
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1860
            Width           =   1785
         End
         Begin VB.ComboBox cbocalfid 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   195
            Width           =   1800
         End
         Begin VB.ComboBox cboeid 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   540
            Width           =   1785
         End
         Begin VB.CommandButton cmdload2 
            Caption         =   "L&oad"
            Height          =   405
            Left            =   6600
            TabIndex        =   24
            Top             =   4140
            Width           =   900
         End
         Begin VB.ComboBox cbowd 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1215
            Width           =   1785
         End
         Begin VB.ComboBox cboww 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   870
            Width           =   1800
         End
         Begin VB.ComboBox Cbodatemeasured 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1530
            Width           =   1800
         End
         Begin VB.ComboBox Cbohh 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1860
            Width           =   1785
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Marbling Score"
            Height          =   270
            Left            =   3240
            TabIndex        =   43
            Top             =   2235
            Width           =   1500
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Color"
            Height          =   270
            Left            =   3240
            TabIndex        =   42
            Top             =   2565
            Width           =   1500
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Maturity"
            Height          =   270
            Left            =   3240
            TabIndex        =   41
            Top             =   3225
            Width           =   1500
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Texture Of Lean"
            Height          =   270
            Left            =   3240
            TabIndex        =   40
            Top             =   2895
            Width           =   1500
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Carcass Date"
            Height          =   270
            Left            =   3225
            TabIndex        =   39
            Top             =   3555
            Width           =   1500
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Yield Grade"
            Height          =   270
            Left            =   3240
            TabIndex        =   32
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Hot Carcass Weight"
            Height          =   270
            Left            =   3240
            TabIndex        =   31
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Kidney (KPH)"
            Height          =   270
            Left            =   3240
            TabIndex        =   30
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Fat Thickness"
            Height          =   270
            Left            =   3240
            TabIndex        =   29
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Rib Eye"
            Height          =   270
            Left            =   3225
            TabIndex        =   28
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Quality Grade (Alpha)"
            Height          =   270
            Left            =   3225
            TabIndex        =   27
            Top             =   1875
            Width           =   1500
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Hip Height"
            Height          =   270
            Left            =   30
            TabIndex        =   19
            Top             =   1875
            Width           =   1230
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Measured"
            Height          =   270
            Left            =   30
            TabIndex        =   18
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Weaning Weight"
            Height          =   270
            Left            =   45
            TabIndex        =   16
            Top             =   900
            Width           =   1230
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Weaning Date"
            Height          =   270
            Left            =   45
            TabIndex        =   15
            Top             =   1230
            Width           =   1230
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "E ID"
            Height          =   270
            Left            =   45
            TabIndex        =   9
            Top             =   570
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Calf ID"
            Height          =   270
            Left            =   45
            TabIndex        =   8
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4635
         Left            =   -74910
         TabIndex        =   4
         Top             =   315
         Width           =   7695
         Begin FPSpread.vaSpread grddata 
            Height          =   4335
            Left            =   105
            OleObjectBlob   =   "Impdata.frx":0054
            TabIndex        =   11
            Top             =   210
            Width           =   7440
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4620
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   7635
         Begin VB.CommandButton cmdloadcols 
            Caption         =   "Load Columns"
            Height          =   360
            Left            =   4065
            TabIndex        =   14
            Top             =   1560
            Width           =   1260
         End
         Begin VB.ComboBox cbosheet 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1590
            Width           =   1365
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load &Sheets"
            Height          =   390
            Left            =   6075
            TabIndex        =   10
            Top             =   255
            Width           =   1230
         End
         Begin VB.TextBox TxtSpreadSheet 
            Height          =   285
            Left            =   3075
            TabIndex        =   6
            Top             =   270
            Width           =   2900
         End
         Begin MSComDlg.CommonDialog CDIPRINTSET 
            Left            =   6225
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label4 
            Caption         =   "Select Sheetname"
            Height          =   255
            Left            =   1260
            TabIndex        =   12
            Top             =   1650
            Width           =   1635
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "<Spreadsheet Filename>"
            Height          =   285
            Left            =   75
            TabIndex        =   7
            Top             =   285
            Width           =   2940
         End
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Import"
      Height          =   385
      Left            =   5955
      TabIndex        =   0
      Top             =   5265
      Width           =   1000
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   7035
      TabIndex        =   1
      Top             =   5250
      Width           =   1000
   End
End
Attribute VB_Name = "frmImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONN As String
Private Function calcscore(ID As String, hh As String) As Double
 'Call calc_205
 Dim numofday
 Dim SQL As String
 Dim DB As database
 Dim tbData As Recordset
 Dim j As Double
 Dim jj As String
 Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
 SQL = "SELECT calfbirth.birthdate, calfwean.cdatemeas, calfbirth.CalfID, sex FROM calfbirth INNER JOIN calfwean ON (calfbirth.CalfID = calfwean.CalfID) AND (calfbirth.HerdID = calfwean.HerdID) WHERE calfbirth.CalfID='" & ID & "'"
 Set tbData = DB.OpenRecordset(SQL, dbOpenDynaset)
 If Not tbData.NoMatch Then
   If IsDate(tbData!birthdate) And IsDate(tbData!cdatemeas) Then
      numofday = DateDiff("d", tbData!birthdate, tbData!cdatemeas)
      If numofday < 0 Then
         MsgBox ("Invalid Date Measured")
         calcscore = 0
         tbData.Close: Set tbData = Nothing
         DB.Close: Set DB = Nothing
         Exit Function
      End If
      
      If Val(tbData!Sex) = 2 Then
        j = -11.7086 + (0.4723 * Val(hh)) - (0.0239 * numofday) + (0.0000146 * (numofday * numofday)) + (0.0000759 * (Val(hh)) * numofday)
      End If
      If Val(tbData!Sex) <> 2 Then
         j = -11.548 + (0.4878 * Val(hh)) - (0.0289 * numofday) + (0.00001947 * (numofday * numofday)) + (0.0000334 * (Val(hh)) * numofday)
      End If
      jj$ = funround2(1, j)
      
      calcscore = Val(jj$)
      tbData.Close: Set tbData = Nothing
      DB.Close: Set DB = Nothing
      Exit Function
   End If
 End If
 
 calcscore = 0
 tbData.Close: Set tbData = Nothing
 DB.Close: Set DB = Nothing
End Function

Private Sub initgrid()
  grddata.MaxCols = 17
  grddata.MaxRows = 0
  
  grddata.SetText 1, 0, "Calf ID"
  grddata.SetText 2, 0, "E ID"
  grddata.SetText 3, 0, "Wean Weight"
  grddata.SetText 4, 0, "Wean Date"
  grddata.SetText 5, 0, "Date Measured"
  grddata.SetText 6, 0, "Hip Height"
  grddata.SetText 7, 0, "Yield Grade"
  grddata.SetText 8, 0, "Hot Carcass Weight"
  grddata.SetText 9, 0, "Fat Thickness"
  grddata.SetText 10, 0, "Kidney (KPH)"
  grddata.SetText 11, 0, "Rib Eye"
  grddata.SetText 12, 0, "Quality Grade (Alpha)"
  grddata.SetText 13, 0, "Marbling Score"
  grddata.SetText 14, 0, "Color"
  grddata.SetText 15, 0, "Texture Of Lean"
  grddata.SetText 16, 0, "Maturity"
  grddata.SetText 17, 0, "Carcass Date"
  
    

  grddata.ColWidth(1) = 2000
  grddata.ColWidth(2) = 2000
  grddata.ColWidth(3) = 2000
  grddata.ColWidth(4) = 2000
  grddata.ColWidth(5) = 2000
  grddata.ColWidth(6) = 2000
  grddata.ColWidth(7) = 2000
  grddata.ColWidth(8) = 2000
  grddata.ColWidth(9) = 2000
  grddata.ColWidth(10) = 2000
  grddata.ColWidth(11) = 2000
  grddata.ColWidth(12) = 2000
  grddata.ColWidth(13) = 2000
  grddata.ColWidth(14) = 2000
  grddata.ColWidth(15) = 2000
  grddata.ColWidth(16) = 2000
  grddata.ColWidth(17) = 2000
  
  grddata.Row = -1
  grddata.Col = 1
  grddata.CellType = SS_CELL_TYPE_EDIT
  grddata.TypeHAlign = SS_CELL_H_ALIGN_LEFT
  grddata.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
  grddata.TypeEditMultiLine = False
  grddata.TypeEditLen = 8
  
  
  grddata.Col = 2
  grddata.CellType = SS_CELL_TYPE_EDIT
  grddata.TypeHAlign = SS_CELL_H_ALIGN_LEFT
  grddata.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
  grddata.TypeEditMultiLine = False
  grddata.TypeEditLen = 20
  
  
End Sub

Private Sub validform(exitcode As Integer)
  Dim i As Long
  Dim ii As Long
  Dim DB As database
  Dim tbData As Recordset
  Dim SQL As String
  Dim var
  Dim var2
  Dim res As Integer
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  For i = 1 To grddata.MaxRows
    grddata.GetText 1, i, var
    SQL = "SELECT * FROM calfbirth WHERE calfid = '" & var & "'"
    Set tbData = DB.OpenRecordset(SQL, dbOpenDynaset)
    If tbData.EOF And tbData.BOF Then
       Beep
       MsgBox "Row " & i & " must have a valid calf ID.", vbOKOnly
       exitcode = 1
       tbData.Close: Set tbData = Nothing
       DB.Close: Set DB = Nothing
       Exit Sub
    End If
        
    If cboeid.TEXT <> "'None'" Then
      grddata.GetText 2, i, var
      If var <> "" Then
        If Len(var) <> 15 Then
            MsgBox "Row " & i & " EID must be 15 characters.", vbOKOnly
            exitcode% = 1
            Exit Sub
        End If
      End If
      
      If var = "" Then
         res = MsgBox("Row " & i & " Has a blank EID.  Do you wish to continue?", vbYesNo)
         If res = vbNo Then
           exitcode = 1
           Exit Sub
         End If
      End If
      grddata.GetText 1, i, var2
      If Check_EID(CStr(var), "Calf", "", "", CStr(var2), "Calf") = False Then
         MsgBox "EID on line " & i & " is on file already", vbOKOnly + vbCritical, Me.Caption
         exitcode% = 1
         Exit Sub
      End If
      For ii = i + 1 To grddata.MaxRows
        grddata.GetText 2, ii, var2
        If var = var2 Then
          MsgBox "The eid on line " & i & " and " & ii & " are Duplicated.", vbOKOnly
          exitcode = 1
           Exit Sub
        End If
      Next ii
    End If
  Next i
  
  
  
  tbData.Close: Set tbData = Nothing
  DB.Close: Set DB = Nothing
End Sub

Private Sub update_info()
  Dim i As Long
  Dim DB As database
  Dim SQL As String
  Dim cid
  Dim eid
  Dim ww
  Dim score As Double
  Set DB = DBEngine(0).OpenDatabase(dbfile$, False, False)
  For i = 1 To grddata.MaxRows
    grddata.GetText 1, i, cid
    grddata.GetText 2, i, eid
    If cboeid.TEXT <> "'None'" Then
      SQL = "update CALFbirth set elecid = '" & eid & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    If cboww.TEXT <> "'None'" Then
      grddata.GetText 3, i, ww
      SQL = "update calfwean set actweight = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    If cbowd.TEXT <> "'None'" Then
      grddata.GetText 4, i, ww
      SQL = "update calfwean set dateweighed = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If Cbodatemeasured.TEXT <> "'None'" Then
      grddata.GetText 5, i, ww
      SQL = "update calfwean set cdatemeas = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
          
    If Cbohh.TEXT <> "'None'" Then
      grddata.GetText 6, i, ww
      SQL = "update calfwean set chipheight = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
      
      score = 0
      score = calcscore(CStr(cid), CStr(ww))
      If score <> 0 Then
        SQL = "update calfwean set score = " & score & " WHERE calfid = '" & cid & "'"
        DB.Execute SQL
      End If
    End If
    
    If cboyieldgrade.TEXT <> "'None'" Then
      grddata.GetText 7, i, ww
      SQL = "update calfCARCASS set YGRADE = " & ww & " WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CBOhotCarcassweight.TEXT <> "'None'" Then
      grddata.GetText 8, i, ww
      SQL = "update calfCARCASS set YWT = " & ww & " WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If cbofatthickness.TEXT <> "'None'" Then
      grddata.GetText 9, i, ww
      SQL = "update calfCARCASS set YFAT = " & ww & " WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CboKidney.TEXT <> "'None'" Then
      grddata.GetText 10, i, ww
      SQL = "update calfCARCASS set YKIDNEY = " & ww & " WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CboRibEye.TEXT <> "'None'" Then
      grddata.GetText 11, i, ww
      SQL = "update calfCARCASS set YRIBEYE = " & ww & " WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CboQualityGrade.TEXT <> "'None'" Then
      grddata.GetText 12, i, ww
      SQL = "update calfCARCASS set QGRADE = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CboMarblingScore.TEXT <> "'None'" Then
      grddata.GetText 13, i, ww
      SQL = "update calfCARCASS set QSCORE = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If Cbocolor.TEXT <> "'None'" Then
      grddata.GetText 14, i, ww
      SQL = "update calfCARCASS set QCOLOR = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
 
    If Cbotextureoflean.TEXT <> "'None'" Then
      grddata.GetText 15, i, ww
      SQL = "update calfCARCASS set QTEXTURE = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If Cbomaturity.TEXT <> "'None'" Then
      grddata.GetText 16, i, ww
      SQL = "update calfCARCASS set QMATURITY = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
    
    If CboCarcassDate.TEXT <> "'None'" Then
      grddata.GetText 17, i, ww
      SQL = "update calfCARCASS set CARCASSDATE = '" & ww & "' WHERE calfid = '" & cid & "'"
      DB.Execute SQL
    End If
      
  Next i
  DB.Close: Set DB = Nothing
  Unload Me
End Sub


Private Sub CMDCancel_Click()
  Unload Me
End Sub

Private Sub cmdcframe_Click()
  If grddata.RowsFrozen < 1 Then Exit Sub
  
End Sub

Private Sub CMDforms_Click()
  Dim DBREP As database
  Dim RS As Recordset
  Dim i As Long
  Dim eid
  Screen.MousePointer = vbHourglass
  report.Initialize
  Set DBREP = DBEngine(0).OpenDatabase(repfile$, False, False)
  DBREP.Execute ("delete * from eidform")
  
  For i = 1 To grddata.MaxRows
    Set RS = DBREP.OpenRecordset("eidform")
    With RS
     .AddNew
     grddata.GetText 2, i, eid
      RS!eid = eid
      RS!eidfirsthalf = Left(eid, 10)
      RS!eidsecondhalf = Right(eid, 5)
     .Update
    End With
  Next i
  RS.Close: Set RS = Nothing
  DBREP.Close: Set DBREP = Nothing
  report.SetReportFileName = dbdir$ & "\" & "EIDINPUT.rpt"
  report.setDbname = repfile$
  report.SetReportCaption = "EID Import Forms"
  report.Setorientation = True
  report.PrintReport
  Screen.MousePointer = vbDefault
End Sub

Private Sub CmdLoad_Click()
 Dim dbExcel As database
 Dim tbl As TableDef
 cbosheet.Clear
 If TxtSpreadSheet.TEXT = "" Then Exit Sub
 Screen.MousePointer = vbHourglass

 Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
 For Each tbl In dbExcel.TableDefs
   cbosheet.AddItem tbl.Name
 Next
 Screen.MousePointer = vbDefault

End Sub

Private Sub cmdload2_Click()
Dim dbExcel As database
Dim snpExcel As Recordset
Dim SQL As String
Dim i As Long
Dim cid As String
Dim eid As String
Dim ww As String
Dim wd As String
Dim dm As String
Dim hh As String
Dim YG, HCW, ft, k, r, qg, ms, c, tl, m, cd As String

 Screen.MousePointer = vbHourglass
 Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
 
If cbocalfid.TEXT <> "'None'" Then cid = "[" & cbocalfid.TEXT & "]" Else cid = cbocalfid.TEXT
If cboeid.TEXT <> "'None'" Then eid = "[" & cboeid.TEXT & "]" Else eid = cboeid.TEXT
If cboww.TEXT <> "'None'" Then ww = "[" & cboww.TEXT & "]" Else ww = cboww.TEXT
If cbowd.TEXT <> "'None'" Then wd = "[" & cbowd.TEXT & "]" Else wd = cbowd.TEXT
If Cbodatemeasured.TEXT <> "'None'" Then dm = "[" & Cbodatemeasured.TEXT & "]" Else dm = Cbodatemeasured.TEXT
If Cbohh.TEXT <> "'None'" Then hh = "[" & Cbohh.TEXT & "]" Else hh = Cbohh.TEXT
If cboyieldgrade.TEXT <> "'None'" Then YG = "[" & cboyieldgrade.TEXT & "]" Else YG = cboyieldgrade.TEXT
If CBOhotCarcassweight.TEXT <> "'None'" Then HCW = "[" & CBOhotCarcassweight.TEXT & "]" Else HCW = CBOhotCarcassweight.TEXT
If cbofatthickness.TEXT <> "'None'" Then ft = "[" & cbofatthickness.TEXT & "]" Else ft = cbofatthickness.TEXT
If CboKidney.TEXT <> "'None'" Then k = "[" & CboKidney.TEXT & "]" Else k = CboKidney.TEXT
If CboRibEye.TEXT <> "'None'" Then r = "[" & CboRibEye.TEXT & "]" Else r = CboRibEye.TEXT
If CboQualityGrade.TEXT <> "'None'" Then qg = "[" & CboQualityGrade.TEXT & "]" Else qg = CboQualityGrade.TEXT
If CboMarblingScore.TEXT <> "'None'" Then ms = "[" & CboMarblingScore.TEXT & "]" Else ms = CboMarblingScore.TEXT
If Cbocolor.TEXT <> "'None'" Then c = "[" & Cbocolor.TEXT & "]" Else c = Cbocolor.TEXT
If Cbotextureoflean.TEXT <> "'None'" Then tl = "[" & Cbotextureoflean.TEXT & "]" Else tl = Cbotextureoflean.TEXT
If Cbomaturity.TEXT <> "'None'" Then m = "[" & Cbomaturity.TEXT & "]" Else m = Cbomaturity.TEXT
If CboCarcassDate.TEXT <> "'None'" Then cd = "[" & CboCarcassDate.TEXT & "]" Else cd = CboCarcassDate.TEXT
  
 
 
 SQL = "select " & cid & ", " & eid & ", " & ww & ", " & wd & ", " & dm & ", " & hh & ", " & YG & ", " & HCW & ", " & ft & ", " & k & ", " & r & ", " & qg & ", " & ms & ", " & c & ", " & tl & ", " & m & ", " & cd & " from [" & cbosheet.TEXT & "]"
 Set snpExcel = dbExcel.OpenRecordset(SQL, dbOpenSnapshot)
 i = 0
 grddata.MaxRows = 0
 While Not snpExcel.EOF
    i = i + 1
    grddata.MaxRows = grddata.MaxRows + 1
    If Not IsNull(snpExcel(cbocalfid.TEXT)) Then grddata.SetText 1, i, CStr(snpExcel(cbocalfid.TEXT))
    If cboeid.TEXT <> "'None'" Then If Not IsNull(snpExcel(cboeid.TEXT)) Then grddata.SetText 2, i, CStr(snpExcel(cboeid.TEXT))
    If cboww.TEXT <> "'None'" Then If Not IsNull(snpExcel(cboww.TEXT)) Then grddata.SetText 3, i, CStr(snpExcel(cboww.TEXT))
    If cbowd.TEXT <> "'None'" Then If Not IsNull(snpExcel(cbowd.TEXT)) Then grddata.SetText 4, i, CStr(snpExcel(cbowd.TEXT))
    If Cbodatemeasured.TEXT <> "'None'" Then If Not IsNull(snpExcel(Cbodatemeasured.TEXT)) Then grddata.SetText 5, i, CStr(snpExcel(Cbodatemeasured.TEXT))
    If Cbohh.TEXT <> "'None'" Then If Not IsNull(snpExcel(Cbohh.TEXT)) Then grddata.SetText 6, i, CStr(snpExcel(Cbohh.TEXT))
    If cboyieldgrade.TEXT <> "'None'" Then If Not IsNull(snpExcel(cboyieldgrade.TEXT)) Then grddata.SetText 7, i, CStr(snpExcel(cboyieldgrade.TEXT))
    If CBOhotCarcassweight.TEXT <> "'None'" Then If Not IsNull(snpExcel(CBOhotCarcassweight.TEXT)) Then grddata.SetText 8, i, CStr(snpExcel(CBOhotCarcassweight.TEXT))
    If cbofatthickness.TEXT <> "'None'" Then If Not IsNull(snpExcel(cbofatthickness.TEXT)) Then grddata.SetText 9, i, CStr(snpExcel(cbofatthickness.TEXT))
    If CboKidney.TEXT <> "'None'" Then If Not IsNull(snpExcel(CboKidney.TEXT)) Then grddata.SetText 10, i, CStr(snpExcel(CboKidney.TEXT))
    If CboRibEye.TEXT <> "'None'" Then If Not IsNull(snpExcel(CboRibEye.TEXT)) Then grddata.SetText 11, i, CStr(snpExcel(CboRibEye.TEXT))
    If CboQualityGrade.TEXT <> "'None'" Then If Not IsNull(snpExcel(CboQualityGrade.TEXT)) Then grddata.SetText 12, i, CStr(snpExcel(CboQualityGrade.TEXT))
    If CboMarblingScore.TEXT <> "'None'" Then If Not IsNull(snpExcel(CboMarblingScore.TEXT)) Then grddata.SetText 13, i, CStr(snpExcel(CboMarblingScore.TEXT))
    If Cbocolor.TEXT <> "'None'" Then If Not IsNull(snpExcel(Cbocolor.TEXT)) Then grddata.SetText 14, i, CStr(snpExcel(Cbocolor.TEXT))
    If Cbotextureoflean.TEXT <> "'None'" Then If Not IsNull(snpExcel(Cbotextureoflean.TEXT)) Then grddata.SetText 15, i, CStr(snpExcel(Cbotextureoflean.TEXT))
    If Cbomaturity.TEXT <> "'None'" Then If Not IsNull(snpExcel(Cbomaturity.TEXT)) Then grddata.SetText 16, i, CStr(snpExcel(Cbomaturity.TEXT))
    If CboCarcassDate.TEXT <> "'None'" Then If Not IsNull(snpExcel(CboCarcassDate.TEXT)) Then grddata.SetText 17, i, CStr(snpExcel(CboCarcassDate.TEXT))
    snpExcel.MoveNext
 Wend
 Screen.MousePointer = vbDefault
End Sub

Private Sub cmdloadcols_Click()
 Dim dbExcel As database
 Dim fld As Field
 Dim tbl As TableDef
 Screen.MousePointer = vbHourglass
 cbocalfid.Clear
 cboeid.Clear
 
 
 cboeid.AddItem "'None'"
 cbowd.AddItem "'None'"
 cboww.AddItem "'None'"
 Cbodatemeasured.AddItem "'None'"
 Cbohh.AddItem "'None'"
 cboyieldgrade.AddItem "'None'"
 CBOhotCarcassweight.AddItem "'None'"
 cbofatthickness.AddItem "'None'"
 CboKidney.AddItem "'None'"
 CboRibEye.AddItem "'None'"
 CboQualityGrade.AddItem "'None'"
 CboMarblingScore.AddItem "'None'"
 Cbocolor.AddItem "'None'"
 Cbotextureoflean.AddItem "'None'"
 Cbomaturity.AddItem "'None'"
 CboCarcassDate.AddItem "'None'"
 
 
 cboeid.ListIndex = 0
 cbowd.ListIndex = 0
 cboww.ListIndex = 0
 Cbodatemeasured.ListIndex = 0
 Cbohh.ListIndex = 0
 cboyieldgrade.ListIndex = 0
 CBOhotCarcassweight.ListIndex = 0
 cbofatthickness.ListIndex = 0
 CboKidney.ListIndex = 0
 CboRibEye.ListIndex = 0
 CboQualityGrade.ListIndex = 0
 CboMarblingScore.ListIndex = 0
 Cbocolor.ListIndex = 0
 Cbotextureoflean.ListIndex = 0
 Cbomaturity.ListIndex = 0
 CboCarcassDate.ListIndex = 0
 
 
 If TxtSpreadSheet.TEXT = "" Then Exit Sub
 Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
 For Each tbl In dbExcel.TableDefs
   If tbl.Name = cbosheet.TEXT Then
     For Each fld In tbl.Fields
       cbocalfid.AddItem fld.Name
       cboeid.AddItem fld.Name
       cbowd.AddItem fld.Name
       cboww.AddItem fld.Name
       Cbodatemeasured.AddItem fld.Name
       Cbohh.AddItem fld.Name
       cboyieldgrade.AddItem fld.Name
       CBOhotCarcassweight.AddItem fld.Name
       cbofatthickness.AddItem fld.Name
       CboKidney.AddItem fld.Name
       CboRibEye.AddItem fld.Name
       CboQualityGrade.AddItem fld.Name
       CboMarblingScore.AddItem fld.Name
       Cbocolor.AddItem fld.Name
       Cbotextureoflean.AddItem fld.Name
       Cbomaturity.AddItem fld.Name
       CboCarcassDate.AddItem fld.Name
     Next
   End If
 Next
 Screen.MousePointer = vbDefault
End Sub

Private Sub CMDOk_Click()
  'CONN = "Provider=SQLOLEDB;Server=SU_ryanb;Database=ranger;uid=Agvance;pwd=AgvSQL2000;driver={sqlserver};"
  'Call impgrower
  'Call impVEND
  'Call impacct
  'Call impprod
  Dim exitcode As Integer
  exitcode = 0
  Call validform(exitcode)
  If exitcode <> 0 Then Exit Sub
  Call update_info
  
End Sub


Private Sub impgrower()
' Dim dbExcel As database
' Dim snpExcel As Recordset
' Dim objcust As clsCustomer
' Dim tbl As TableDef
'
'
' Screen.MousePointer = vbHourglass
'
' Dim fld As Field
' Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
' For Each tbl In dbExcel.TableDefs
'   Debug.Print tbl.Name
'    For Each fld In tbl.Fields
'      Debug.Print fld.Name
'    Next
' Next
' Set snpExcel = dbExcel.OpenRecordset("Sheet1$", dbOpenSnapshot)
'
'
'
' While Not snpExcel.EOF
'   Set objcust = New clsCustomer
'          With objcust
'            .ID = Right(snpExcel("id"), 6)
'            .Growid = Right(snpExcel("id"), 6)
'            '.Location = snpExcel("f9")
'            '.xref1 = snpExcel("xref1")
'            '.xref2 = snpExcel("xref2")
'            '.xref3 = snpExcel("xref3")
'            '.xref4 = snpExcel("xref4")
'            .GROWNAME2 = snpExcel("Last Name or Company Name")
'            If Not IsNull(snpExcel("First name")) Then .GROWNAME1 = snpExcel("First name")
'            If Not IsNull(snpExcel("Contact(ADDRESS1)")) Then .CAREOF = PadTics(snpExcel("Contact(ADDRESS1)"))
'            If Not IsNull(snpExcel("PHYSICAL(ADDRESS 2)")) Then .ADDRESS = PadTics(snpExcel("PHYSICAL(ADDRESS 2)"))
'            If Not IsNull(snpExcel("PO BOX (ADDRESS3)")) Then .Address3 = PadTics(snpExcel("PO BOX (ADDRESS3)"))
'            If Not IsNull(snpExcel("city")) Then .City = snpExcel("City")
'            If Not IsNull(snpExcel("State")) Then .State = Left(snpExcel("State"), 2)
'            If Not IsNull(snpExcel("zip")) Then .Zip = Left(snpExcel("Zip"), 10)
'            If Len(snpExcel("phone")) = 7 Then
'              .PHONE = Left(snpExcel("phone"), 3) & "-" & Right(snpExcel("phone"), 4)
'            Else
'              If Len(snpExcel("phone")) = 10 Then
'                .PHONE = Left(snpExcel("phone"), 3) & "-" & Mid(snpExcel("phone"), 4, 3) & "-" & Right(snpExcel("phone"), 4)
'              Else
'                If Not IsNull(snpExcel("Phone")) Then .PHONE = Left(snpExcel("phone"), 15)
'             End If
'            End If
'            If Len(snpExcel("fax")) = 7 Then
'              .PHONE2 = Left(snpExcel("fax"), 3) & "-" & Right(snpExcel("fax"), 4)
'            Else
'              If Len(snpExcel("fax")) = 10 Then
'                .PHONE2 = Left(snpExcel("fax"), 3) & "-" & Mid(snpExcel("fax"), 4, 3) & "-" & Right(snpExcel("fax"), 4)
'              Else
'                If Not IsNull(snpExcel("fax")) Then .PHONE2 = Left(snpExcel("fax"), 15)
'             End If
'            End If
'            If Not IsNull(snpExcel("terms")) Then
'              Call addterms(snpExcel("Terms"))
'              .Discount = snpExcel("Terms")
'            End If
'
'            'Call addsalesman(snpExcel("f12"))
'            '.Salesman = snpExcel("f12")
'            .GeoCode = "A"
'            If snpExcel("Sales Tax (California Mill Tax)") = "Tax" Then
'              .taxable = True
'            Else
'              .taxable = False
'            End If
'            '.Reason = snpExcel("Sales Tax Exemption Reason")
'            '.PESTNUM = snpExcel("f12")
'            'If IsDate(snpExcel("Expiration Date")) Then .Pestexp = snpExcel("Expiration Date")
'          End With
'       Call objcust.save(CONN, "A", "", "SQLSERVER")
'       Set objcust = Nothing
'  snpExcel.MoveNext
' Wend
'
' Screen.MousePointer = vbDefault
' MsgBox "Customer Import Complete.", vbOKOnly + vbInformation, Me.Caption
End Sub


Private Sub impacct()
' Dim dbExcel As database
' Dim snpExcel As Recordset
' Dim objacct As clsGLAccounts
' Dim tbl As TableDef
'
'
' Screen.MousePointer = vbHourglass
'
' Dim fld As Field
' Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
' For Each tbl In dbExcel.TableDefs
'   Debug.Print tbl.Name
'    For Each fld In tbl.Fields
'      Debug.Print fld.Name
'    Next
' Next
' Set snpExcel = dbExcel.OpenRecordset("Sheet1$", dbOpenSnapshot)
'
'
'
' While Not snpExcel.EOF
'   Set objacct = New clsGLAccounts
'          With objacct
'            .ID = Right(snpExcel("1000-00"), 10)
'            .Desc = Left(PadTics(snpExcel("Checking - U S National Bank")), 40)
'            .plcat = snpExcel("Asset (Current)")
'            .Debcred = GetDebCredStatus(.plcat)
'          End With
'
'       Call objacct.save(CONN, "A", "", "SQLSERVER")
'       Set objacct = Nothing
'  snpExcel.MoveNext
' Wend
'
' Screen.MousePointer = vbDefault
' MsgBox "acct Import Complete.", vbOKOnly + vbInformation, Me.Caption
End Sub



Private Sub impprod()
' Dim dbExcel As database
' Dim snpExcel As Recordset
' Dim dbagvance As database
' Dim Deptid As String
' Dim ProdID As String
' Dim pName  As String
' Dim selunits As String
' Dim invunits As String
' Dim billdiv As String
' Dim Wt As String
' Dim PackUnit As String
' Dim PackSize As String
' Dim EPA As String
' Dim Manufacturer As String
' Dim Consistency As String
' Dim sql(4) As String
' Dim objGenBus As clsGenBus
' Dim tu As String
' Dim AMTONHAND As String
' Dim last As Double, list As Double, avgcost As Double
' Dim lev1 As Double, lev2 As Double, lev3 As Double
'
'
' Screen.MousePointer = vbHourglass
'
' Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
' Set objGenBus = New clsGenBus
'
'
' Dim fld As Field
' Dim tbl As TableDef
' For Each tbl In dbExcel.TableDefs
'   Debug.Print tbl.Name
'   For Each fld In tbl.Fields
'        Debug.Print fld.Name
'   Next
' Next
'
'
'
'
'
'
'
'
'
'
'
' 'Set snpExcel = dbExcel.OpenRecordset("'Product ID''s to Add To SSI'$", dbOpenSnapshot)
' Set snpExcel = dbExcel.OpenRecordset("ProdLsttxt$", dbOpenSnapshot)
' 'Set snpExcel = dbExcel.OpenRecordset("'Product ID''s Complete AGR List'$", dbOpenSnapshot)
' While Not snpExcel.EOF
'   Call adddept(snpExcel("Dept ID"))
'   Deptid = snpExcel("Dept ID")
'   ProdID = Left(snpExcel("Product ID"), 10)
'   pName = Left(snpExcel("Product Name"), 50)
'   AMTONHAND = snpExcel("Amt On Hand")
'   If Not IsNull(snpExcel("Selling Units")) Then
'     tu = snpExcel("Selling Units")
'   Else
'     tu = "each"
'   End If
'   Call Update_Units(tu)
'   selunits = tu
'   If Not IsNull(snpExcel("Inventory Units")) Then
'     tu = snpExcel("Inventory Units")
'   Else
'     tu = "each"
'   End If
'   Call Update_Units(tu)
'   invunits = tu
'   If Not IsNull(snpExcel("Billing Divisor")) Then
'     billdiv = snpExcel("Billing Divisor")
'   Else
'     billdiv = 1
'   End If
'   list = snpExcel("List Price")
'   avgcost = snpExcel("Average Cost")
'   If IsNull(snpExcel("Weight")) Then
'     Wt = 0
'   Else
'     Wt = snpExcel("Weight")
'   End If
'   If Not IsNull(snpExcel("Pkg Units")) Then
'     tu = snpExcel("Pkg Units")
'   Else
'     tu = "each"
'   End If
'   Call Update_Units(tu)
'   PackUnit = tu
'   PackSize = snpExcel("Pkg Size")
'   last = snpExcel("Last Cost")
'   'If IsNull(snpExcel("EPA#")) Then
'     EPA = ""
'   'Else
'   '  EPA = snpExcel("EPA#")
'   'End If
'   'If snpExcel("f10") <> "" Then
'   '  Manufacturer = "'" & snpExcel("f10") & "'"
'   '  Call addmanu(snpExcel("f10"))
'   'Else
'     Manufacturer = "Null"
'   'End If
'   Consistency = "" 'snpExcel("Consistency")
'   lev1 = snpExcel("Level 1")
'   lev2 = snpExcel("Level 2")
'   lev3 = snpExcel("Level 3")
'   sql(0) = "INSERT INTO  product (DEPARTID, prodid,PRODNAME,MANUFAC,INVUNITS,PackUnit,PackSize,UNITWGHT,eparegnum,dryorliq,AMTONHAND) values ('" & Deptid & "','" & ProdID & "','" & pName & "'," & Manufacturer & ",'" & invunits & "','" & PackUnit & "','" & PackSize & "','" & Wt & "','" & EPA & "','" & Consistency & "','" & AMTONHAND & "') "
'   sql(1) = "INSERT INTO  prodsaf (DEPARTID, prodid) values ('" & Deptid & "','" & ProdID & "') "
'   sql(2) = "INSERT INTO  prodset (DEPARTID, prodid) values ('" & Deptid & "','" & ProdID & "') "
'   sql(3) = "INSERT INTO  prodprce (DEPARTID, prodid, billdiv,BILLUNITS,list, LASTPURCST, AVGCOST, PRCELEV1,PRCELEV2,PRCELEV3) values ('" & Deptid & "','" & ProdID & "','" & billdiv & "','" & selunits & "'," & list & "," & last & "," & avgcost & "," & lev1 & "," & lev2 & "," & lev3 & ") "
'   sql(4) = "INSERT INTO  prodacct (DEPTID, prodid) values ('" & Deptid & "','" & ProdID & "') "
'
'   Call objGenBus.DirectExec(CONN, sql)
'
'   snpExcel.MoveNext
'
' Wend
'
' Set objGenBus = Nothing
'
'
'
' snpExcel.Close: Set snpExcel = Nothing
' dbExcel.Close: Set dbExcel = Nothing
' Screen.MousePointer = vbDefault
' MsgBox "Product Import Complete.", vbOKOnly + vbInformation, Me.Caption
End Sub




Private Sub impVEND()
' Dim dbExcel As database
' Dim snpExcel As Recordset
' Dim dbagvance As database
'
'
' Dim fld As Field
' Dim tbl As TableDef
'
' Screen.MousePointer = vbHourglass
'
' Set dbExcel = DBEngine(0).OpenDatabase(TxtSpreadSheet.TEXT, False, False, "Excel 3.0;")
' For Each tbl In dbExcel.TableDefs
'   Debug.Print tbl.Name
'   For Each fld In tbl.Fields
'        Debug.Print fld.Name
'   Next
' Next
' Set snpExcel = dbExcel.OpenRecordset("sheet1$", dbOpenSnapshot)
' Dim objVendor As clsVendor
' While Not snpExcel.EOF
'     Set objVendor = New clsVendor
'     With objVendor
'       .venid = Left(snpExcel("Vendor ID"), 10)
'       .Name = Left(PadTics(snpExcel("Vendor")), 50)
'       If Not IsNull(snpExcel("Address 1")) Then .addr1 = Left(snpExcel("Address 1"), 50)
'       If Not IsNull(snpExcel("Address 2")) Then .addr2 = Left(snpExcel("Address 2"), 50)
'       If Not IsNull(snpExcel("city")) Then .City = snpExcel("city")
'       If Not IsNull(snpExcel("State")) Then .State = Left(snpExcel("state"), 2)
'       If Not IsNull(snpExcel("zip")) Then .Zip = Trim(snpExcel("zip"))
'       'If snpExcel("f7") <> Null Then
'       '  .Location = snpExcel("f7")
'       'Else
'         .Location = "Main"
'       'End If
'       If Len(snpExcel("phone")) = 7 Then
'         .PHONE = Left(snpExcel("phone"), 3) & "-" & Right(snpExcel("phone"), 4)
'       Else
'         If Len(snpExcel("phone")) = 10 Then
'           .PHONE = Left(snpExcel("phone"), 3) & "-" & Mid(snpExcel("phone"), 4, 3) & "-" & Right(snpExcel("phone"), 4)
'         Else
'           If Not IsNull(snpExcel("phone")) Then .PHONE = Left(snpExcel("phone"), 20)
'        End If
'       End If
'
'       If Len(snpExcel("Alt# Phone")) = 7 Then
'         .PHONE2 = Left(snpExcel("Alt# Phone"), 3) & "-" & Right(snpExcel("Alt# Phone"), 4)
'       Else
'         If Len(snpExcel("Alt# Phone")) = 10 Then
'           .PHONE2 = Left(snpExcel("f10"), 3) & "-" & Mid(snpExcel("Alt# Phone"), 4, 3) & "-" & Right(snpExcel("Alt# Phone"), 4)
'         Else
'           If Not IsNull(snpExcel("Alt# Phone")) Then .PHONE2 = Left(snpExcel("Alt# Phone"), 20)
'        End If
'       End If
'
'       If Len(snpExcel("Fax")) = 7 Then
'         .fax = Left(snpExcel("Fax"), 3) & "-" & Right(snpExcel("Fax"), 4)
'       Else
'         If Len(snpExcel("fax")) = 10 Then
'           .fax = Left(snpExcel("Fax"), 3) & "-" & Mid(snpExcel("Fax"), 4, 3) & "-" & Right(snpExcel("Fax"), 4)
'         Else
'           If Not IsNull(snpExcel("fax")) Then .fax = Left(snpExcel("Fax"), 20)
'        End If
'       End If
'       .actnum = snpExcel("GL Acct #")
'
'
'       'If snpExcel("f12") <> Null Then .web = snpExcel("f12")
'
'     End With
'
'     Call objVendor.save(CONN, "A", "", "SQLSERVER")
'     Set objVendor = Nothing
'
'  snpExcel.MoveNext
' Wend
'
'
'
' snpExcel.Close: Set snpExcel = Nothing
' dbExcel.Close: Set dbExcel = Nothing
' Screen.MousePointer = vbDefault
' MsgBox "Vendor Import Complete.", vbOKOnly + vbInformation, Me.Caption
End Sub


Public Function ZeroPad(le%, thestring$) As String
'le% is the desired length of the zeropadded string
'thestring$ is the original unpadded string
Dim IsNeg As Boolean

If Left$(thestring$, 1) = "-" Then
 le% = le% - 1
 thestring$ = Right$(thestring$, Len(thestring$) - 1)
 IsNeg = True
Else
 IsNeg = False
End If
thestring$ = Trim(thestring$)
ZeroPad = String$(le% - Len(thestring$), "0") + thestring$
If IsNeg Then ZeroPad = "-" + ZeroPad
End Function




Public Sub PARSELINE(PARSELINE$, ResultArray$(), delim As Integer, ARRAY1 As Integer)
 'this sub could be rewritten to use vb 6's new split function
 Dim comma%
 comma% = 0: ARRAY1 = 0
 Do
  comma% = InStr(PARSELINE$, Chr$(delim)) ' position of 1st delimiter
  If comma% > 0 Then  ' found delimiter
    ResultArray$(ARRAY1) = Trim(Left$(PARSELINE$, comma% - 1))  ' get item
    PARSELINE$ = Trim(Mid$(PARSELINE$, comma% + 1))
    ARRAY1 = ARRAY1 + 1
   Else ' found no delimiter
    ResultArray$(ARRAY1) = Trim(PARSELINE$)
    PARSELINE$ = ""
  End If
 Loop While Len(PARSELINE$) ' loop until no more characters
' Array1 is string array index of last returned string element
End Sub


Public Function StripChar(oldstr As String, char2strip As String) As String
'this function could be replaced by vb6's new replace function
Dim t As Integer, TheByte$
Dim le As Integer

le = Len(oldstr)
For t = 1 To le
  TheByte$ = Mid$(oldstr, t, 1)
  If TheByte$ <> char2strip Then StripChar = StripChar & TheByte$
Next t
End Function




Private Sub Form_Load()
  Call initgrid
End Sub


Private Sub grddata_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
   Call ClearGridSelected(Me!grddata)
   grddata.MaxRows = grddata.DataRowCnt
 End If
End Sub

Private Sub TxtSpreadSheet_Change()
 'Call SetButtons
End Sub


Private Sub TxtSpreadSheet_DblClick()
Dim savdir$, SAVDRIVE$

savdir$ = CurDir
If Mid$(savdir$, 2, 1) = ":" Then
 SAVDRIVE$ = Left$(savdir$, 2)
End If

 On Local Error GoTo LeHandle
 With CDIPRINTSET
  .CancelError = True
  .Flags = cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoLongNames + cdlOFNPathMustExist + cdlOFNShareAware
  .DialogTitle = "Select a Customer Excel Spread Sheet"
  .Filename = "*.xls"
  .Filter = "Excel Spread Sheet|*.xls"
  .DefaultExt = "*.xls"
  .ShowOpen
  TxtSpreadSheet.TEXT = .Filename
 End With

   If Mid$(savdir$, 2, 1) = ":" Then
     ChDrive (SAVDRIVE$)
    End If
    
    ChDir (savdir$)

LeHandle:
On Local Error GoTo 0
End Sub


Private Sub TxtSpreadSheet_GotFocus()
 If Len(TxtSpreadSheet.TEXT) < 1 Then
   Call TxtSpreadSheet_DblClick
 End If
End Sub


