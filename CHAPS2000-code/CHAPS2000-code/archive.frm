VERSION 5.00
Object = "{B11ECDA8-C130-11CE-9BE9-00AA00575482}#1.0#0"; "MHLIST32.OCX"
Begin VB.Form FrmArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archive Database"
   ClientHeight    =   2760
   ClientLeft      =   4050
   ClientTop       =   1470
   ClientWidth     =   6195
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   6195
   Begin MhglbxLib.Mh3dList LSTFiles 
      Height          =   2295
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   5100
      _Version        =   65536
      _ExtentX        =   8996
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Mh3dList1"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TintColor       =   16711935
      Caption         =   "Mh3dList1"
      ColTitleButtons =   0   'False
      BevelStyleInner =   0
      BevelSizeInner  =   0
      BorderType      =   1
      BorderColor     =   0
      Case            =   0
      Col             =   2
      ColCharacter    =   9
      ColScale        =   2
      ColSizing       =   0
      DividerStyle    =   0
      FillColor       =   16777215
      FontStyle       =   0
      LightColor      =   16777215
      MultiSelect     =   0
      PictureHeight   =   0
      PictureWidth    =   0
      AdjustHeight    =   0
      ScrollBars      =   1
      ShadowColor     =   8421504
      WallPaper       =   0
      Sorted          =   0   'False
      TextColor       =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      ColInstr        =   0
      TitleHeight     =   0
      TitleFontBold   =   0   'False
      TitleFontItalic =   0   'False
      TitleFontName   =   "MS Sans Serif"
      TitleFontSize   =   8.25
      TitleFontStrike =   0   'False
      TitleFontUnder  =   0   'False
      TitleFontStyle  =   0
      TitleBevelStyle =   0
      TitleBevelSize  =   0
      TitleColor      =   0
      FocusColor      =   0
      HighColor       =   16777215
      VirtualList     =   0   'False
      BufferSize      =   100
      SortOrder       =   ""
      SelectedColor   =   8388608
      Transparent     =   0   'False
      TransparentColor=   0
      TitleFillColor  =   12632256
      Platform        =   0
      FireDrawItem    =   0   'False
      DrawItemLeft    =   0
      DrawItemRight   =   0
      DataSourceList  =   ""
      ListDividersH   =   -1  'True
      ListDividersV   =   -1  'True
      TitleDividers   =   -1  'True
      DataField       =   ""
      DataFieldCount  =   0
      ColTitle0       =   "Filename"
      ColWidth0       =   18
      ColTitle1       =   "Date"
      ColWidth1       =   25
      ColTitle2       =   "Size"
      ColWidth2       =   15
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "Delete"
      Height          =   385
      Left            =   5190
      TabIndex        =   5
      Top             =   885
      Width           =   1000
   End
   Begin VB.CommandButton CMDArchive 
      Caption         =   "Archive"
      Height          =   385
      Left            =   5190
      TabIndex        =   4
      Top             =   480
      Width           =   1000
   End
   Begin VB.CommandButton CMDDone 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   385
      Left            =   5175
      TabIndex        =   1
      Top             =   2355
      Width           =   1000
   End
   Begin VB.Label LBLCurrent 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   2175
      TabIndex        =   3
      Top             =   60
      Width           =   3150
   End
   Begin VB.Label Label1 
      Caption         =   "Current Database"
      Height          =   225
      Left            =   825
      TabIndex        =   2
      Top             =   75
      Width           =   1290
   End
End
Attribute VB_Name = "FrmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sKey$
Dim sArchiveDir$
Dim iLastIndx%

Private Sub CMDArchive_Click()
Dim sIndx$
Dim iLen%
CMDArchive.Enabled = False
On Error GoTo ErrHandler
sIndx = Space(80)
iLen = GetPrivateProfileString("Chaps", "Last Saved", "1", sIndx, Len(sIndx), "chaps.ini")
sIndx = Left(sIndx, iLen)
If Val(sIndx) = 1000 Then sIndx = 1
FileCopy dbfile$, sArchiveDir$ & sKey$ & "." & sIndx
WritePrivateProfileString "Chaps", "Last Saved", CStr(sIndx + 1), "chaps.ini"
Call ScanDir
CMDArchive.Enabled = True
Exit Sub
ErrHandler:
MsgBox Err.Description
CMDArchive.Enabled = True
End Sub

Private Sub CmdDelete_Click()
Dim iRet%
CMDDelete.Enabled = False
On Error GoTo ErrHandler
iRet = MsgBox("Are you sure you want to delete this file?", vbYesNo, Me.Caption)
If iRet = vbNo Then
    CMDDelete.Enabled = True
   Exit Sub
End If
LSTFiles.Col = 0
Kill sArchiveDir$ & LSTFiles.ColText
Call ScanDir
CMDDelete.Enabled = True
Exit Sub
ErrHandler:
MsgBox Err.Description
CMDDelete.Enabled = True
End Sub

Private Sub CmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim indx%
Call centermdiform(Me, mdimain, 0, 0)
sKey$ = GetActivationKey
LBLCurrent.Caption = dbfile$
For indx = Len(sKey$) To 1 Step -1
   If Mid(sKey$, indx, 1) = "." Then Exit For
Next
sKey$ = Mid(sKey$, 1, indx - 1)
sArchiveDir$ = dbdir$
Call ScanDir
End Sub

Private Sub ScanDir()
Dim indx%
Dim sFileTime$, sFileSize$
Dim iLastFound%
LSTFiles.ScreenUpdate = False
LSTFiles.Clear
For indx = 1 To 999
   If FileExist(sArchiveDir & sKey & "." & indx) Then
      sFileTime = FileDateTime(sArchiveDir & sKey & "." & indx)
      sFileSize = FileLen(sArchiveDir & sKey & "." & indx)
      LSTFiles.AddItem sKey & "." & indx & vbTab & sFileTime & vbTab & sFileSize
   End If
Next
If LSTFiles.ListCount > 0 Then LSTFiles.ListIndex = 0
LSTFiles.ScreenUpdate = True
End Sub
