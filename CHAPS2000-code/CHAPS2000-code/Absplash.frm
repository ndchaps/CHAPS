VERSION 5.00
Begin VB.Form AbSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "application title"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "Absplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   6075
      TabIndex        =   7
      Top             =   2340
      Width           =   6135
      Begin VB.Label lblUserInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "user information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "user name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   480
      Top             =   480
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   3735
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3315
      Width           =   1335
   End
   Begin VB.Label LblBornOnDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   17
      Top             =   3300
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1985, 1987, 1988, 1990, 1993, 1999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   16
      Top             =   4290
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "North Datoka State University Extension Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   15
      Top             =   3885
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "North Dakota Beef Cattle Improvement Association"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   14
      Top             =   3705
      Width           =   4695
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   795
      TabIndex        =   1
      Top             =   1785
      Width           =   4815
   End
   Begin VB.Label lblFileDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "file description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licensed to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2085
      Width           =   6135
   End
   Begin VB.Label lblPathEXE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "path and exe information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2910
      Width           =   6135
   End
   Begin VB.Label lblTrademark 
      BackStyle       =   0  'Transparent
      Caption         =   "trademark information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1950
      Width           =   6135
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "copyright information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1710
      Width           =   6135
   End
   Begin VB.Line linDivide 
      Index           =   1
      X1              =   120
      X2              =   6240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line linDivide 
      Index           =   0
      X1              =   120
      X2              =   6240
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   11
      Top             =   3540
      Width           =   4695
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "version information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   75
      Width           =   2445
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "application title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   0
      Top             =   -60
      Width           =   3675
   End
   Begin VB.Image imgIcon 
      Height          =   855
      Left            =   120
      Picture         =   "Absplash.frx":000C
      Stretch         =   -1  'True
      Top             =   795
      Width           =   6120
   End
End
Attribute VB_Name = "AbSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'      Copyright © 1997 Microsoft Corporation.  All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------

Option Explicit

' API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

' API Constants
Private Const GWL_STYLE         As Long = (-16)
Private Const WS_CAPTION        As Long = &HC00000
Private Const WS_CAPTION_NOT    As Long = &HFFFFFFFF - WS_CAPTION

Private Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGKEYSYSINFO    As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Private Const gREGVALSYSINFOLOC As String = "MSINFO"
Private Const gREGVALSYSINFO    As String = "PATH"

' NT location of user name and company
Private Const gNTREGKEYINFO     As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Private Const gNTREGVALUSER     As String = "RegisteredOwner"
Private Const gNTREGVALCOMPANY  As String = "RegisteredOrganization"

' Win95 locataion of user name and company
Private Const g95REGKEYINFO     As String = "Software\Microsoft\MS Setup (ACME)\User Info"
Private Const g95REGVALUSER     As String = "DefName"
Private Const g95REGVALCOMPANY  As String = "DefCompany"

' Change these to what you want the default name and user info to be
Private Const DEFAULT_USER_NAME As String = "USER INFORMATION NOT AVAILABLE"
Private Const DEFAULT_USER_INFO As String = vbNullString

' Information for warning information at bottom of form
Private gWarningInfo      As String

Private mBoxHeight              As Integer
Private mStyle                  As StyleType
Private mTitleBarHidden         As Boolean

' Type declarations
Private Type StyleType
    OldStyle As Long
    NewStyle As Long
End Type 'StyleType

Private Sub Form_Load()
    'lblWarning.caption = gWarningInfo

'   Fill in all of the information that comes from the App object
    With App
        Caption = "About C.H.A.P.S. 2000" ' .TITLE
        lblTitle.Caption = "C.H.A.P.S. 2000" '.TITLE
        
        If .CompanyName <> "" Then
            lblCompanyName.Caption = "A product of " & .CompanyName
        Else
            lblCompanyName.Caption = ""
        End If
        
        lblVersion.Caption = "Version " & .Major & "." & .Minor & "." & _
                             .Revision & " (32-bit)"
        lblCopyright.Caption = .LegalCopyright
        lblTrademark.Caption = .LegalTrademarks
        lblPathEXE.Caption = .Path & "\" & .EXEName & " "
        lblFileDescription.Caption = .FileDescription
        LblBornOnDate.Caption = "Date of " & .EXEName & ".exe is " & FileDateTime(.Path & "\" & .EXEName & ".exe")
         
    End With 'App
    
'   Get "default" height of About Box
    mBoxHeight = Height
End Sub

Private Sub CMDOk_Click()
    Hide ' If you want to unload the form, change this to Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Public Sub About(frmParent As Form, Optional lUserName As String, _
                 Optional lUserInfo As String, Optional Warning As String)
    lblWarning.Caption = Warning
    'imgIcon.Picture = LoadPicture(App.Path & "\cow1.bmp")
    CMDOk.Enabled = True
    cmdSysInfo.Enabled = True
    
'   Add user information to form
    If lUserName <> "" Then
        lblUserName.Caption = lUserName
        lblUserInfo.Caption = lUserInfo
    Else
        lblUserName.Caption = GetUserName
        lblUserInfo.Caption = GetUserCompany
    End If
    
'   Modify the form style to show the title bar
    ShowTitleBar
    
'   A resize event is needed in order to apply the changes to the form style.  Setting
'   the height to 0 should do it.
    If Height = mBoxHeight Then
        Height = 0
    End If
    
'   Set height of About Box to "default" height
    Height = mBoxHeight
    
    Show vbModal, frmParent
End Sub

Public Sub SplashOn(frmParent As Form, Optional MinDisplay As Long, _
                    Optional lUserName As String, Optional lUserInfo As String, Optional Warning As String)
    'gWarningInfo = Warning
    lblWarning.Caption = Warning
    If Not Visible Then
        Dim lHeight As Integer
        
        'imgIcon.Picture = LoadPicture(App.Path & "\cow1.bmp")
        CMDOk.Enabled = False
        cmdSysInfo.Enabled = False
    
'       If a delay is specified, set up the Timer
        If MinDisplay > 0 Then
            Timer1.Interval = MinDisplay
            Timer1.Enabled = True
        End If
        
'       Add user information to form
        If lUserName <> "" Then
            lblUserName.Caption = lUserName
            lblUserInfo.Caption = lUserInfo
        Else
            lblUserName.Caption = GetUserName
            lblUserInfo.Caption = GetUserCompany
        End If
        
'       Modify the form style to hide the title bar
        HideTitleBar
        
'       Need to cause a form resize in order to get updated ScaleHeight value
        'lHeight = Height
        'Height = 0
        'Height = lHeight
        
'       Set height to hide the "About Box Only" information
        'Height = linDivide(1).Y1 + (Height - ScaleHeight)
        
'       Show the form
        Show vbModeless, frmParent

'       For some reason, need a Refresh to make sure Splash Screen gets painted
        Refresh
    End If
End Sub

Public Sub SplashOff()
    If Visible Then
'       Wait until any minimum display time elapses
        Do While Timer1.Enabled
            DoEvents
        Loop
        
        Hide ' If you want to unload the form, change this to Unload Me

'       Modify the form style to show the title bar
        ShowTitleBar
        
'       Set height of About Box to "default" height
        Height = mBoxHeight
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Sub HideTitleBar()
'   Change the style of the form to not show a title bar
    If mTitleBarHidden Then Exit Sub
    
    mTitleBarHidden = True
    
    With mStyle
        .OldStyle = GetWindowLong(hWnd, GWL_STYLE)
        .NewStyle = .OldStyle And WS_CAPTION_NOT
        SetWindowLong hWnd, GWL_STYLE, .NewStyle
    End With 'mStyle
End Sub

Private Sub ShowTitleBar()
'   Change the style of the form to show a title bar
    If Not mTitleBarHidden Then Exit Sub
    mTitleBarHidden = False
    SetWindowLong hWnd, GWL_STYLE, mStyle.OldStyle
End Sub

Private Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existence Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Temporary Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occurred...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function GetUserName() As String
    Dim KeyVal As String
            
'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   None of the above
    Else
        GetUserName = DEFAULT_USER_NAME
    End If
End Function

Private Function GetUserCompany() As String
    Dim KeyVal As String
    
'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   None of the above
    Else
        GetUserCompany = DEFAULT_USER_INFO
    End If
End Function
