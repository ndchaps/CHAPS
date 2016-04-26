Attribute VB_Name = "SSIExecute"
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const WM_CLOSE = &H10

Public Declare Function GetLastError _
Lib "kernel32" () As Long
Public Declare Function FormatMessage _
Lib "kernel32" Alias "FormatMessageA" _
(ByVal dwFlags As Long, _
lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Function LastSystemError() As String
    '
' better system error
    '
Dim sError As String * 500
Dim lErrNum As Long
Dim lErrMsg As Long
    '
lErrNum = GetLastError
lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)
LastSystemError = Trim(sError)
    '
End Function

Public Sub CloseProgram(Program As String)
 Dim A As Long
 Dim hWnd As Long
 Dim anull As String
 hWnd = FindWindow(anull, Program)
 If hWnd <> 0 Then A = SendMessage(hWnd, WM_CLOSE, 0, 0)
End Sub

Public Function LoadProgram(ProgramTitle As String, ProgramPath$) As Long
 Dim retcode As Integer
 On Local Error GoTo LoadProgram
 AppActivate ProgramTitle
 LoadProgram = Maximize(ProgramTitle)
 If LoadProgram = 0 Then GoTo LoadProgram
Exit Function
 
LoadProgram:
 On Local Error GoTo LeHandle
 Screen.MousePointer = vbHourglass
 If ProgramPath$ = "" Then GoTo LeHandle
 If Not fileexist(ProgramPath$) Then GoTo LeHandle
 LoadProgram = Shell(ProgramPath$, 3)
 Screen.MousePointer = vbDefault
Exit Function
 
LeHandle:
 MsgBox ProgramTitle & " is not installed" & vbCrLf & vbCrLf & "It can be purchased through SSI." & vbCrLf & vbCrLf & "For additional information" & vbCrLf & "Please call SSI" & vbCrLf & " phone #: (800)-752-7912", vbOKOnly + vbExclamation, MdiMain.caption
 Screen.MousePointer = vbDefault
 LoadProgram = -1
End Function

Public Function Maximize(thetitle$) As Integer
'pass in the title of a window and this function will show that window
'maximized. uses api routine showwindow.
'Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

 Dim hWnd As Long
 Dim Newhwnd As Long
 Dim anull As String
 hWnd = FindWindow(anull, thetitle$)
 Maximize = hWnd
 If hWnd <> 0 Then
   Newhwnd = ShowWindow(hWnd, SW_SHOWMAXIMIZED)
 End If
End Function

Public Sub Minimize(thetitle$)
'pass in the title of a window and this function will Minimize that window.
'Uses api routine showwindow.
'Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


 Dim hWnd As Long
 Dim anull As String
 hWnd = FindWindow(anull, thetitle$)
 If hWnd <> 0 Then
   hWnd = ShowWindow(hWnd, SW_SHOWMINIMIZED)
 End If
End Sub

