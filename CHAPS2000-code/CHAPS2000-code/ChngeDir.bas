Attribute VB_Name = "FileModulesforChangeDir"
Option Explicit
Private Const BFFM_INITIALIZED = 1
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                 (ByVal hWnd As Long, ByVal wMsg As Long, _
                  ByVal wParam As Long, lParam As Any) As Long

Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Public Function BrowseCallbackProc(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
  Select Case uMsg
    Case BFFM_INITIALIZED
      ' Set the dialog's pre-selected folder from the pointer to the path
      ' we allocated in bi.lParam above (passed in the lpData param).
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal StrFromPtrA(lpData))
  End Select
End Function

Private Function StrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  StrFromPtrA = sRtn
End Function

