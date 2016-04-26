Attribute VB_Name = "FileModules"
Option Explicit

Declare Function DiskSpaceFree Lib "VB5STKIT.DLL" Alias "DISKSPACEFREE" () As Long

Public Function get_free_space(drive$) As Long
 Dim SaveDrive$
 SaveDrive$ = Left$(CurDir$, 1)
 ChDrive Left$(drive$, 1)
 get_free_space = DiskSpaceFree()
 ChDrive SaveDrive$
End Function

Function FileExist(Filename$) As Boolean
 Dim buffer As Integer
 On Local Error Resume Next
 buffer = FreeFile
 Open Filename$ For Input As #buffer
 If Err = 53 Then Close buffer: FileExist = False: Exit Function
 If LOF(buffer) < 1 Then Close buffer: FileExist = False: Exit Function
 Close buffer
 FileExist = True
End Function

Public Function GetPath(Filename As String) As String
 GetPath = CurDir$
 Dim t As Integer
 For t = Len(Filename) To 1 Step -1
  If Mid$(Filename, t, 1) = "\" Or Mid$(Filename, t, 1) = ":" Then
    GetPath = Left$(Filename, t)
    Exit For
  End If
 Next t
End Function
Public Function GetFile(Filename As String) As String
 GetFile = Filename
 Dim t As Integer
 For t = Len(Filename) To 1 Step -1
  If Mid$(Filename, t, 1) = "\" Or Mid$(Filename, t, 1) = ":" Then
    GetFile = Mid$(Filename, t + 1)
    Exit For
  End If
 Next t
End Function


