Attribute VB_Name = "Error_Module_Routines"
Option Explicit
Public GMODNAME$
Public TEXT$(10)
Public GERRNUM$
Public GERRSOURCE$

Public Function WorkSpaceOpenObjects(theWorkspace As Integer) As String
Dim i%, j%, Text$
'theWorkSpace is the workspace collections index, usually 0
For i% = 0 To DBEngine(theWorkspace).Databases.count - 1
  Text$ = Text$ & DBEngine(theWorkspace).Databases(i%).name & vbCrLf
  For j% = 0 To DBEngine(theWorkspace).Databases(i%).Recordsets.count - 1
   Text$ = Text$ & vbTab & DBEngine(theWorkspace).Databases(i%).Recordsets(j%).name & vbCrLf
  Next j%
Next i%
WorkSpaceOpenObjects = Text$
MsgBox WorkSpaceOpenObjects, vbInformation, "Open WorkSpace(" & theWorkspace & ") Objects"
End Function

Sub POP_ERROR(TEXT$())
 Dim T
 Load frmpop_error
 frmpop_error!lblmessage.caption = ""
 For T = 1 To 5
  If TEXT$(T) <> "" Then frmpop_error!lblmessage.caption = frmpop_error!lblmessage.caption & TEXT$(T) & Chr$(13) & Chr$(10)
 Next T
 frmpop_error.Show vbModal
End Sub

