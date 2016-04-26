Attribute VB_Name = "ChapsDemo"
Option Explicit
Public gIsRegistered As Boolean
Public gIsDemo As Boolean

Function GetActivationKey() As String
Dim indx%
Dim sKey$
sKey = Space(80)
indx = GetPrivateProfileString("Chaps", "Key", "", sKey, Len(sKey), "chaps.ini")   '
sKey = Mid(sKey, 1, indx)
GetActivationKey = sKey
End Function

Private Function GetCheckSum(sKey$) As String
Dim indx%
Dim sChar$
Dim sRet&
For indx = 1 To Len(sKey)
   sChar = Mid(sKey, indx, 1)
   sRet = Val(sRet) + GetNumber(sChar)
Next
GetCheckSum = Val(sRet)
End Function

Private Function GetNumber(sChar$) As Integer
'Select Case UCase(sChar)
'   Case "A"
'      GetNumber = 26
'   Case "B"
'      GetNumber = 25
'   Case "C"
'      GetNumber = 24
'   Case "D"
'      GetNumber = 23
'   Case "E"
'      GetNumber = 22
'   Case "F"
'      GetNumber = 21
'   Case "G"
'      GetNumber = 20
'   Case "H"
'      GetNumber = 19
'   Case "I"
'      GetNumber = 18
'   Case "J"
'      GetNumber = 17
'   Case "K"
'      GetNumber = 16
'   Case "L"
'      GetNumber = 15
'   Case "M"
'      GetNumber = 14
'   Case "N"
'      GetNumber = 13
'   Case "O"
'      GetNumber = 12
'   Case "P"
'      GetNumber = 11
'   Case "Q"
'      GetNumber = 10
'   Case "R"
'      GetNumber = 9
'   Case "S"
'      GetNumber = 8
'   Case "T"
'      GetNumber = 7
'   Case "U"
'      GetNumber = 6
'   Case "V"
'      GetNumber = 5
'   Case "W"
'      GetNumber = 4
'   Case "X"
'      GetNumber = 3
'   Case "Y"
'      GetNumber = 2
'   Case "Z"
'      GetNumber = 1
'   Case Else
'      GetNumber = Val(sChar)
GetNumber = Asc(sChar)
'End Select
End Function

Function IsValidKey(sKey As String) As Boolean
Dim sSuffix$
Dim indx%
Dim sPrefix$
IsValidKey = True
For indx = Len(sKey) To 1 Step -1
   If Mid(sKey, indx, 1) = "." Then Exit For
Next
If indx = 0 Then IsValidKey = False: Exit Function
sSuffix = Mid(sKey, indx + 1)
sPrefix = Left(sKey, indx - 1)
If Val(sSuffix) = Val(GetCheckSum(sPrefix)) Then 'compare checksums
   IsValidKey = True
   WritePrivateProfileString "Chaps", "Key", sKey, "chaps.ini"
Else
   IsValidKey = False
   Exit Function
End If
If sKey = "EJCRU422.529" Then
   gIsDemo = True
Else
   gIsDemo = False
End If
End Function

Function IsRegistered() As Boolean
IsRegistered = False
If IsValidKey(GetActivationKey) = True Then
   IsRegistered = True
   gIsRegistered = True
Else
   IsRegistered = False
   gIsRegistered = False
End If
End Function

Function IsValidSireEntry() As Boolean
Dim DB As database, RS As Recordset
IsValidSireEntry = True
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select count(sireid) as countsireid from sireprof", dbOpenSnapshot)
If Not RS.EOF Then
   If Field2Num(RS!countsireid) > 10 Then
      IsValidSireEntry = False
   End If
End If
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End Function

Function IsValidCowEntry() As Boolean
Dim DB As database, RS As Recordset
IsValidCowEntry = True
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select count(cowid) as countsireid from cowprof", dbOpenSnapshot)
If Not RS.EOF Then
   If Field2Num(RS!countsireid) > 10 Then
      IsValidCowEntry = False
   End If
End If
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End Function

Function IsValidCalfEntry() As Boolean
Dim DB As database, RS As Recordset
IsValidCalfEntry = True
Set DB = DBEngine(0).OpenDatabase(dbfile$, exclusiveyn%, readonlyyn%)
Set RS = DB.OpenRecordset("select count(calfid) as countsireid from calfbirth", dbOpenSnapshot)
If Not RS.EOF Then
   If Field2Num(RS!countsireid) > 50 Then
      IsValidCalfEntry = False
   End If
End If
RS.Close: Set RS = Nothing
DB.Close: Set DB = Nothing
End Function
