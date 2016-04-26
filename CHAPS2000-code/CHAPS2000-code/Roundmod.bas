Attribute VB_Name = "Round_modules"
Option Explicit

Function funround2(decacc%, j) As String
Dim jj$
If IsNull(j) Then
   funround2 = 0
   Exit Function
End If
If j = "" Then
   funround2 = 0
   Exit Function
End If

If decacc% = 0 Then
   If j >= 0 Then j = Int(j + 0.5001001)
   If j < 0 Then j = Int(j + 0.4998999)
   jj$ = Trim$(Str$(j))
   funround2 = jj$
   Exit Function
End If

If decacc% = 1 Then
   If j >= 0 Then j = Int(j * 10 + 0.5001001) / 10
   If j < 0 Then j = Int(j * 10 + 0.4998999) / 10
   jj$ = Str$(j)
   If j = Int(j) Then jj$ = jj$ + ".0"
   jj$ = Trim$(jj$)
   funround2 = jj$
   Exit Function
End If

If decacc% = 2 Then
   If j >= 0 Then j = Int(j * 100 + 0.5001001) / 100
   If j < 0 Then j = Int(j * 100 + 0.4998999) / 100
   jj$ = Str$(j)
   If j = Int(j) Then jj$ = jj$ + ".00"
   If Mid$(jj$, Len(jj$) - 1, 1) = "." Then jj$ = jj$ + "0"
   jj$ = Trim$(jj$)
   funround2 = jj$
   Exit Function
End If
If decacc% = 3 Then
   If j >= 0 Then j = Int(j * 1000 + 0.5001001) / 1000
   If j < 0 Then j = Int(j * 1000 + 0.4998999) / 1000
   jj$ = Str$(j)
   If j = Int(j) Then jj$ = jj$ + ".000"
   If Mid$(jj$, Len(jj$) - 1, 1) = "." Then jj$ = jj$ + "00"
   If Mid$(jj$, Len(jj$) - 2, 1) = "." Then jj$ = jj$ + "0"
   jj$ = Trim$(jj$)
   funround2 = jj$
   Exit Function
End If
If decacc% = 4 Then
   If j >= 0 Then j = Int(j * 10000 + 0.5001001) / 10000
   If j < 0 Then j = Int(j * 10000 + 0.4998999) / 10000
   jj$ = Str$(j)
   If j = Int(j) Then jj$ = jj$ + ".0000"
   If Mid$(jj$, Len(jj$) - 1, 1) = "." Then jj$ = jj$ + "000"
   If Mid$(jj$, Len(jj$) - 2, 1) = "." Then jj$ = jj$ + "00"
   If Mid$(jj$, Len(jj$) - 3, 1) = "." Then jj$ = jj$ + "0"
   jj$ = Trim$(jj$)
   funround2 = jj$
   Exit Function
End If

End Function

Function trunk2(decacc%, j) As String
'this function will trunk a three decimal place number to the number of places sent in decacc%

  Dim jj$
  If IsNull(j) Then
    trunk2 = ""
    Exit Function
  End If
  
  If j = "" Then
    trunk2 = ""
    Exit Function
  End If

  jj$ = funround2(3, j)

  If decacc% = 0 Then
    trunk2 = Left(jj$, Len(jj$) - 4)
    Exit Function
  End If

  If decacc% = 1 Then
    trunk2 = Left(jj$, Len(jj$) - 2)
    Exit Function
  End If

  If decacc% = 2 Then
    trunk2 = Left(jj$, Len(jj$) - 1)
    Exit Function
  End If
End Function


