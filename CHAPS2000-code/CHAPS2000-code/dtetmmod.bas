Attribute VB_Name = "date_time_modules"
Option Explicit

Function converttime(tme$) As String
 Dim HR, AMPM$
 HR = Val(Left$(tme$, 2)): AMPM$ = "AM"
 If HR >= 12 Then AMPM$ = "PM"
 If HR > 12 Then HR = HR - 12
  If HR = 0 Then HR = 12
 converttime$ = LTrim$(Str$(HR)) + Mid$(tme$, 3, 3) + AMPM$
End Function



