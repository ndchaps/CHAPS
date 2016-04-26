Attribute VB_Name = "Datemod"
Option Explicit
Public Sub CboTimeLoad(thecbo As control, inc%)
Dim hr%, m%, h%, thehour$, theminute$, theampm$

'inc% = send this in, it can be 2,3,4,5,6,10,15,20, or 30 minute incs.
'thecbo is the name of the combo box to load with time selections

If inc% > 30 Then inc% = 30
For h% = 0 To 23
 hr% = h%
 If h% < 12 Then
  theampm$ = "AM"
 Else
  hr% = h% - 12
  theampm$ = "PM"
 End If
 For m% = 0 To 59 Step inc%
  thehour$ = format$(hr%, "##")
  If Val(thehour$) = 0 Then thehour$ = "12"
  theminute$ = format$(m%, "##")
  If Len(theminute$) < 2 Then theminute$ = "0" & theminute$
  If Val(theminute$) = 0 Then theminute$ = "00"
  'thetime$ = thehour$ & ":" & theminute$ & theampm$
  thecbo.AddItem thehour$ & ":" & theminute$ & " " & theampm$
 Next m%
Next h%
  thecbo.ListIndex = 0
End Sub


Public Function JDate()
  Dim TheDate As Date ' Declare variables.
  Dim FirstDate As Date
  TheDate = "12/31/" & DatePart("yyyy", Now)
  FirstDate = "01/01/1998"
  JDate = (DateDiff("y", FirstDate, Now)) + 1
End Function

Public Function GTime()
  GTime = Right$(DatePart("yyyy", Now), 2) & JDate & left$(format(Time, "short Time"), 2) & Right$(format(Time, "short Time"), 2) & "00.00"
End Function

Public Function SSIDate(datein$) As String
  'This function is to make sure a date is in the form mm/dd/yyyy
  'datein$ = any valid date
  'ssidate = the date in the form mm/dd/yyyy
  
  Dim mnth$, dy$, yr$
  If Mid$(datein$, 3, 1) = "/" And Mid$(datein$, 6, 1) = "/" Then
    SSIDate = datein$
    Exit Function
  End If
  mnth$ = month(datein$)
  dy$ = day(datein$)
  yr$ = year(datein$)
  If Len(mnth$) = 1 Then mnth$ = "0" & mnth$
  If Len(dy$) = 1 Then dy$ = "0" & dy$
  
  SSIDate = mnth$ & "/" & dy$ & "/" & yr$
End Function

Public Sub GetDate(gcaldate)
'sub used by all cmdCal click events
'Input: gcaldate which should be the maskedbox.text associated
'       with the cmdCal clicked.
'Output: gcaldate is set with the calendar selected date
'       so cmdCal will set maskedbox.text=gcaldate the line
'       after the call to this sub.
'
'        Created 5/15/97  Mark
'
 If gcaldate = "--/--/----" Then
   gcaldate = SSIDate(Date)
 End If
 frmcalendar.Show vbModal
End Sub

