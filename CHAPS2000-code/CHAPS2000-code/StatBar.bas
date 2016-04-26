Attribute VB_Name = "Status_Bar_Routines"
Option Explicit

Public Function DisplayStatus(N, message$) As String
'*************************************************
' Programmer: Lynn Owens
' Date Created: 03-24-1997
'
' Modified Dates:
'
' Description: This routine will Display a status
'  on the N th panel of mdimain!statusbar1
'
' Paramerters In:
'  Message$: the message to display on the display
'   area
'
' Returns:  The Message the is currently in the
'  Display area
'
'
' Example:
' this example will display "Hello World" in the display panel
' SaveMessage$ = DisplayStatus("Hello World")
'
'******************************************************************************
 DisplayStatus = MdiMain!StatusBar1.Panels(N).TEXT
 MdiMain!StatusBar1.Panels(N).TEXT = message$
 MdiMain!StatusBar1.Refresh
 'Call mdimain_paint
 'MDIMAIN!StatusBar1.SimpleText = message$
End Function


