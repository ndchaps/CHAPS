Attribute VB_Name = "Wait"
Option Explicit

Sub WaitSecs(secs)

'*************************************************************
' SUB NAME: WaitSecs
'   written in Visual Basic 3.0
'
' PURPOSE:
'   Waits a specified number of seconds
'
' INPUT PARAMETERS:
'   secs: Number of seconds to wait (ex: 30)
'
' EXAMPLE(S):
'   Call WaitSecs(1)
'*************************************************************

    Dim sTart!, temp%
    sTart! = Timer
    While Timer < sTart! + secs + 1
         temp% = DoEvents()
    Wend
End Sub

