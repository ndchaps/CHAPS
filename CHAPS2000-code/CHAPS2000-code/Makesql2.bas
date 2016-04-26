Attribute VB_Name = "Create_SQL_Statments2"
Option Explicit
Public Sub create_sql_selection(thecontrol As control, Col(), FieldVar$(), HMFIELDS%, FORMULA$)
 Dim THOR$, column%, theand$
 Dim t As Integer
 '
 ' this routine will create a select from a list box(microhelp)
 ' to send into the crystal report
 ' example you have selected 2 customers the selection
 ' would look like this
 ' {tbcustomer!custid} = 'WIRBAR' AND {TBCUSTOMER!CUSTID} = 'BAKLAR'
 '
 '
 '''''''''''''''''''''''''
 ' THECONTROL - THE LISTBOX THAT CONTAINS THE TAG INFORMATION EX. FRMSELECT_GROWER_REPORTS!LSTGROWERS
 '
 ' COL()      - A ARRAY OF COLUMNS THAT ARE IN THE LIST BOX THAT YOU WANT THE SELECTION FROMULA ON
 '              IF ONLY ONE PRIMARY KEY THEN THERE WOULD BE ONLY ONE ENTRY
 '              IF MORE THEN ON THEN THERE WOULD BE MORE THEN ONE ENTRY
 '              EXAMPLE PRODUCTS TABLE AND YOU HAVE SELECTED PRODUCT ID'S
 '              COL(1) = 0, COL(2) = 1 WHERE COLUMN ONE IN THE LIST BOX IS THE DEPART ID AND COLUMN 2 IS THE PRODUCTID
 '
 ' FIELDVAR$() - THE FIELD IN THE TABLE THAT THE COLUMN REPRESENTS
 '               EXAMPLE FROM ABOUVE FIELDVAR$(1) WOULD = TBPRODUCT!DEPTID AND FIELDVAR$(2) = TBPRODUCT!PRODID
 '
 ' HMFIELDS%   - IS THE NUMBER OF COLUMNS USED FROM THE LIST BOX THE ABOUVE EXAMPLE HMFIELDS% WOULD = 2
 '
 ' FORMULA$    - THE FORMULA SENT BACK
 ''''''''''''''''''''''''''''''''''''
 THOR$ = ""
 FORMULA$ = ""
 If thecontrol.SelectedCount > 0 Then
   For t = 0 To thecontrol.ListCount - 1
    If thecontrol.Tagged(t) = True Then
      FORMULA$ = FORMULA$ + THOR$
      theand$ = ""
      For column% = 1 To HMFIELDS%
       thecontrol.Col = Col(column%)
       FORMULA$ = FORMULA$ & theand$ & FieldVar$(column%) & " = '" & thecontrol.ColList(t) & "'"
       theand$ = " AND "
      Next column%
      THOR$ = " Or "
    End If
   Next t
 End If
End Sub


Public Sub create_sql_selection_NUM(thecontrol As control, Col(), FieldVar$(), HMFIELDS%, FORMULA$)
 Dim THOR$, column%, theand$
 Dim t As Integer
 '
 ' this routine will create a select from a list box(microhelp)
 ' to send into the crystal report
 ' example you have selected 2 customers the selection
 ' would look like this
 ' {tbcustomer!custid} = 'WIRBAR' AND {TBCUSTOMER!CUSTID} = 'BAKLAR'
 '
 '
 '''''''''''''''''''''''''
 ' THECONTROL - THE LISTBOX THAT CONTAINS THE TAG INFORMATION EX. FRMSELECT_GROWER_REPORTS!LSTGROWERS
 '
 ' COL()      - A ARRAY OF COLUMNS THAT ARE IN THE LIST BOX THAT YOU WANT THE SELECTION FROMULA ON
 '              IF ONLY ONE PRIMARY KEY THEN THERE WOULD BE ONLY ONE ENTRY
 '              IF MORE THEN ON THEN THERE WOULD BE MORE THEN ONE ENTRY
 '              EXAMPLE PRODUCTS TABLE AND YOU HAVE SELECTED PRODUCT ID'S
 '              COL(1) = 0, COL(2) = 1 WHERE COLUMN ONE IN THE LIST BOX IS THE DEPART ID AND COLUMN 2 IS THE PRODUCTID
 '
 ' FIELDVAR$() - THE FIELD IN THE TABLE THAT THE COLUMN REPRESENTS
 '               EXAMPLE FROM ABOUVE FIELDVAR$(1) WOULD = TBPRODUCT!DEPTID AND FIELDVAR$(2) = TBPRODUCT!PRODID
 '
 ' HMFIELDS%   - IS THE NUMBER OF COLUMNS USED FROM THE LIST BOX THE ABOUVE EXAMPLE HMFIELDS% WOULD = 2
 '
 ' FORMULA$    - THE FORMULA SENT BACK
 ''''''''''''''''''''''''''''''''''''
 THOR$ = ""
 FORMULA$ = ""
 If thecontrol.SelectedCount > 0 Then
   For t = 0 To thecontrol.ListCount - 1
    If thecontrol.Tagged(t) = True Then
      FORMULA$ = FORMULA$ + THOR$
      theand$ = ""
      For column% = 1 To HMFIELDS%
       thecontrol.Col = Col(column%)
       FORMULA$ = FORMULA$ & theand$ & FieldVar$(column%) & " = " & thecontrol.ColList(t) & " "
       theand$ = " AND "
      Next column%
      THOR$ = " Or "
    End If
   Next t
 End If
End Sub



Public Sub create_sql_not_selection(thecontrol As control, Col(), FieldVar$(), HMFIELDS%, FORMULA$)
 Dim THOR$, column%, theand$
 Dim t As Integer
 '
 ' this routine will create a select from a list box(microhelp)
 ' to send into the crystal report
 ' example you have selected 2 customers the selection
 ' would look like this
 ' {tbcustomer!custid} = 'WIRBAR' AND {TBCUSTOMER!CUSTID} = 'BAKLAR'
 '
 '
 '''''''''''''''''''''''''
 ' THECONTROL - THE LISTBOX THAT CONTAINS THE TAG INFORMATION EX. FRMSELECT_GROWER_REPORTS!LSTGROWERS
 '
 ' COL()      - A ARRAY OF COLUMNS THAT ARE IN THE LIST BOX THAT YOU WANT THE SELECTION FROMULA ON
 '              IF ONLY ONE PRIMARY KEY THEN THERE WOULD BE ONLY ONE ENTRY
 '              IF MORE THEN ON THEN THERE WOULD BE MORE THEN ONE ENTRY
 '              EXAMPLE PRODUCTS TABLE AND YOU HAVE SELECTED PRODUCT ID'S
 '              COL(1) = 0, COL(2) = 1 WHERE COLUMN ONE IN THE LIST BOX IS THE DEPART ID AND COLUMN 2 IS THE PRODUCTID
 '
 ' FIELDVAR$() - THE FIELD IN THE TABLE THAT THE COLUMN REPRESENTS
 '               EXAMPLE FROM ABOUVE FIELDVAR$(1) WOULD = TBPRODUCT!DEPTID AND FIELDVAR$(2) = TBPRODUCT!PRODID
 '
 ' HMFIELDS%   - IS THE NUMBER OF COLUMNS USED FROM THE LIST BOX THE ABOUVE EXAMPLE HMFIELDS% WOULD = 2
 '
 ' FORMULA$    - THE FORMULA SENT BACK
 ''''''''''''''''''''''''''''''''''''
 THOR$ = ""
 FORMULA$ = ""
 'If thecontrol.SelectedCount > 0 Then
   For t = 0 To thecontrol.ListCount - 1
    If thecontrol.Tagged(t) = False Then
      FORMULA$ = FORMULA$ + THOR$
      theand$ = ""
      For column% = 1 To HMFIELDS%
       thecontrol.Col = Col(column%)
       FORMULA$ = FORMULA$ & theand$ & FieldVar$(column%) & " = '" & thecontrol.ColList(t) & "'"
       theand$ = " AND "
      Next column%
      THOR$ = " Or "
    End If
   Next t
' End If
End Sub

Public Sub query_from_form(theform As Form, FORMULA$)
 Dim indx%
 Dim theand$
 If FORMULA$ = "" Then theand$ = "" Else theand$ = " AND"
 For indx% = 0 To theform.Controls.count - 1
  If TypeOf theform.Controls(indx%) Is TextBox Then
    If theform.Controls(indx%).text <> "" Then
      If InStr(theform.Controls(indx%).text, "*") Or InStr(theform.Controls(indx%).text, "?") Then
        FORMULA$ = FORMULA$ & theand$ & " " & theform.Controls(indx%).tag & " like '" & theform.Controls(indx%).text & "'"
       Else
        FORMULA$ = FORMULA$ & theand$ & " " & theform.Controls(indx%).tag & " = '" & theform.Controls(indx%).text & "'"
      End If
      theand$ = " AND"
    End If
  End If
  If TypeOf theform.Controls(indx%) Is ListBox Then
  End If
  If TypeOf theform.Controls(indx%) Is ComboBox Then
    If theform.Controls(indx%).ListIndex <> -1 Then
      If UCase$(theform.Controls(indx%).text) <> "NONE" Then
        If theform.Controls(indx%).tag = "grower.priceing" Then
          FORMULA$ = FORMULA$ & theand$ & " " & theform.Controls(indx%).tag & " = " & theform.Controls(indx%).ListIndex
          theand$ = " AND"
        Else
          FORMULA$ = FORMULA$ & theand$ & " " & theform.Controls(indx%).tag & " = '" & theform.Controls(indx%).text & "'"
          theand$ = " AND"
        End If
      End If
    End If
  End If
  If TypeOf theform.Controls(indx%) Is CheckBox Then
    If theform.Controls(indx%).Value Then
      FORMULA$ = FORMULA$ & theand$ & " " & theform.Controls(indx%).tag
      theand$ = " and"
    End If
  End If
 Next indx
' sql$ = "Select * from grower into "
End Sub

Public Sub create_sql_selectionlp(thecontrol As control, Col(), FieldVar$(), HMFIELDS%, FORMULA$)
 Dim THOR$, column%, theand$
 Dim t As Integer
 '
 ' this routine will create a select from a list box(ListPro)
 ' for a ms access sql stmt
 ' example you have selected 2 customers the selection
 ' would look like this
 ' (tbcustomer!custid) = 'WIRBAR' AND (TBCUSTOMER!CUSTID) = 'BAKLAR'
 '
 '
 '''''''''''''''''''''''''
 ' THECONTROL - THE LISTBOX THAT CONTAINS THE TAG INFORMATION EX. FRMSELECT_GROWER_REPORTS!LSTGROWERS
 '
 ' COL()      - A ARRAY OF COLUMNS THAT ARE IN THE LIST BOX THAT YOU WANT THE SELECTION FROMULA ON
 '              IF ONLY ONE PRIMARY KEY THEN THERE WOULD BE ONLY ONE ENTRY
 '              IF MORE THEN ON THEN THERE WOULD BE MORE THEN ONE ENTRY
 '              EXAMPLE PRODUCTS TABLE AND YOU HAVE SELECTED PRODUCT ID'S
 '              COL(1) = 0, COL(2) = 1 WHERE COLUMN ONE IN THE LIST BOX IS THE DEPART ID AND COLUMN 2 IS THE PRODUCTID
 '
 ' FIELDVAR$() - THE FIELD IN THE TABLE THAT THE COLUMN REPRESENTS
 '               EXAMPLE FROM ABOUVE FIELDVAR$(1) WOULD = TBPRODUCT!DEPTID AND FIELDVAR$(2) = TBPRODUCT!PRODID
 '
 ' HMFIELDS%   - IS THE NUMBER OF COLUMNS USED FROM THE LIST BOX THE ABOUVE EXAMPLE HMFIELDS% WOULD = 2
 '
 ' FORMULA$    - THE FORMULA SENT BACK
 ''''''''''''''''''''''''''''''''''''
 THOR$ = ""
 FORMULA$ = ""
 If thecontrol.SelCount > 0 Then
   For t = 0 To thecontrol.SelCount
    thecontrol.row = thecontrol.NextSel()
    If thecontrol.row <> -1 Then
      FORMULA$ = FORMULA$ + THOR$
      theand$ = ""
      For column% = 1 To HMFIELDS%
       thecontrol.Col = Col(column%)
       FORMULA$ = FORMULA$ & theand$ & FieldVar$(column%) & " = '" & thecontrol.ColList & "'"
       theand$ = " AND "
      Next column%
      THOR$ = " Or "
    End If
   Next t
 End If
End Sub
