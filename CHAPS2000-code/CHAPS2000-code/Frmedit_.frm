VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmedit_data 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Database Data"
   ClientHeight    =   4125
   ClientLeft      =   1740
   ClientTop       =   1515
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpread.vaSpread GrdData 
      Height          =   4005
      Left            =   60
      OleObjectBlob   =   "Frmedit_.frx":0000
      TabIndex        =   2
      Top             =   60
      Width           =   4725
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "Cancel"
      Height          =   385
      Left            =   4875
      TabIndex        =   1
      Top             =   525
      Width           =   1000
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   385
      Left            =   4875
      TabIndex        =   0
      Top             =   60
      Width           =   1000
   End
End
Attribute VB_Name = "frmedit_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private indx$(50), indexnum%
Dim DbName
Private Sub cmdcancel_Click()
 Unload Me
End Sub

Public Sub clear_grid(gridcntrl As Control)
   ' this clears out grid data not including headings
   ' gridcntrl = formname!gridname
   ' formname!gridname = name of grid control
 Dim Col As Long, Row As Long
 Dim ret As Long
 Screen.MousePointer = vbHourglass
 For Col = 1 To gridcntrl.MaxCols
  For Row = 1 To gridcntrl.MaxRows
   'gridcntrl.COL = COL
   'gridcntrl.row = row
   'gridcntrl.Text = ""
   gridcntrl.SetText Col, Row, ""
  Next Row
 Next Col
 Screen.MousePointer = vbDefault
End Sub

Public Sub clear_grid2(gridcntrl As Control, startrow As Long, endrow As Long, startcol As Long, endcol As Long)
 Dim Col As Long, Row As Long
 Dim ret As Long
 Screen.MousePointer = vbHourglass
 For Col = startcol To endcol
  For Row = startrow To endrow
   gridcntrl.SetText Col, Row, ""
  Next Row
 Next Col
 Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSave_Click()
   Dim RS As Recordset
   Dim my_field As Field
'   Dim a As DAO.FieldAttributeEnum

   Dim i
   Dim dbpm As database
   i = 0
   Screen.MousePointer = vbHourglass
   Set dbpm = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%) ', gConnect)
   Set RS = dbpm.OpenRecordset(indx$(1), dbOpenTable)
   Dim my_tabledef As TableDef
   Set my_tabledef = dbpm.TableDefs(indx$(1))
   RS.Index = "primarykey"
   Select Case indexnum% - 1
     Case "1"
        RS.Seek "=", indx$(2)
     Case "2"
        RS.Seek "=", indx$(2), indx$(3)
     Case "3"
        RS.Seek "=", indx$(2), indx$(3), indx$(4)
     Case "4"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5)
     Case "5"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6)
     Case "6"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7)
     Case "7"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8)
     Case "8"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9)
    Case "9"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10)
   Case "10"
        RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10), indx$(11)
   End Select
   i = 0
   If Not RS.NoMatch Then
      RS.Edit
      For Each my_field In my_tabledef.Fields
        i = i + 1
        GrdData.Row = i
        GrdData.Col = 1
       ' For Each a In my_field.Attributes
       '  Debug.Print a.dbautonum
       ' Next
    On Local Error Resume Next
  '    If my_field.Attributes = dbUpdatableField Then
        Select Case my_field.Type
         Case dbDate
          If GrdData.TEXT <> "" Then Call Date2Field(RS(my_field.Name), GrdData.TEXT)
         Case dbLong, dbDouble, dbInteger
          If GrdData.TEXT <> "" Then RS(my_field.Name).Value = Val(GrdData.TEXT)
         Case dbText
          If GrdData.TEXT <> "" Then RS(my_field.Name).Value = GrdData.TEXT
         Case dbBoolean
          If GrdData.TEXT = 1 Then
            RS(my_field.Name).Value = True
           Else
            RS(my_field.Name).Value = False
          End If
        End Select
  '  End If
       Next
       RS.Update
   End If
   RS.Close: Set RS = Nothing
   dbpm.Close: Set dbpm = Nothing
   Screen.MousePointer = vbDefault
   Unload Me
End Sub

Private Sub Form_Activate()
 Dim i%, found%
 Dim dbpm As database
 Dim my_tabledef As TableDef
 Dim RS As Recordset
 Dim my_field As Field
 If Me.Tag = "" Then Exit Sub
 indexnum% = 0
 GrdData.Col = 0
 While Me.Tag <> ""
  indexnum% = indexnum% + 1
  found% = InStr(Me.Tag, Chr$(9))
  indx$(indexnum%) = LTrim$(RTrim$(left$(Me.Tag, found% - 1)))
  Me.Tag = Right$(Me.Tag, Len(Me.Tag) - found%)
 Wend
 Set dbpm = DBEngine(0).OpenDatabase(dbfile, exclusiveyn%, readonlyyn%) ', gConnect)
 Set my_tabledef = dbpm.TableDefs(indx$(1))
 Set RS = dbpm.OpenRecordset(indx$(1), dbOpenTable)
 RS.MoveLast
 GrdData.MaxRows = my_tabledef.Fields.count
 GrdData.MaxCols = 1
 Call clear_grid(Me!GrdData)
 GrdData.Row = 0
 GrdData.SetText 1, 0, "Data"
 RS.Index = "primarykey"
 If Not RS.NoMatch Then
   Select Case indexnum% - 1
    Case "1"
     RS.Seek "=", indx$(2)
    Case "2"
     RS.Seek "=", indx$(2), indx$(3)
    Case "3"
     RS.Seek "=", indx$(2), indx$(3), indx$(4)
    Case "4"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5)
    Case "5"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6)
    Case "6"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7)
    Case "7"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8)
    Case "8"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9)
    Case "9"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10)
    Case "10"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10), indx$(11)
    Case "11"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10), indx$(11), indx$(12)
    Case "12"
     RS.Seek "=", indx$(2), indx$(3), indx$(4), indx$(5), indx$(6), indx$(7), indx$(8), indx$(9), indx$(10), indx$(11), indx$(12), indx$(13)
   
   End Select
 End If
 i% = 0
 For Each my_field In my_tabledef.Fields
  i% = i% + 1
  GrdData.Col = 0
  GrdData.Row = i%
  GrdData.TEXT = my_field.Name
  If Len(my_field.Name) * 125 > GrdData.ColWidth(0) Then
    GrdData.ColWidth(0) = Len(my_field.Name) * 125
  End If
  GrdData.Col = 1
  Select Case my_field.Type
   Case dbBoolean
    GrdData.Row = i%
    GrdData.CellType = 10 'SS_CELL_TYPE_CHECKBOX
   Case dbDate
    GrdData.Row = i%
    GrdData.CellType = 0 'SS_CELL_TYPE_DATE
   Case dbLong
    GrdData.Row = i%
    GrdData.CellType = 3 'SS_CELL_TYPE_INTEGER
   Case dbInteger
    GrdData.Row = i%
    GrdData.CellType = 3 ' SS_CELL_TYPE_INTEGER
   Case dbFloat
    GrdData.Row = i%
    GrdData.CellType = 2 'SS_CELL_TYPE_FLOAT
   Case Else
    GrdData.Row = i%
    GrdData.CellType = 1 'SS_CELL_TYPE_EDIT
    GrdData.TypeEditLen = my_field.size
  End Select
  If RS(my_field.Name).Value <> "" Then
    If GrdData.ColWidth(1) < Len(RS(my_field.Name).Value) * 125 Then
      GrdData.ColWidth(1) = Len(RS(my_field.Name).Value) * 125
      If GrdData.ColWidth(1) > 3060 Then GrdData.ColWidth(1) = 3060
    End If
    GrdData.Col = 1
    GrdData.Row = i%
    Select Case my_field.Type
     Case dbBoolean
      If RS(my_field.Name).Value Then
        GrdData.TEXT = 1
       Else
        GrdData.TEXT = 0
      End If
     Case Else
      GrdData.TEXT = RS(my_field.Name).Value
    End Select
  End If
 Next
 RS.Close: Set RS = Nothing
 dbpm.Close: Set dbpm = Nothing
End Sub


Property Let setDbname(Thedb As String)
 DbName = Thedb
End Property

Private Sub Form_Unload(Cancel As Integer)
 Set frmedit_data = Nothing
End Sub


