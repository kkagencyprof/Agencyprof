VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form kbuch 
   Caption         =   "Kassenbuch"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Rechnug"
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abrechnung"
      Height          =   255
      Left            =   9120
      Style           =   1  'Grafisch
      TabIndex        =   16
      ToolTipText     =   "Programm ausdrucken"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ListBox gd1ids 
      Height          =   2955
      IntegralHeight  =   0   'False
      Left            =   7320
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   240
      Top             =   3360
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   120
      Picture         =   "kbuch.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Formular schiessen"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anzeigen:"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ComboBox endlist 
      Height          =   315
      ItemData        =   "kbuch.frx":0250
      Left            =   8160
      List            =   "kbuch.frx":0252
      TabIndex        =   4
      Text            =   "23:59"
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox beglist 
      Height          =   315
      ItemData        =   "kbuch.frx":0254
      Left            =   5400
      List            =   "kbuch.frx":0256
      TabIndex        =   3
      Text            =   "00:00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label t_epb 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Brutto Endbetrag:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Summe MwSt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9600
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label t_epm 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Netto Endbetrag:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9240
      TabIndex        =   11
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label t_epn 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "bis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Von:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "kbuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim posv As Double, posb As Double

Private Sub Command1_Click()
Dim rrr
Dim selt As String, self As String, epreisnetto As Double
Dim epn As Double, epm As Double, epb As Double, r As ADODB.Recordset, c$, dtmp As Double
Dim mwst As Double, lvitem

Dim d2infile As String, d2insub As String
d2infile = "kbuch": d2insub = "Command1_Click"
gd1.ListItems.Clear
gd1ids.Clear

self = datum2sql(trm(Text5.text)) & " " & beglist.text
selt = datum2sql(trm(Text1.text)) & " " & endlist.text
If Len(self & selt) < 32 Then Exit Sub

MousePointer = 11: DoEvents

c$ = "SELECT * " + _
     "FROM kassenbuch " + _
     "WHERE ((dtg>='" & self & "') and (dtg<='" & selt & "')) " + _
     "order by dtg;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
epn = 0: epm = 0: epb = 0
While Not r.EOF
  epn = epn + r!epreisnetto
  epm = epm + r!mwst
  dtmp = r!epreisnetto + r!mwst
  epb = epb + dtmp

  Set lvitem = gd1.ListItems.add(, , "1")
  lvitem.SubItems(1) = r!dtg
  lvitem.SubItems(2) = r!bezeichnung
  lvitem.SubItems(3) = fixeur(r!epreisnetto)
  lvitem.SubItems(4) = fixeur(r!mwst)
  lvitem.SubItems(5) = fixeur(dtmp)
  lvitem.SubItems(6) = r!vorgang
  gd1ids.AddItem r!id
  r.MoveNext
Wend
t_epn.Caption = fixeur(epn)
t_epb.Caption = fixeur(epb)
t_epm.Caption = fixeur(epm)
MousePointer = 0

End Sub

Private Sub Command2_Click()
'd2infile = "kbuch": d2insub = "Command2_Click"
Unload Me
End Sub

Public Sub Command3_Click()
Dim r As ADODB.Recordset
Dim lvitem, i%, V$, b$
Dim self As String
Dim selt As String

Dim d2infile As String, d2insub As String
d2infile = "kbuch": d2insub = "Command3_Click"
Set lvitem = gd1.SelectedItem

For i% = 1 To gd1.ListItems.Count
 Set lvitem = gd1.ListItems(i%)
 If lvitem.Selected = True Then
   If V$ = "" Then V$ = lvitem.SubItems(1)
   b$ = lvitem.SubItems(1)
 End If
Next i%
Text1.text = datfromsql(word1(V$)): beglist.text = word2(V$)
Text5.text = datfromsql(word1(b$)): endlist.text = word2(b$)
Command3.Enabled = False

self = datum2sql(trm(Text5.text)) & " " & beglist.text
selt = datum2sql(trm(Text1.text)) & " " & endlist.text
Call form1.kassenzettel("Kartenverkauf", self, selt)


End Sub

Private Sub Command6_Click()
Dim self As String
Dim selt As String

'd2infile = "kbuch": d2insub = "Command6_Click"
self = datum2sql(trm(Text5.text)) & " " & beglist.text
selt = datum2sql(trm(Text1.text)) & " " & endlist.text
Call form1.kassenabrechnung("Kartenverkauf", self, selt)

End Sub

Private Sub Form_Resize()
'd2infile = "kbuch": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Load()
'd2infile = "kbuch": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Dim i%, rrr, klrv%, r As ADODB.Recordset, c$, dbpara$
Dim colHeader

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
gd1.View = lvwReport

Set colHeader = gd1.ColumnHeaders.add(, , transe("Anzahl"), 700)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum/Zeit"), 1500)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Bezeichnung"), 3000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Netto"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("MwSt"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Brutto"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Vorgang"), 1000)
Text1.text = Date
Text5.text = Date
beglist.Clear
endlist.Clear
For i% = 0 To 23
  c$ = Format$(i%, "0#") & ":00"
  beglist.AddItem c$
  endlist.AddItem c$
Next i%
beglist.text = "00:00"
endlist.text = "23:59"
Command3.Enabled = False
kbuch.Caption = transe("Kassenbuch")
Command3.Caption = transe("Rechnug")
Command6.Caption = transe("Abrechnung")
Command6.ToolTipText = transe("Abrechnung ausdrucken")
Command2.ToolTipText = transe("Formular schliessen")
Command1.Caption = transe("Anzeigen:")
Label8(1).Caption = transe("Brutto Endbetrag:")
Label8(0).Caption = transe("Summe MwSt")
Label8(2).Caption = transe("Netto Endbetrag:")
Label2.Caption = transe("bis:")
Label1.Caption = transe("Von:")
Show
Call Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "kbuch": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub


Public Sub gd1_Click()
Dim id$, aid$, tid$, lvitem, rrr, i%

'd2infile = "kbuch": d2insub = "gd1_Click"
Set lvitem = gd1.SelectedItem
On Error Resume Next
id$ = lvitem.SubItems(6)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

For i% = 1 To gd1.ListItems.Count
 Set lvitem = gd1.ListItems(i%)
 If id$ = lvitem.SubItems(6) Then gd1.ListItems(i%).Selected = True
Next i%
Command3.Enabled = True
Call gd1.SetFocus

End Sub

Private Sub Text1_DblClick()
'd2infile = "kbuch": d2insub = "Text1_DblClick"
  With frmCalendar
    .init Text1, Text1.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1.text = datfromsql(datfromsql(Format(.SelectedDate, "yyyy-mm-dd")))
    End If
  End With
  Unload frmCalendar

End Sub

Private Sub Text5_DblClick()
'd2infile = "kbuch": d2insub = "Text5_DblClick"
  With frmCalendar
    .init Text5, Text5.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text5.text = Format(.SelectedDate, "yyyy-mm-dd")
    End If
  End With
  Unload frmCalendar

End Sub
