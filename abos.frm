VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form abos 
   Caption         =   "Aboverwaltung"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "abos.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Neues Abo anlegen"
      Top             =   2520
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1230
      IntegralHeight  =   0   'False
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "abos.frx":0392
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   3960
      Width           =   375
   End
   Begin VB.ListBox abotermine 
      Height          =   2790
      IntegralHeight  =   0   'False
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
   End
   Begin VB.ListBox aboliste 
      Height          =   2400
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   1440
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "AboPreis"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Räume:"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Plätze geplant:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "abos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nochg As Integer

Public Sub aboliste_Click()
Dim rrr
Dim i As Integer, id$, j%, r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "abos": d2insub = "aboliste_Click"
nochg = 1
abotermine.Clear
Text1.text = ""
Text2.text = ""
i = aboliste.ListIndex
If i < 0 Then Exit Sub
id$ = aboliste.List(i)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
c$ = "select abosproraum,preis from hbabos where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Text1.text = trm(r!abosproraum)
  Text2.text = fixeur(trm("0" + trm(r!preis)))
End If
c$ = "select * from hbabotermine where aboid='" + id$ + "' order by dtg"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  c$ = trm(r!dtg) + " " + trm(r!adrid) + " | " + trm(r!pid) + Space$(80) + "(ID:" + trm(r!id)
  abotermine.AddItem c$
  r.MoveNext
Wend
List1.Clear
c$ = "SELECT hbabos.id, hbabos.Name, hblist.raum, hblist.hid " + _
     "FROM (hbabotermine INNER JOIN hbabos ON hbabotermine.aboid = hbabos.id) INNER JOIN hblist ON (hbabotermine.pid = hblist.pgid) AND (hbabotermine.adrid = hblist.hid) " + _
     "Where (((hbabos.id)='" + id$ + "')) ORDER BY hbabos.id, hblist.raum;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
c$ = ""
While Not r.EOF
  If c$ <> trm(r!hid) + " " + trm(r!raum) Then
    c$ = trm(r!hid) + " " + trm(r!raum)
    List1.AddItem c$
  End If
  r.MoveNext
Wend

nochg = 0

End Sub

Private Sub abotermine_DblClick()
Dim rrr
Dim i As Integer, id$, j%, r As ADODB.Recordset, bn$

Dim d2infile As String, d2insub As String
d2infile = "abos": d2insub = "abotermine_DblClick"
i = abotermine.ListIndex
If i < 0 Then Exit Sub
id$ = abotermine.List(i)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
id$ = "select * from hbabotermine where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, id$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Load splan
  On Error Resume Next
  Call splan.SetFocus
  On Error GoTo 0
  splan.noupd% = 1
  splan.hid.text = r!adrid
  splan.pgid.text = r!pid
  DoEvents
  bn$ = trm(Mid(r!dtg, InStr(r!dtg, " ")))
  splan.beglist.text = bn$
  splan.Text5.text = word1(r!dtg)
  splan.noupd% = 0
  Call splan.Command2_Click
  Call splan.beglist_Change
  For i = 0 To splan.termlist.ListCount - 1
    If splan.Text5.text + " " + splan.beglist.text = splan.termlist.List(i) Then
      splan.termlist_drw% = 0
      splan.termlist.ListIndex = i
      splan.termlist_drw% = 1
      Exit For
    End If
  Next i
End If

End Sub

Private Sub abotermine_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, c$, j%


'd2infile = "abos": d2insub = "abotermine_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  idx% = abotermine.ListIndex
  If idx% < 0 Then Exit Sub
  id$ = abotermine.List(idx%)
  j% = InStr(id$, "(ID:") + 4
  id$ = Mid$(id$, j%)
  c$ = "delete from hbabotermine where id='" + id$ + "'"
  Call form1.sqlqry(c$)
  Call aboliste_Click
End If

End Sub

Private Sub Command1_Click()
'd2infile = "abos": d2insub = "Command1_Click"
Unload Me

End Sub

Private Sub Command25_Click()
Dim c$

'd2infile = "abos": d2insub = "Command25_Click"
c$ = trm(InputBox("Name des neuen Abos:", "Neues Abo anlegen", ""))
If c$ <> "" Then
  c$ = "insert into hbabos (id,name) values('" + form1.newid("hbabos", "id", 6) + "','" + c$ + "')"
  Call form1.sqlqry(c$)
  Call rlist1
End If

End Sub

Private Sub Form_Load()
Dim i

'd2infile = "abos": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
nochg = 1
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Show
Call rlist1
BackColor = form1.cleancolor()

End Sub
Private Sub Form_Resize()
'd2infile = "abos": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub


Private Sub Form_Unload(Cancel As Integer)
'd2infile = "abos": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Sub rlist1()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "abos": d2insub = "rlist1"
  aboliste.Clear
  abotermine.Clear
  c$ = "select * from hbabos order by name"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    aboliste.AddItem trm(r!name) + " (" + trm(r!abosproraum) + " Plätze)" + Space$(80) + "(ID:" + r!id
    r.MoveNext
  Wend
End Sub

Private Sub Text1_Change()
Dim c$, j%, i As Integer, id$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "abos": d2insub = "Text1_Change"
If nochg = 1 Then Exit Sub
i = aboliste.ListIndex
If i < 0 Then Exit Sub
id$ = aboliste.List(i)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
j% = Val(Text1.text)
c$ = "update hbabos set abosproraum=" + trm(j%) + " where id='" + id$ + "'"
Call form1.sqlqry(c$)

End Sub

Private Sub Text2_Change()
Dim c$, j%, i As Integer, id$

'd2infile = "abos": d2insub = "Text2_Change"
If nochg = 1 Then Exit Sub
i = aboliste.ListIndex
If i < 0 Then Exit Sub
id$ = aboliste.List(i)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
c$ = "update hbabos set preis=" + d2db(Text2.text) + " where id='" + id$ + "'"
Call form1.sqlqry(c$)

End Sub
