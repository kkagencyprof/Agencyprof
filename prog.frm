VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form prog 
   Caption         =   "Programme - AgencyProf"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   Icon            =   "prog.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Publishers"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   49
      ToolTipText     =   "Neues Programm erstellen"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GEMA"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   48
      ToolTipText     =   "Neues Programm erstellen"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1560
      Picture         =   "prog.frx":000C
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Löschen"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Programm kopieren"
      Height          =   255
      Left            =   3600
      TabIndex        =   47
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ersetzen"
      Height          =   255
      Left            =   7200
      TabIndex        =   46
      ToolTipText     =   "Programmpunkt eine Position nach unten setzen"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Timer stimer 
      Left            =   3120
      Top             =   360
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   45
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2400
      Picture         =   "prog.frx":04FC
      Style           =   1  'Grafisch
      TabIndex        =   43
      ToolTipText     =   "Programm in die Zwischenablage kopieren"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Picture         =   "prog.frx":0A2E
      Style           =   1  'Grafisch
      TabIndex        =   41
      ToolTipText     =   "Zeige To Do-Liste"
      Top             =   3240
      Width           =   375
   End
   Begin VB.ComboBox oksel 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   600
      TabIndex        =   40
      Top             =   3660
      Width           =   2775
   End
   Begin VB.CommandButton Command19 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   39
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CheckBox timpaleft 
      Height          =   255
      Left            =   9480
      TabIndex        =   37
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Bühnenplan"
      Height          =   255
      Left            =   9480
      TabIndex        =   36
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8640
      Picture         =   "prog.frx":1BA8
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Besetzung löschen"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   11280
      Top             =   120
   End
   Begin VB.ComboBox besetz 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   3600
      TabIndex        =   32
      Top             =   4080
      Width           =   5055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Name ändern"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      ToolTipText     =   "Einen neuen Namen für ein ausgehähltes Programm vergeben"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "gespielt?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   30
      Top             =   360
      Width           =   1335
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3840
      Top             =   5280
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      Picture         =   "prog.frx":2098
      Style           =   1  'Grafisch
      TabIndex        =   25
      ToolTipText     =   "Speichern"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox chgs 
      Height          =   645
      Left            =   4560
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox List3 
      Height          =   3375
      Left            =   9360
      TabIndex        =   23
      ToolTipText     =   "In welchen Programmen wurde ein Werk gespielt?"
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   8640
      Picture         =   "prog.frx":243F
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Markierten Programmpunkt löschen"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "prog.frx":292F
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Programm ausdrucken"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ab"
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      ToolTipText     =   "Programmpunkt eine Position nach unten setzen"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "auf"
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      ToolTipText     =   "Programmpunkt eine Position nach oben setzen"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      ToolTipText     =   "Programmpunkt hinzufügen"
      Top             =   360
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   3600
      TabIndex        =   17
      ToolTipText     =   "Programmablauf"
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   15
      Text            =   "Text1"
      ToolTipText     =   "Titel der Veranstaltung"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   13
      Text            =   "Text1"
      ToolTipText     =   "Uhrzeit Ende der Veranstaltung"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   11
      Text            =   "Text1"
      ToolTipText     =   "Uhrzeit Beginn"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   9
      Text            =   "Text1"
      ToolTipText     =   "Enddatzm der Veranstaltung"
      Top             =   6495
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Text            =   "Text1"
      ToolTipText     =   "Anfangsdatum der Veranstaltung"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "Ort der Veranstaltung"
      Top             =   5745
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
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
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Neues Programm erstellen"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "prog.frx":2DCF
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Schliessen"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Gewähltes Programm"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Vorhandene Programme"
      Top             =   840
      Width           =   3255
   End
   Begin VB.CheckBox mbes 
      Caption         =   "mit Besetzung"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Suche: "
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm mit Doppelklick auswählen"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Schlagw. li."
      Height          =   255
      Left            =   9480
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label maxbesetz 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3600
      TabIndex        =   35
      ToolTipText     =   "Aus der Maximalbesetzung können Sie einen Bühnenplan erstellen"
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Werk in Programmen"
      Height          =   375
      Left            =   9360
      TabIndex        =   29
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmpunkte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3720
      TabIndex        =   28
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Markierter Part"
      Height          =   255
      Left            =   7440
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   16
      Top             =   6135
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   14
      Top             =   6855
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   12
      Top             =   6855
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   10
      Top             =   6510
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Top             =   6495
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   720
      Shape           =   4  'Gerundetes Rechteck
      Top             =   5640
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4815
      Left            =   3480
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4815
      Left            =   9240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "prog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nflds As Integer, prv$, neuwerkid$, callbck$, okselmi%
Dim tponly$, prgprvn$, snotb41 As Double, usuchvz As Double, msec As Double

Private Sub besetz_Click()
Dim kid$, p%, bid$, wid$, c$

kid$ = besetz.List(besetz.ListIndex)
p% = InStr(kid$, "ID:")
If p% = 0 Then Exit Sub
prog.BackColor = form1.dirtycolor()
DoEvents
bid$ = Mid$(kid$, p% + 3)

kid$ = List2.List(List2.ListIndex)
wid$ = Mid$(kid$, InStr(kid$, "(MYID:") + 6)
wid$ = trm(Left$(wid$, InStr(wid$, " ") - 1))

c$ = "update programmliste set besetztid='" & bid$ & "' where id='" & wid$ & "'"
Call form1.sqlqry(c$)
Timer1.Interval = 100
Timer1.Enabled = True

End Sub

Private Sub Command1_Click()

Hide
Unload prog

End Sub


Private Sub Command10_Click()
Dim r As ADODB.Recordset, cmd$, rrr
Dim kid$

Dim d2infile As String, d2insub As String
d2infile = "prog": d2insub = "Command10_Click"

If List2.ListIndex < 0 Then Exit Sub

Label4.Caption = "Werk in Programmen"
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub
kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
cmd$ = "select programmid from programmliste where werkid='" + kid$ & "'"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If r.EOF Then
  MsgBox "Keine Programme gefunden."
  Exit Sub
End If
r.MoveFirst
List3.Clear
While Not r.EOF
  List3.AddItem r!programmid
  r.MoveNext
Wend

Call List2_DblClick
DoEvents
Call werkvz.Command10_Click
DoEvents
DoEvents
werkvz.Hide

End Sub

Private Sub Command11_Click()
prgprvn$ = Text1(0).text
Text1(0).Enabled = True
Call Text1(0).SetFocus
Command8.Enabled = True

End Sub

Private Sub Command12_Click()
Dim kid$, wid$, c$

kid$ = List2.List(List2.ListIndex)
wid$ = Mid$(kid$, InStr(kid$, "(MYID:") + 6)
wid$ = trm(Left$(wid$, InStr(wid$, " ") - 1))
c$ = "update programmliste set besetztid='' where id='" & wid$ & "'"
Call form1.sqlqry(c$)
Call List2_Click

End Sub

Private Sub Command13_Click()
Dim ProgID$

ProgID$ = Text1(0).text
Load bplan
On Error Resume Next
Call bplan.SetFocus
On Error GoTo 0
bplan.pgid.text = ProgID$

End Sub

Private Sub Command14_Click()
Dim oid$, auftrid$

oid$ = oksel.text

If oid$ <> "" Then
  prog.MousePointer = 11
  auftrid$ = ""
  If InStr(oid$, "(ID:") = 0 Then
    auftrid$ = ""
  Else
    auftrid$ = Mid$(oid$, InStr(oid$, "(ID:") + 4)
  End If
  If auftrid$ <> "" Then
    MousePointer = 11

    Unload auftritt
    DoEvents
    Load auftritt
    Call auftritt.SetFocus
    Call auftritt.showrec(auftrid$, 0)
    MousePointer = 0
  End If
End If

End Sub

Private Sub Command15_Click()
Dim cmd$, rtmp As QueryDef, up$, werkid$, ProgID$
Dim stmp As ADODB.Recordset, ask%, i%, id$

i% = List2.ListIndex
If i% < 0 Then Exit Sub

id$ = List2.List(i%)
If InStr(id$, "(WID:") = 0 Then Exit Sub
id$ = Mid$(id$, InStr(id$, "(WID:") + 5)
ProgID$ = Text1(0).text
If ProgID$ = "" Then
  MsgBox (transe("Bitte wählen Sie zuerst ein Programm"))
  Exit Sub
End If
neuwerkid$ = ""
Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus
Call werkvz.callbackinit("prog")
While neuwerkid$ = "": DoEvents: Wend

If neuwerkid$ = "" Or neuwerkid$ = "_LOGOUT_" Then Exit Sub

ask% = MsgBox(transe("Ersetzen") & ":" & vbCrLf & form1.getwerknamebyid(id$) & vbCrLf & form1.getwerknamebyid(neuwerkid$) & vbCrLf & transe("ja: in allen Programmen ersetzen") & vbCrLf & transe("nein: nur hier ersetzen"), vbYesNo + vbCritical + vbDefaultButton2, transe("ersetzen") & "?")
If ask% = vbNo Then
  cmd$ = "update programmliste set werkid='" + neuwerkid$ + "' where programmid='" + ProgID + "' and werkid='" + id$ + "'"
Else
  cmd$ = "update programmliste set werkid='" + neuwerkid$ + "' where werkid='" + id$ + "'"
End If
Call form1.sqlqry(cmd$)
Call List1_Click

End Sub

Private Sub Command16_Click()
Dim rtmp As ADODB.Recordset, rrr
Dim neuid As String, i As Integer, nid$, c$, ProgID$

Call savecheck
ProgID$ = Text1(0).text
If ProgID$ = "" Then
  MsgBox ("Bitte wählen Sie zuerst ein Programm")
  Exit Sub
End If
neuid = InputBox(transe("Neues Programm"), "")
If trm(neuid) = "" Then Exit Sub
MousePointer = 11: DoEvents
Call form1.sqlqry("INSERT INTO programm (programmID) VALUES('" & neuid & "')")
Set rtmp = New ADODB.Recordset
rrr = form1.adoopen(rtmp, "SELECT * FROM programmliste where programmid='" + ProgID$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, "", "")
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  nid$ = form1.newid("programmliste", "id", 16)
  c$ = "insert into programmliste (id,ProgrammID,WerkID,Position,besetztid) values('" + nid$ + "','"
  c$ = c$ + neuid + "','"
  c$ = c$ + trm(rtmp!werkid) + "',"
  c$ = c$ + trm(rtmp!Position) + ",'"
  c$ = c$ + trm(rtmp!besetztid) + "')"
  Call form1.sqlqry(c$)
  rtmp.MoveNext
Wend
Call rlist1
For i = 0 To List1.ListCount - 1
  If List1.List(i) = neuid Then
    List1.ListIndex = i
    Exit For
  End If
Next i
MousePointer = 0
End Sub

Private Sub Command17_Click()
Dim id$, oid$, auftrid$

id$ = Text1(0).text
If id$ = "" Then Exit Sub
oid$ = oksel.text
auftrid$ = ""
If oid$ <> "" Then
  If InStr(oid$, "(ID:") > 0 Then
    auftrid$ = Mid$(oid$, InStr(oid$, "(ID:") + 4)
  End If
End If
prog.MousePointer = 11
Call form1.prgdruck(id$, form1.getusersetting("gemaliste", "prgdrucknoheadgema.rtf"), mbes.value, auftrid$)
prog.MousePointer = 0

End Sub

Private Sub Command18_Click()
Dim id$, oid$, auftrid$

id$ = Text1(0).text
If id$ = "" Then Exit Sub
oid$ = oksel.text
auftrid$ = ""
If oid$ <> "" Then
  If InStr(oid$, "(ID:") > 0 Then
    auftrid$ = Mid$(oid$, InStr(oid$, "(ID:") + 4)
  End If
End If
prog.MousePointer = 11
Call form1.prgdruck(id$, form1.getusersetting("publishersliste", "prgdrucknoheadpublishers.rtf"), mbes.value, auftrid$)
prog.MousePointer = 0


End Sub

Private Sub Command19_Click()
Call form1.handbuchcall("08-Programme.htm")

End Sub

Private Sub Command2_Click()
Dim rtmp As QueryDef
Dim neuid As String, i As Integer

Call savecheck
neuid = InputBox(transe("Neues Programm"), "")
If trm(neuid) = "" Then Exit Sub
Call form1.sqlqry("INSERT INTO programm (programmID) VALUES('" & neuid & "')")
Call rlist1
For i = 0 To List1.ListCount - 1
  If List1.List(i) = neuid Then
    List1.ListIndex = i
    Exit For
  End If
Next i
End Sub

Private Sub Command3_Click()
Dim cmd$, rtmp As QueryDef, up$, werkid$, ProgID$
Dim stmp As ADODB.Recordset

ProgID$ = Text1(0).text
If ProgID$ = "" Then
  MsgBox ("Bitte wählen Sie zuerst ein Programm")
  Exit Sub
End If
neuwerkid$ = ""
Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus
Call werkvz.callbackinit("prog")
While neuwerkid$ = "": DoEvents: Wend

If neuwerkid$ = "" Or neuwerkid$ = "_LOGOUT_" Then Exit Sub

cmd$ = "insert into programmliste (id,programmid,werkid,position) values('" + form1.newid("programmliste", "id", 20) & "','" + ProgID$ & "','" + neuwerkid$ & "'," & trm(10 * (List2.ListCount + 1)) & ")"
Call form1.sqlqry(cmd$)
Call List1_Click

End Sub
Public Sub callback(wrkid$)

neuwerkid$ = wrkid$

End Sub

Private Sub Command33_Click()
Dim tx$, i%

i% = List1.ListIndex
If i% < 0 Then Exit Sub
Clipboard.Clear
tx$ = List1.List(i%)
tx$ = form1.rdprog(tx$)
Clipboard.settext tx$

End Sub

Private Sub Command4_Click()
Dim i%, pid$, id$

i% = List2.ListIndex
If i% < 1 Then Exit Sub
pid$ = Text1(0).text
If pid$ = "" Then Exit Sub

id$ = List2.List(i%)
id$ = Mid$(id$, InStr(id$, "(MYID:") + 6)
id$ = trm(Left$(id$, InStr(id$, "(WID:") - 1))
form1.sqlqry ("update programmliste set position=" & i% & " where id='" & id$ & "' and programmid='" + pid$ & "'")
id$ = List2.List(i% - 1)
id$ = Mid$(id$, InStr(id$, "(MYID:") + 6)
id$ = trm(Left$(id$, InStr(id$, "(WID:") - 1))
form1.sqlqry ("update programmliste set position=" & i% + 1 & " where id='" & id$ & "' and programmid='" + pid$ & "'")
Call List1_Click
List2.ListIndex = i% - 1


End Sub

Private Sub Command5_Click()
Dim i%, pid$, id$

i% = List2.ListIndex
If i% >= List2.ListCount - 1 Then Exit Sub
pid$ = Text1(0).text
If pid$ = "" Then Exit Sub

id$ = List2.List(i%)
id$ = Mid$(id$, InStr(id$, "(MYID:") + 6)
id$ = trm(Left$(id$, InStr(id$, "(WID:") - 1))
form1.sqlqry ("update programmliste set position=" & i% + 2 & " where id='" & id$ & "' and programmid='" + pid$ & "'")
id$ = List2.List(i% + 1)
id$ = Mid$(id$, InStr(id$, "(MYID:") + 6)
id$ = trm(Left$(id$, InStr(id$, "(WID:") - 1))
form1.sqlqry ("update programmliste set position=" & i% + 1 & " where id='" & id$ & "' and programmid='" + pid$ & "'")
Call List1_Click
List2.ListIndex = i% + 1


End Sub

Private Sub Command6_Click()
Dim id$, oid$, auftrid$

id$ = Text1(0).text
If id$ = "" Then Exit Sub
oid$ = oksel.text
auftrid$ = ""
If oid$ <> "" Then
  If InStr(oid$, "(ID:") > 0 Then
    auftrid$ = Mid$(oid$, InStr(oid$, "(ID:") + 4)
  End If
End If
prog.MousePointer = 11
Call form1.prgdruck(id$, form1.getusersetting("programmvorlageohnekopf", "prgdrucknohead.rtf"), mbes.value, auftrid$)
prog.MousePointer = 0
End Sub

Private Sub Command7_Click()
Dim i%, pid$, id$

i% = List2.ListIndex
If i% < 0 Then Exit Sub
pid$ = Text1(0).text
If pid$ = "" Then Exit Sub

id$ = List2.List(i%)
id$ = Mid$(id$, InStr(id$, "(MYID:") + 6)
id$ = trm(Left$(id$, InStr(id$, "(WID:") - 1))
form1.sqlqry ("delete from programmliste where id='" & id$ & "' and programmid='" + pid$ & "'")
List2.RemoveItem i%

End Sub


Private Sub Command8_Click()
Dim i%

For i% = 0 To chgs.ListCount - 1
  form1.sqlqry (chgs.List(i%))
Next i%
chgs.Clear
BackColor = form1.cleancolor()
Command8.Enabled = False
End Sub

Private Sub Command9_Click()
Dim r As ADODB.Recordset, li%, id$, p%, rrr

Dim d2infile As String, d2insub As String
d2infile = "prog": d2insub = "Command9_Click"
li% = List1.ListIndex

id$ = List1.List(li%)
p% = InStr(id$, " in "): If p% > 0 Then id$ = Left$(id$, p% - 1)
p% = InStr(id$, " am "): If p% > 0 Then id$ = Left$(id$, p% - 1)
If id$ = "" Then Exit Sub

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM programmliste where programmid='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  MsgBox "Das Programm enthät Werke und kann deshalb nicht gelöscht werden."
  Exit Sub
End If
form1.sqlqry ("delete from programm where programmid='" + id$ + "'")
List1.RemoveItem li%

End Sub

Private Sub Form_Load()
Dim klrv%, s%, mew, meh

axsResizer1.SaveControlPositions
Width = 9350

okselmi% = -1
Command14.Enabled = False
nflds = 6
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
mew = form1.mylastwidth(Me.name, 0)
meh = form1.mylastheight(Me.name, 0)
If meh > 0 And mew > 0 Then
  Me.Width = mew
  Me.Height = meh
End If
Command10.Enabled = False
Call form1.formpos(Me)
s% = form1.myfontsize()
List1.Font.Size = s%
List2.Font.Size = s%
besetz.Font.Size = s%
maxbesetz.Font.Size = s%
usuchvz = 600#:
msec = 1# / (24# * 3600# * 1000#)

Randomize
Label5.Visible = False

timpaleft.Visible = False
Command13.Height = 495
Command13.Top = 4320
klrv% = Val(form1.mylastFormVar(Me.name, "timpaleft", "0"))
If klrv% <> 0 Then klrv% = 1
timpaleft.value = klrv%
klrv% = Val(form1.mylastFormVar(Me.name, "mbes", "0"))
If klrv% <> 0 Then klrv% = 1
mbes.value = klrv%
prog.Caption = transe("Programme - AgencyProf")
Command33.ToolTipText = transe("Programm in die Zwischenablage kopieren")
Command19.ToolTipText = transe("Hilfeseite öffnen")
Command13.Caption = transe("Bühnenplan")
Command12.ToolTipText = transe("Besetzung löschen")
Command11.Caption = transe("Name")
Command10.Caption = transe("gespielt?")
Command9.ToolTipText = transe("Löschen")
Command8.ToolTipText = transe("Speichern")
List3.ToolTipText = transe("In welchen Programmen wurde ein Werk gespielt?")
Command7.ToolTipText = transe("Markierten Programmpunkt löschen")
Command6.ToolTipText = transe("Programm ausdrucken")
Command5.Caption = transe("abw.")
Command15.Caption = transe("ersetzen")
Command5.ToolTipText = transe("Programmpunkt eine Position nach unten setzen")
Command4.Caption = transe("aufw.")
Command4.ToolTipText = transe("Programmpunkt eine Position nach oben setzen")
Command3.ToolTipText = transe("Programmpunkt hinzufügen")
List2.ToolTipText = transe("Programmablauf")
Command2.ToolTipText = transe("Neues Programm erstellen")
Command1.ToolTipText = transe("Formular schliessen")
Text1(0).ToolTipText = transe("Gewähltes Programm")
List1.ToolTipText = transe("Vorhandene Programme")
Label6.Caption = transe("Programm mit Doppelklick auswählen")
Label5.Caption = transe("Schlagw. li.")
maxbesetz.ToolTipText = transe("Aus der Maximalbesetzung können Sie einen Bühnenplan erstellen")
Label4.Caption = transe("Werk in Programmen")
Label3.Caption = transe("Programmpunkte")
Label2.Caption = transe("Markierter Part")

Show

Call rlist1
stimer.Enabled = True
End Sub
Public Sub rlist1()
Dim rtmp As ADODB.Recordset, cmd$, rrr, t$, swrd As String, s1$
Dim d2infile As String, d2insub As String, i%, stxt As String

d2infile = "prog": d2insub = "rlist1"
List1.Clear
Call nulldsp

If tponly$ = "" Then
  cmd$ = "SELECT * FROM programm"
  s1$ = trm(Text2.text): i% = 0: stxt = ""
  While Len(s1$) > 0 And i% < 5
    swrd = word1(s1$)
    If stxt <> "" Then stxt = stxt + "and "
    stxt = stxt + "instr(ProgrammID,'" + swrd + "')>0 "
    s1$ = word2bis(s1$)
    i% = i% + 1
  Wend
  If stxt <> "" Then cmd$ = cmd$ + " where " + stxt
Else
  cmd$ = "SELECT programm.* FROM programm INNER JOIN tpprogli ON programm.ProgrammID = tpprogli.prgid "
  cmd$ = cmd$ + "WHERE (((tpprogli.tpid)='" + tponly$ + "'))"
  tponly$ = ""
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!programmid) Then
    t$ = rtmp!programmid
    If Not IsNull(rtmp!Veranstaltungsort) And InStr(t$, trm(rtmp!Veranstaltungsort)) = 0 Then t$ = t$ & " in " & rtmp!Veranstaltungsort
    If Not IsNull(rtmp!anfangsdatum) Then t$ = t$ & " am " & rtmp!anfangsdatum
    List1.AddItem t$
  End If
  rtmp.MoveNext
Wend
rtmp.Close

End Sub


Sub nulldsp()
Dim i%, rrr

besetz.text = ""
maxbesetz.Caption = ""
oksel.Clear
Command14.Enabled = False
oksel.text = ""
okselmi% = -1
List2.Clear
Command10.Enabled = False
i% = 0


For i% = 0 To nflds
  On Error Resume Next
  Label1(i%).Caption = form1.sqla.TableDefs("programm").Fields(i%).name
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  Text1(i%).text = ""
Next i%

End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)
exuld:
On Error GoTo 0

End Sub

Private Sub List1_Click()
Dim rtmp As ADODB.Recordset, wid$, sid$, i%, li%, id$, p%, rrr, renum%, bz$
Dim d2infile As String, d2insub As String, c$

d2infile = "prog": d2insub = "List1_Click"
Call savecheck
i% = 0

li% = List1.ListIndex
Call nulldsp

id$ = List1.List(li%)
p% = InStr(id$, " in "): If p% > 0 Then id$ = Left$(id$, p% - 1)
p% = InStr(id$, " am "): If p% > 0 Then id$ = Left$(id$, p% - 1)
If id$ = "" Then Exit Sub

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM programm where programmid='" + id$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  For i% = 0 To nflds
    If Not IsNull(rtmp.Fields(i%)) Then
      Text1(i%).text = rtmp.Fields(i%)
    End If
  Next i%
End If
renum% = 0

List2.Clear
Command10.Enabled = False
On Error Resume Next
Set rtmp = New ADODB.Recordset
rrr = form1.adoopen(rtmp, "SELECT * FROM programmliste where programmid='" + id$ + "' order by position", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  wid$ = trm(rtmp!werkid)
  sid$ = ""
  If Left$(wid$, 4) = "SBZ:" Then
    sid$ = Mid$(wid$, 5)
    wid$ = form1.getsatzidbywerkid(sid$)
  End If
  bz$ = ""
  If besetz.Enabled = True Then
  On Error Resume Next
  If trm(rtmp!besetztid) <> "" Then
    bz$ = form1.bestzstr(rtmp!besetztid)
  Else
    bz$ = form1.defaultbesetzt(wid$)
  End If
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    besetz.Enabled = False
    MsgBox "Die Datenbankstruktur muss erst geändert werden. (besetztid fehlt)." & vbCrLf & "Bitte lontaktieren Sie Ihren Support"
  End If
  End If
  If sid$ = "" Then
    List2.AddItem form1.getdauerbywerkid(wid$) & " Min., " & form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) & " (" & bz$ & ")" & Space$(120) & "(MYID:" & rtmp!id & " (WID:" & wid$
  Else
    List2.AddItem form1.getsatznamebyid(sid$) & " " + transe("aus") + " " & form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) & " (" & bz$ & ")" & Space$(120) & "(MYID:" & rtmp!id & " (WID:" & "SBZ:" + sid$
  End If
  DoEvents
  If rtmp!Position <> List2.ListCount Then renum% = 1
  rtmp.MoveNext
Wend
Call rdmaxbes

If renum% = 1 Then
  For i% = 0 To List2.ListCount - 1
    wid$ = List2.List(i%)
    wid$ = Mid$(wid$, InStr(wid$, "(MYID:") + 6)
    wid$ = trm(Left$(wid$, InStr(wid$, "(WID:") - 1))
    form1.sqlqry ("update programmliste set position=" & i% + 1 & " where id='" & wid$ & "' and programmid='" & id$ & "'")
  Next i%
End If
Label4.Caption = "Programm in Terminen"
List3.Clear
c$ = "SELECT auftrittsid,bezeichnung,datum,ort FROM auftritthigru inner join auftritt on auftritt.id=auftritthigru.auftrittsid where felddaten='" + id$ + "' and feldname='Programm' order by auftritt.datum"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  List3.AddItem datum2sql(trm(rtmp!datum)) + " " + trm(rtmp!ort) + " " + trm(rtmp!bezeichnung) + " " + Space$(160) + "(AID:" + rtmp!auftrittsid
  rtmp.MoveNext
Wend

End Sub
Private Sub rdmaxbes()
Dim mb$, i%

On Error GoTo exme
mb$ = ""
If mbes.value > 0 Then
  For i% = 0 To List2.ListCount - 1
   List2.ListIndex = i%
   DoEvents
   mb$ = form1.neumaxbesetz(mb$, besetz.text)
   maxbesetz.Caption = mb$
   DoEvents
  Next i%
  maxbesetz.Caption = mb$
End If
Exit Sub

exme:
On Error GoTo 0

End Sub
Private Sub List1_DblClick()
Dim pid$

pid$ = Text1(0).text
If pid$ = "" Then Exit Sub
If callbck$ <> "" Then
  Select Case callbck$
    Case "tplan": Call tplan.callback(pid$)
                  Call tplan.SetFocus
    Case "auftritt": Call auftritt.callback(pid$)
                     Call auftritt.SetFocus
    Case Else
  End Select
  callbck$ = ""
  'Command1.Caption = "&Schliessen"
End If
callbck$ = ""
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then Call Command9_Click

End Sub



Private Sub List2_Click()
Dim r As ADODB.Recordset, kid$, wid$, rrr, ad$
Dim d2infile As String, d2insub As String

d2infile = "prog": d2insub = "List2_Click"
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
Command10.Enabled = True
If besetz.Enabled = False Then Exit Sub

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM b_loc where wid='" & wid$ & "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
besetz.Clear
While Not r.EOF
  ad$ = form1.bestzstr(r!id)
  besetz.AddItem ad$ & Space$(80) & "ID:" & r!id
  r.MoveNext
Wend

If InStr(kid$, "(MYID:") = 0 Then Exit Sub
wid$ = Mid$(kid$, InStr(kid$, "(MYID:") + 6)
wid$ = trm(Left$(wid$, InStr(wid$, " ") - 1))
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT besetztid ,werkid FROM programmliste where id='" + wid$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
If trm(r!besetztid) = "" Then
  besetz.text = form1.defaultbesetzt(r!werkid)
Else
  besetz.text = form1.bestzstr(r!besetztid)
End If
End If

End Sub

Private Sub List2_DblClick()
Dim rtmp As ADODB.Recordset, kid$, wid$, sid$, i%
Dim d2infile As String, d2insub As String

d2infile = "prog": d2insub = "List2_DblClick"
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
If Left$(wid$, 4) = "SBZ:" Then
    sid$ = Mid$(wid$, 5)
    wid$ = form1.getsatzidbywerkid(sid$)
End If
Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus
kid$ = form1.getkompnamebywerkid(wid$)
Call werkvz.showkompdetailbyname(kid$)
Call werkvz.Timer2_Timer
'Call werkvz.showwerkdetail(wid$) war hier nich sooo gut. so bessa
For i% = 0 To werkvz.List2.ListCount - 1
  If InStr(werkvz.List2.List(i%), form1.getwerknamebyid(wid$)) = 1 Then
    werkvz.List2.ListIndex = i%
    Exit Sub
  End If
Next i%

End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then Call Command7_Click

End Sub

Private Sub List3_Click()
Dim idx%, i%, s$

idx% = List3.ListIndex
s$ = List3.List(idx%)
For i% = 0 To List1.ListCount - 1
  If Left(List1.List(i%), Len(s$)) = s$ Then
    List1.ListIndex = i%
    Exit Sub
  End If
Next i%

End Sub

Private Sub List3_DblClick()
Dim idx%, i%, s$, c$, r As ADODB.Recordset, rrr, s0$, p%
Dim d2infile As String, d2insub As String

d2infile = "prog": d2insub = "List3_DblClick"
idx% = List3.ListIndex
s0$ = List3.List(idx%)
p% = InStr(s0$, "(AID:")
If p% > 0 Then
  s$ = Mid(s0$, p% + 5)
  c$ = "select * from auftritt where id='" & s$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    Load auftritt
    On Error Resume Next
    Call auftritt.SetFocus
    On Error GoTo 0
    Call auftritt.showrec(s$, 0)
  End If
End If
End Sub

Private Sub maxbesetz_DblClick()
Dim bboxy%, n
Dim blaeser_rest_calc, p%, bcnt%, blaesercount, instrumente, bbox%, ob_pro_blaeserzeile
Dim kb, wert$, brt, blaeser_pro_zeile, blaeser_rest, blaeserzeile, need, blaeserzeilen
Dim timx As Single, timdx As Single, b0off As Single, s$, l$, fl, ob, kl, fa
Dim hb, i%, ho, tr, po, tu, bb, pa, tri, be, sc, v1, v2, vi, ce
Dim d2infile As String, d2insub As String

d2infile = "prog": d2insub = "maxbesetz_DblClick"
s$ = maxbesetz.Caption
Load bplan
Call bplan.delme_Click
If s$ = "" Then Exit Sub
l$ = s$
bplan.pgid = Text1(0).text
  fl = Val(s$)
  l$ = Mid(s$, InStr(s$, "/") + 1)
  ob = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  kl = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  fa = Val(l$): If fa < 0 Then fa = 0
  l$ = Mid(l$, InStr(l$, "/") + 1)
  hb = fl + ob + kl + fa
  i% = 1: While Mid(l$, i%, 1) <> "/": i% = i% + 1: Wend
  While Mid(l$, i%, 1) <> " " And i% > 1: i% = i% - 1: Wend
  l$ = Mid$(l$, i%)

  ho = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  tr = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  po = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  tu = Val(l$): If tu < 0 Then tu = 0
  l$ = Mid(l$, InStr(l$, "/") + 1)
  bb = ho + tr + po + tu
  i% = 1: While Mid(l$, i%, 1) <> "/": i% = i% + 1: Wend
  While Mid(l$, i%, 1) <> " " And i% > 1: i% = i% - 1: Wend
  l$ = trm(Mid$(l$, i%))

  pa = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  tri = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  be = Val(l$): If be < 0 Then be = 0
  l$ = Mid(l$, InStr(l$, "/") + 1)
  sc = pa + tri + be
  i% = 1: While Mid(l$, i%, 1) <> "/": i% = i% + 1: Wend
  While Mid(l$, i%, 1) <> " " And i% > 1: i% = i% - 1: Wend
  l$ = trm(Mid$(l$, i%))

  v1 = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  v2 = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  vi = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  ce = Val(l$)
  l$ = Mid(l$, InStr(l$, "/") + 1)
  kb = Val(l$)
  s = v1 + v2 + vi + ce + kb

  wert$ = InputBox(transe("Geben Sie die Bühnenbreite (in cm) vor:"), "", trm((bb + hb) * 100))
nxtry:
  brt = Val(wert$)
  If brt <= 0 Then Exit Sub
  bplan.Text2.text = wert$
  blaeser_pro_zeile = Int(brt / 100)
  If (blaeser_pro_zeile Mod 2) = 0 Then blaeser_pro_zeile = blaeser_pro_zeile - 1
  b0off = brt / 2 - (blaeser_pro_zeile - 1) * 100 / 2
  blaeser_rest = hb + bb
  blaeserzeile = 0
  bplan.bbreit(0) = "80"
  bplan.bbreit(1) = "80"
  bplan.Text3.text = blaeser_pro_zeile
  While blaeser_rest > blaeser_pro_zeile
    blaeser_rest = blaeser_rest - blaeser_pro_zeile
    bplan.Check1(3).value = 1
    Call bplan.Check1_Click(3)
    DoEvents
    Call bplan.p1_MouseDown(1, 0, b0off, 100 + blaeserzeile * 100)
    DoEvents
    blaeserzeile = blaeserzeile + 1
  Wend
  If blaeser_rest < 4 Then
    need = 5 - blaeser_rest
    wert$ = InputBox(transe("Ungünstige Bühnenbreite ich empfehle (in cm):"), "", brt + need * 100)
    If Val(wert$) <> brt Then
      Call bplan.delme_Click
      GoTo nxtry
    End If
  End If
  blaeserzeilen = blaeserzeile + 1
  bplan.Text3.text = blaeser_rest
  blaeser_rest_calc = blaeser_rest
  If (blaeser_rest Mod 2) = 0 Then blaeser_rest_calc = blaeser_rest + 1
  DoEvents
  bplan.Check1(3).value = 1
  Call bplan.Check1_Click(3)
  DoEvents
  Call bplan.p1_MouseDown(1, 0, brt / 2 - (blaeser_rest_calc - 1) * 100 / 2, 100 + blaeserzeile * 100)
  DoEvents

  p% = 0

  bcnt% = 0
  blaesercount = hb + bb

  'oboen beginnen vor dem dirigenten in basic:
  instrumente = ob
  bbox% = Int(blaeserzeile * blaeser_pro_zeile + blaeser_rest / 2)
  ob_pro_blaeserzeile = Int(instrumente / (blaeserzeile + 1))
  If ob_pro_blaeserzeile < 2 Then ob_pro_blaeserzeile = 2
  While bcnt% < instrumente
    Call bplan.List1DClick(bbox%, "Ob" & Val(bcnt% + 1))
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
    If (bcnt% Mod ob_pro_blaeserzeile) = 0 Or bbox% > blaesercount - 1 Then
      bbox% = bbox% - blaeser_pro_zeile - ob_pro_blaeserzeile
      bboxy% = bbox% / blaeser_pro_zeile
      bbox% = Int(bboxy% * blaeser_pro_zeile + blaeser_pro_zeile / 2)
      'If bbox% Mod 2 = 0 Then bbox% = bbox% - 1
    End If
  Wend

  'trompeten rechts der oboen:
  instrumente = tr
  bcnt% = 0
  bbox% = Int(blaeserzeile * blaeser_pro_zeile + blaeser_rest / 2)
  ob_pro_blaeserzeile = Int(instrumente / (blaeserzeile + 1))
  If ob_pro_blaeserzeile < 2 Then ob_pro_blaeserzeile = 2
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    If bbox% > blaesercount - 1 Then bbox% = 0
    Call bplan.List1DClick(bbox%, "Trp" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
    'genug instrumente in dieser Zeile oder Zeile voll
    If (bcnt% Mod ob_pro_blaeserzeile) = 0 Or bbox% > blaesercount - 1 Then
      bbox% = bbox% - blaeser_pro_zeile - ob_pro_blaeserzeile
      bboxy% = bbox% / blaeser_pro_zeile
      bbox% = bboxy% * blaeser_pro_zeile + (blaeser_pro_zeile / 2)
    End If
  Wend

  'flöten links der oboen:
  instrumente = fl
  bcnt% = 0
  ob_pro_blaeserzeile = Int(instrumente / (blaeserzeile + 1))
  If ob_pro_blaeserzeile < 2 Then ob_pro_blaeserzeile = 2
  bbox% = Int(blaeserzeile * blaeser_pro_zeile + blaeser_rest / 2) - ob_pro_blaeserzeile
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    If bbox% > blaesercount - 1 Then bbox% = 0
    Call bplan.List1DClick(bbox%, "Fl" & bcnt% + 1)
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
    'genug instrumente in dieser Zeile oder Zeile voll
    If (bcnt% Mod ob_pro_blaeserzeile) = 0 Or bbox% > blaesercount - 1 Then
      bbox% = bbox% - blaeser_pro_zeile - ob_pro_blaeserzeile
      bboxy% = bbox% / blaeser_pro_zeile
      bbox% = bboxy% * blaeser_pro_zeile + (blaeser_pro_zeile / 2) - ob_pro_blaeserzeile
    End If
  Wend

  'das war easy, Hörner links
  instrumente = ho
  bcnt% = 0
  bbox% = 0
  ob_pro_blaeserzeile = Int(instrumente / (blaeserzeile + 1))
  If ob_pro_blaeserzeile < 2 Then ob_pro_blaeserzeile = 2
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    If bbox% > blaesercount - 1 Then bbox% = 0
    Call bplan.List1DClick(bbox%, "Hr" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
    'genug instrumente in dieser Zeile oder Zeile voll
    If (bcnt% Mod ob_pro_blaeserzeile) = 0 Or bbox% > blaesercount - 1 Then
      bbox% = bbox% + blaeser_pro_zeile
      bboxy% = bbox% / blaeser_pro_zeile
      bbox% = bboxy% * blaeser_pro_zeile
    End If
  Wend

  'Tuben rechts
  instrumente = tu
  bcnt% = 0
  ob_pro_blaeserzeile = Int(instrumente / (blaeserzeile + 1))
  If ob_pro_blaeserzeile < 2 Then ob_pro_blaeserzeile = 2
  bbox% = blaeser_pro_zeile - ob_pro_blaeserzeile
  While bcnt% < instrumente
rtrtu:
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    If bbox% > blaesercount - 1 Then
      bbox% = 0
      GoTo rtrtu
    End If
    Call bplan.List1DClick(bbox%, "Tu" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
    'genug instrumente in dieser Zeile oder Zeile voll
    If (bcnt% Mod ob_pro_blaeserzeile) = 0 Or bbox% > blaesercount - 1 Then
      bbox% = bbox% + blaeser_pro_zeile - ob_pro_blaeserzeile
      bboxy% = bbox% / blaeser_pro_zeile
      bbox% = (bboxy% * blaeser_pro_zeile - 1) - ob_pro_blaeserzeile
    End If
  Wend

  'ab jetzt: auffüllen was frei ist
  instrumente = kl
  bcnt% = 0
  bbox% = 0
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    Call bplan.List1DClick(bbox%, "Kl" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
  Wend

  instrumente = fa
  bcnt% = 0
  bbox% = 0
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    Call bplan.List1DClick(bbox%, "Fag" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
  Wend

  instrumente = po
  bcnt% = 0
  bbox% = 0
  While bcnt% < instrumente
    While bplan.getobd(bbox%) <> ""
      bbox% = bbox% + 1
    Wend
    Call bplan.List1DClick(bbox%, "Pos" & Val(bcnt% + 1))
    DoEvents
    bcnt% = bcnt% + 1
    bbox% = bbox% + 1
  Wend




'  For i% = 1 To fa
'    Call bplan.List1DClick(p%, "Fa" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To kl
'    Call bplan.List1DClick(p%, "Kl" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To ob
'    Call bplan.List1DClick(p%, "Ob" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To fl
'    Call bplan.List1DClick(p%, "Fl" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%

'  For i% = 1 To tu
'    Call bplan.List1DClick(p%, "Tu" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To po
'    Call bplan.List1DClick(p%, "Po" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To tr
'    Call bplan.List1DClick(p%, "Tr" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%
'  For i% = 1 To ho
'    Call bplan.List1DClick(p%, "Ho" & Val(i%))
'    p% = p% + 1
'    DoEvents
'  Next i%

bplan.Text4.text = "4"
bplan.Check1(4).value = 1
Call bplan.Check1_Click(4)
DoEvents
Call bplan.p1_MouseDown(1, 0, brt / 2, 100 + blaeserzeile * 100 + 400)
DoEvents
n = 0
For i% = 0 To bplan.List1.ListCount - 1
  If InStr(bplan.List1.List(i%), "Streicher") > 0 Then
    Select Case n:
      Case 3: Call bplan.List1DClick(i%, trm(v1) & " Vl1")
      Case 2: Call bplan.List1DClick(i%, trm(v2) & "Vl2")
      Case 1: Call bplan.List1DClick(i%, trm(vi) & " Va")
      Case 0: Call bplan.List1DClick(i%, trm(ce) & " VC")
      Case Else:
    End Select
    n = n + 1
  End If
Next i%
If kb > 0 Then
  If timpaleft.value = 0 Then
    timx = 100
  Else
    timx = bplan.p1.ScaleWidth - 100
  End If
  bplan.bbreit(0) = "80"
  bplan.bbreit(1) = "80"
  bplan.Text3.text = "1"
  bplan.Check1(3).value = 1
  Call bplan.Check1_Click(3)
  DoEvents
  Call bplan.p1_MouseDown(1, 0, timx, 100 + blaeserzeile * 100 + 200)
  Call bplan.List1DClick(bplan.List1.ListCount - 1, trm(kb) & " KB")
End If


If timpaleft.value = 0 Then
  timx = bplan.p1.ScaleWidth - 130
  timdx = 20
Else
  timx = 130
  timdx = -20
End If
bplan.sbreit(0) = "70"
bplan.sbreit(1) = "70"
bplan.Check1(0).value = 1
Call bplan.Check1_Click(0)
DoEvents
Call bplan.p1_MouseDown(1, 0, timx, 100 + blaeserzeile * 100 + 200)
DoEvents
Call bplan.List1DClick(bplan.List1.ListCount - 1, "Tim")
DoEvents
bplan.Check1(0).value = 1
Call bplan.Check1_Click(0)
DoEvents
Call bplan.p1_MouseDown(1, 0, timx + timdx, 100 + blaeserzeile * 100 + 240)
DoEvents
Call bplan.List1DClick(bplan.List1.ListCount - 1, "pa")
DoEvents
bplan.Check1(0).value = 1
Call bplan.Check1_Click(0)
DoEvents
Call bplan.p1_MouseDown(1, 0, timx + timdx, 100 + blaeserzeile * 100 + 280)
DoEvents
Call bplan.List1DClick(bplan.List1.ListCount - 1, "ni")
DoEvents


bplan.bbreit(0) = "100"
bplan.bbreit(1) = "100"
bplan.Text3.text = "1"
bplan.Check1(3).value = 1
Call bplan.Check1_Click(3)
DoEvents
Call bplan.p1_MouseDown(1, 0, brt / 2, 100 + blaeserzeile * 100 + 400)
Call bplan.List1DClick(bplan.List1.ListCount - 1, "Dirigent")
DoEvents

bplan.Text1.text = 100 + blaeserzeile * 100 + 550
DoEvents
Call bplan.Command2_Click
End Sub

Private Sub mbes_Click()

If mbes.value > 0 Then
  Label5.Visible = True
  timpaleft.Visible = True
  Command13.Height = 255
  Command13.Top = 4560
  Call rdmaxbes
Else
  Label5.Visible = False
  timpaleft.Visible = False
  Command13.Height = 495
  Command13.Top = 4320
End If
Call form1.setmylastFormVar(Me.name, "mbes", trm(mbes.value))

End Sub


Private Sub oksel_Click()

okselmi% = List1.ListIndex
Command14.Enabled = True

End Sub

Private Sub oksel_DropDown()
Dim r As ADODB.Recordset, s As ADODB.Recordset, rrr
Dim id$, li%, p%, cmd$
Dim d2infile As String, d2insub As String

d2infile = "prog": d2insub = "oksel_DropDown"
li% = List1.ListIndex
If li% < 0 Then Exit Sub
oksel.Clear
Command14.Enabled = False
id$ = List1.List(li%)
p% = InStr(id$, " in "): If p% > 0 Then id$ = Left$(id$, p% - 1)
p% = InStr(id$, " am "): If p% > 0 Then id$ = Left$(id$, p% - 1)
If id$ = "" Then Exit Sub
cmd$ = "select auftrittsid from auftritthigru where felddaten='" + id$ + "' and feldname='Programm'"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If r.EOF Then Exit Sub

r.MoveFirst
While Not r.EOF
  cmd$ = "select ort,datum from auftritt where id='" & r!auftrittsid & "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not s.EOF Then oksel.AddItem s!ort & ", den  " & datfromsql(s!datum) & Space$(80) & "(ID:" & r!auftrittsid
  r.MoveNext
Wend

End Sub

Private Sub stimer_Timer()
Dim s$

'd2infile = "Form1": d2insub = "Timer1_Timer"
If now() < snotb41 Then Exit Sub
stimer.Enabled = False
s$ = Text2.text
Call rlist1

End Sub

Private Sub Text1_GotFocus(Index As Integer)

prv$ = Text1(Index).text

End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim s$, id$, nwert$, fld$, ask%, cmd$

If Index > 0 Then

id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  Exit Sub
End If
nwert$ = trm(Text1(Index).text)
If nwert$ <> prv$ Then
  fld$ = Label1(Index).Caption
  If Index = 2 Or Index = 3 Then nwert$ = datum2sql(nwert$)
  If nwert$ = "" Then
    nwert$ = "NULL"
  Else
    nwert$ = "'" + nwert$ + "'"
  End If
  chgs.AddItem "update programm set " + fld$ + "=" + nwert$ + " where programmid='" + id$ + "'"
  prog.BackColor = form1.dirtycolor()
  Command8.Enabled = True
End If

End If 'index=0
If Index = 0 Then
  If prgprvn$ <> "" And prgprvn$ <> Text1(0).text Then
    ask% = MsgBox("Soll das Programm " & prgprvn$ & " umbenannt werden in " & Text1(0).text, vbYesNo + vbCritical + vbDefaultButton2, "Programm umbenennen?")
    If ask% = vbYes Then
      MousePointer = 11
      cmd$ = "update programm set programmid='" & Text1(0).text & "' where programmid='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update programmliste set programmid='" & Text1(0).text & "' where programmid='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update tpprogli set prgid='" & Text1(0).text & "' where prgid='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update usr_" + utabn("künstlerauftritt") + " set programm='" & Text1(0).text & "' where programm='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update usr_orchesterauftritt set programm='" & Text1(0).text & "' where programm='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update usr_promo set programm='" & Text1(0).text & "' where programm='" & prgprvn$ & "';": Call form1.sqlqry(cmd$)
      cmd$ = "update auftritthigru set felddaten='" & Text1(0).text & "' where ((felddaten='" & prgprvn$ & "') and (feldname='Programm'));": Call form1.sqlqry(cmd$)
      MousePointer = 0
    End If
  End If
  prgprvn$ = ""
  Command8.Enabled = False
  Call rlist1
End If
End Sub
Public Sub callbackinit(frm$, onlytp$)
callbck$ = frm$
tponly$ = onlytp$
Call rlist1
Command1.Visible = False
'Command1.Caption = "Doppelklick aufs Programm"
'Command1.Enabled = False
End Sub
Public Sub selectone(id$)
Dim i%

For i% = 0 To List1.ListCount - 1
  If id$ = Left(List1.List(i%), Len(id$)) Then
    List1.ListIndex = i%
    i% = List1.ListCount
  End If
Next i%

End Sub

Sub savecheck()
Dim antw As Integer

If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  If antw = vbYes Then Call Command8_Click
End If
BackColor = form1.cleancolor()
End Sub

Private Sub Text2_Change()
Dim sText As String

sText = trm(Text2.text)
stimer.Enabled = False
stimer.Interval = 200
stimer.Enabled = True
snotb41 = now() + usuchvz * msec

End Sub

Private Sub Timer1_Timer()
Call List2_Click
prog.BackColor = form1.cleancolor()
DoEvents
Timer1.Enabled = False
Timer1.Interval = 0
End Sub

Private Sub timpaleft_Click()

Call form1.setmylastFormVar(Me.name, "timpaleft", trm(timpaleft.value))

End Sub
