VERSION 5.00
Begin VB.Form auftrittshintergrund 
   Caption         =   "Hintergrunddaten für Auftritte und Adressen"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check2 
      Caption         =   "&Share with Horde"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   41
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton chklst 
      Caption         =   "Checklist"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   40
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Textmarken"
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Folge"
      Height          =   255
      Left            =   2520
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "kopieren"
      Height          =   255
      Left            =   720
      TabIndex        =   37
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7920
      TabIndex        =   36
      Text            =   "DDDDDD"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   120
      Picture         =   "auftrittshintergrund.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Formular schiessen"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Left            =   6600
      Top             =   3000
   End
   Begin VB.CommandButton Command16 
      Caption         =   "löschen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "upd"
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "del"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   3000
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5760
      Top             =   2280
   End
   Begin VB.CommandButton Command13 
      Caption         =   "diese erstellen"
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Abgleich"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "alle erstellen"
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Farbe setzen"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox kfrb 
      Height          =   285
      Index           =   2
      Left            =   8880
      TabIndex        =   24
      Text            =   "B"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox kfrb 
      Height          =   285
      Index           =   1
      Left            =   8400
      TabIndex        =   23
      Text            =   "G"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox kfrb 
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   22
      Text            =   "R"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ggf. Daten löschen"
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Top             =   960
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Namen ändern"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   0
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7560
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8520
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&OK, speichern"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   2640
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "auftrittshintergrund.frx":0250
      Left            =   7560
      List            =   "auftrittshintergrund.frx":0252
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   585
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ab"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "auf"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ab"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "auf"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "neu"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ne&u"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "RGB in HEX"
      Height          =   255
      Left            =   6840
      TabIndex        =   35
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Abkz."
      Height          =   255
      Left            =   480
      TabIndex        =   33
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Kalenderfarbe (Rot-Grün-Blau)"
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "(0=löschen)"
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Wert"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   9240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "aus Tabelle"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Zeilen"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Feldname"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "auftrittshintergrund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cid$, auftstart As Integer
Dim wrkJet As Workspace
Dim sqla As Database, prvt$, nosv4 As Integer
Dim kfarbe%(2), notb4%

Sub rlist2()
Dim rtmp As ADODB.Recordset, at As ADODB.Recordset, i%, rrr
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "rlist2"
Set at = New ADODB.Recordset
at.CursorLocation = adUseServer
rrr = form1.adoopen(at, "SELECT * FROM auftrittstypen where id='" + cid$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If Not at.EOF Then
  For i% = 0 To 2
    kfrb(i%).text = "" & at.Fields(2 + i%).value
    If kfrb(i%).text = "" Then kfrb(i%).text = "192"
    kfrb(i%).Enabled = True
  Next i%
  Text4.text = "": If Not IsNull(at!abkz) Then Text4.text = at!abkz
Else
  For i% = 0 To 2
    kfrb(i%).Enabled = False
  Next i%
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftrittsfelder where typ='" + cid$ + "' order by position", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
List2.Clear
i% = -1
While Not rtmp.EOF
  List2.AddItem rtmp!feldname & " Zeilen: " & rtmp!zeilen & Space$(40) & "(ID:" & rtmp!id
  rtmp.MoveNext
Wend
For i% = 0 To List2.ListCount - 1
  If Left$(List2.List(i%), 10) = "NeuesFeld " Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%

End Sub
Sub rlist1()
Dim rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "rlist1"
List1.Clear
List2.Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id,sortierung FROM adresstypen order by sortierung", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

List1.AddItem "--Adressen"
While Not rtmp.EOF
  List1.AddItem rtmp!id
  If IsNull(rtmp!sortierung) Or rtmp!sortierung <> List1.ListCount * 10 Then
    List2.AddItem "update adresstypen set sortierung=" + trm(str$(List1.ListCount * 10)) + " where id='" + trm(rtmp!id) + "'"
  End If
  rtmp.MoveNext
Wend
auftstart = List1.ListCount
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftrittstypen order by sortierung", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

List1.AddItem "--Auftritt"
notb4% = List1.ListCount
While Not rtmp.EOF
  List1.AddItem rtmp!id
  If rtmp!sortierung <> List1.ListCount * 10 Then
    List2.AddItem "update auftrittstypen set sortierung=" + trm(str$(List1.ListCount * 10)) + " where id='" + trm(rtmp!id) + "'"
  End If
  rtmp.MoveNext
Wend
While List2.ListCount > 0
  form1.sqlqry (List2.List(0))
  List2.RemoveItem 0
Wend
End Sub

Private Sub Check2_Click()
Command8.Enabled = True
End Sub

Private Sub Combo1_Click()
Dim rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Combo1_Click"
Combo2.Clear
Select Case Combo1.text
  Case "programm": Combo2.AddItem "ID"
  Case "adrselect":
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresstypen", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

            While Not rtmp.EOF
                Combo2.AddItem rtmp!id
                rtmp.MoveNext
            Wend
  Case Else:
End Select

End Sub

Private Sub Command1_Click()
Hide
Unload auftrittshintergrund
End Sub

Private Sub Command10_Click()
Dim cmd$

cmd$ = "update auftrittstypen set kalenderfarbe_r='" + trm(kfrb(0).text) + "' where id='" & cid$ + "'": form1.sqlqry (cmd$)
cmd$ = "update auftrittstypen set kalenderfarbe_g='" + trm(kfrb(1).text) + "' where id='" & cid$ + "'": form1.sqlqry (cmd$)
cmd$ = "update auftrittstypen set kalenderfarbe_b='" + trm(kfrb(2).text) + "' where id='" & cid$ + "'": form1.sqlqry (cmd$)
Call form1.upd_colorcache(cid$, Val(kfrb(0).text), Val(kfrb(1).text), Val(kfrb(2).text))

End Sub

Private Sub Command11_Click()
Dim o%, u%, d%, gfi%, t%, i%, tn$, cmd$, d0, d1, flst$, fld$, X

auftrittshintergrund.MousePointer = 11

On Error Resume Next
Kill "sqlupd.txt"
Kill "sqldel.txt"
On Error GoTo 0

o% = FreeFile

u% = FreeFile
Open "sqlupd.txt" For Output As #u%
d% = FreeFile
Open "sqldel.txt" For Output As #d%

gfi% = 0
For t% = 0 To List1.ListCount - 1
  List1.ListIndex = t%
  DoEvents
  If gfi% = 1 Then
    i% = List2.ListCount
    flst$ = ""
    tn$ = "usr_" & utabn(List1.List(List1.ListIndex))
    Print #d%, "drop TABLE " + tn$ + ";"
    Print #u%, "CREATE TABLE " + tn$ + " (id varchar (120) default '0' not null, "
    If i% > 0 Then
      For i% = 0 To List2.ListCount - 1
        List2.ListIndex = i%
        DoEvents
        fld$ = trm(Text1.text)
        If fld$ <> "" Then
          If InStr(flst$, "-" + LCase(fld$) + "-") = 0 Then
            cmd$ = fld$ + " longtext,"
            flst$ = flst$ + "-" + LCase(fld$) + "-"
            Print #u%, cmd$
          End If
        End If
      Next i%
    End If
    Print #u%, "primary key(id));"
  Else
    If InStr(List1.List(t%), "-Auftritt") > 0 Then
      gfi% = 1
    End If
  End If
Next t%

Close #d%
Close #u%
Call form1.chgcreate("FLUSH TABLES;")
Call form1.chgappend("FLUSH STATUS;")
Call form1.crShell("sqlchg", False)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 4

Call form1.crShell("sqldel", False)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 4

Call form1.crShell("sqlupd", True)
X = Shell("notepad.exe sqlupd.txt", 1)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 5

auftrittshintergrund.MousePointer = 0


End Sub

Private Sub Command12_Click()
Dim r As ADODB.Recordset, i%, cmd$, rrr, d0, d1
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Command12_Click"
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  If Text1.text <> sqla.TableDefs("usr_" & utabn(cid)).Fields(i% + 1).name Then
    i% = List2.ListCount + 100
  End If
Next i%
If i% < List2.ListCount + 10 Then
  MsgBox ("tables in sync, no action required")
  Exit Sub
End If
Call form1.chgcreate("FLUSH TABLES;")
cmd$ = "select * from tmp_" + LCase(cid$)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If rrr = 0 Then
  r.Close
  Call form1.chgappend("drop table tmp_" + LCase(cid$) + ";")
End If
Call form1.chgappend("create table tmp_" + LCase(cid$) + " (id varchar (40) default '0' not null , primary key(id)); ")
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  Call form1.chgappend("alter table tmp_" + LCase(cid$) + " add column " + Text1.text + " longtext;")
Next i%
Call form1.chgappend("insert into tmp_" + LCase(cid$) + " select * from usr_" & utabn(cid$) + ";")

Call form1.chgappend("drop table usr_" & utabn(cid$) + ";")
Call form1.chgappend("create table usr_" & utabn(cid$) + " (id varchar (40) default '0' not null , primary key(id)); ")
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  Call form1.chgappend("alter table usr_" & utabn(cid$) + " add column " + Text1.text + " longtext;")
Next i%
Call form1.chgappend("insert into usr_" & utabn(cid$) + " select * from tmp_" + LCase(cid$) + ";")
Call form1.chgappend("FLUSH TABLES;")
Call form1.crShell("sqlchg", True)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 5
sqla.Close
Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())

End Sub

Private Sub Command13_Click()
Dim o%, u%, d%, tn$, i%, cmd$, d0, d1, X

Call Command18_Click
DoEvents
auftrittshintergrund.MousePointer = 11: DoEvents
On Error Resume Next
Kill "sqlupd.txt"
Kill "sqldel.txt"
On Error GoTo 0

o% = FreeFile

u% = FreeFile
Open "sqlupd.txt" For Output As #u%
d% = FreeFile
Open "sqldel.txt" For Output As #d%


tn$ = "usr_" & utabn(List1.List(List1.ListIndex))
Print #d%, "drop TABLE " + tn$ + ";"
Print #u%, "CREATE TABLE " + tn$ + " (id varchar (120) default '0' not null, "
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  If trm(Text1.text) <> "" Then
    cmd$ = Text1.text + " longtext,"
    Print #u%, cmd$
  End If
Next i%
Print #u%, "primary key(id));"
Print #u%, "FLUSH TABLES;"

Close #d%
Close #u%
X = Shell("notepad.exe sqlupd.txt", vbNormalFocus)
Call form1.chgcreate("FLUSH TABLES;")
Call form1.chgappend("FLUSH STATUS;")
Call form1.crShell("sqlchg", False)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 2

Call form1.crShell("sqldel", False)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 2

Call form1.crShell("sqlupd", False)
d0 = Time
Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 3
sqla.Close
Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
auftrittshintergrund.MousePointer = 0
'kein test, treiber merkt es nicht!
'Call form1.fieldcheck(tn$, "id")
'If form1.isfieldmissing(tn$, "id") Then
'  MsgBox ("Tabelle " + tn$ + " nicht anlegbar")
'End If
End Sub

Private Sub Command14_Click()
Dim d0, d1

    Call form1.crShell("sqldel", False)
    d0 = Time
    Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 5
    On Error Resume Next
    Kill "sqldel.txt"
    On Error GoTo 0
    Call Timer1_Timer
End Sub

Private Sub Command15_Click()
Dim d0, d1

    Call form1.crShell("sqlupd", False)
    d0 = Time
    Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 5
    On Error Resume Next
    Kill "sqlupd.txt"
    On Error GoTo 0
    Call Timer1_Timer

End Sub

Private Sub Command16_Click()
Dim r As ADODB.Recordset, i%, tb$, tt$, was$, rrr
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Command16_Click"
i% = List1.ListIndex
If i% < 0 Then Exit Sub
tb$ = "auftrittstypen"
tt$ = "auftritthigru"
If i% < notb4% Then tb$ = "adresstypen"

was$ = List1.List(List1.ListIndex)

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM auftrittsfelder where typ='" + was$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  MsgBox "Es existieren Felder für diesen Typ. Löschen verweigert."
  Exit Sub
End If

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT count(*) as cnt FROM auftritthigru where auftrittstyp='" + was$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If r!cnt > 0 Then
  MsgBox "Es existieren Daten dieses Typs. Löschen verweigert."
  Exit Sub
End If
form1.sqlqry ("delete from " + tb$ + " where id='" + was$ + "'")
Call rlist1
If i% < List1.ListCount - 1 And i% >= 0 Then
  List1.ListIndex = i% + 1
End If
End Sub

Private Sub Command17_Click()
Dim neuid As String, r As ADODB.Recordset, c As String, pos%, tb$, rrr
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Command17_Click"
If List1.ListIndex < 0 Then
  MsgBox "Klicken Sie erst in die Liste, um Adresse oder Termin zu festzulegen"
  Exit Sub
End If
neuid = InputBox(transe("Neuer Hintergrunddatentyp"), transe("Neuer Hintergrunddatentyp"))
neuid = strrepl(neuid, "-", "")
neuid = strrepl(neuid, " ", "")
If trm(neuid) = "" Then Exit Sub

If List1.ListIndex < notb4% Then
  pos% = (notb4% - 1) * 10
  tb$ = "adresstypen"
Else
  tb$ = "auftrittstypen"
  pos% = (List1.ListCount + 1) * 10
End If
Call form1.sqlqry("insert into " + tb$ + " (id,sortierung) values('" + _
               neuid + "'," + trm(str$(pos%)) + ")")
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM auftrittsfelder where typ='" + List1.List(List1.ListIndex) + "'", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  c = "insert into auftrittsfelder (id,typ,feldname,zeilen,position) values('" + _
    form1.newid("auftrittsfelder", "id", 18) + "','" + _
    neuid + "','" + r!feldname + "'," + trm(r!zeilen) + "," + trm(r!Position) + ")"
  Call form1.sqlqry(c)
  r.MoveNext
Wend
Call rlist1

End Sub

Private Sub Command18_Click()
Dim i As Integer, id$

MousePointer = 11: DoEvents
For i = 0 To List2.ListCount - 1
  id$ = List2.List(i)
  id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
  form1.sqlqry ("update auftrittsfelder set position=" & i & " where id='" & id$ + "'")
Next i
MousePointer = 0

End Sub

Private Sub Command19_Click()
Dim o%, u%, i%, X, tn

Call Command18_Click
DoEvents
auftrittshintergrund.MousePointer = 11: DoEvents
On Error Resume Next
Kill "txtmrk.txt"
On Error GoTo 0

o% = FreeFile

u% = FreeFile
Open "txtmrk.txt" For Output As #u%
For i% = 0 To List2.ListCount - 1
  tn = cut_d1(List2.List(i), " ")
  Print #u%, "this__" + List1.List(List1.ListIndex) + "__" + tn
Next i%
Close #u%
auftrittshintergrund.MousePointer = 0
X = Shell("notepad.exe txtmrk.txt", 1)
End Sub

Private Sub Command2_Click()
Dim anz0, rrr, d0, d1, anz1

auftrittshintergrund.MousePointer = 11
If List1.ListIndex < notb4% Then
  Call form1.sqlqry("insert into auftrittsfelder (id,typ,FeldName,zeilen,position) values('" + form1.newid("auftrittsfelder", "id", 20) + "','" + cid$ + "','NeuesFeld',1," & trm(List2.ListCount) + ")")
  Call rlist2
  auftrittshintergrund.MousePointer = 0
  Exit Sub
End If

If Left$(cid$, 2) <> "--" Then
  On Error Resume Next
  anz0 = sqla.TableDefs("usr_" & utabn(cid$)).Fields.Count
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then anz0 = 0
  Call form1.chgcreate("FLUSH TABLES;")
  Call form1.chgappend("ALTER TABLE usr_" & utabn(cid$) + " ADD NeuesFeld LONGTEXT;")
  Call form1.crShell("sqlchg", False)
  d0 = Time
  Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 4
  sqla.Close
  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
  anz1 = sqla.TableDefs("usr_" & utabn(cid$)).Fields.Count
  If anz1 = anz0 + 1 Then
    Call form1.sqlqry("insert into auftrittsfelder values('" + form1.newid("auftrittsfelder", "id", 20) + "','" + cid$ + "','NeuesFeld',1," & trm(List2.ListCount) + ")")
    On Error Resume Next
    Kill "sqlchg.txt"
    On Error GoTo 0
  Else
    MsgBox "fehlgeschlagen, starten sie zuerst den Server neu."
  End If
  Call rlist2
End If
auftrittshintergrund.MousePointer = 0
End Sub


Private Sub Command3_Click()
Dim neuid As String, pos%, tb$

If List1.ListIndex < 0 Then
  MsgBox "Klicken Sie erst in die Liste, um Adresse oder Termin zu festzulegen"
  Exit Sub
End If
neuid = InputBox(transe("Neuer Hintergrunddatentyp"), transe("Neuer Hintergrunddatentyp"))
neuid = strrepl(neuid, "-", "")
neuid = strrepl(neuid, " ", "")
If trm(neuid) = "" Then Exit Sub

If List1.ListIndex < notb4% Then
  pos% = (notb4% - 1) * 10
  tb$ = "adresstypen"
Else
  tb$ = "auftrittstypen"
  pos% = (List1.ListCount + 1) * 10
End If
Call form1.sqlqry("insert into " + tb$ + " (id,sortierung) values('" + _
               neuid + "'," + trm(str$(pos%)) + ")")
Call form1.sqlqry("insert into auftrittsfelder values('" + _
   form1.newid("auftrittsfelder", "id", 20) + "','" + neuid + "','NeuesFeld',1," & trm(List2.ListCount) + ")")

Call rlist1

End Sub

Private Sub Command4_Click()
Dim i%, id$
i% = List2.ListIndex
If i% < 0 Then Exit Sub

id$ = List2.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
form1.sqlqry ("update auftrittsfelder set position=" & i% & " where id='" & id$ + "'")

id$ = List2.List(i% - 1)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
form1.sqlqry ("update auftrittsfelder set position=" & i% + 1 & " where id='" & id$ + "'")
Call rlist2
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), "." + Text1.text) > 0 Or word1(List2.List(i%)) = Text1.text Then
    List2.ListIndex = i%
    Exit For
  End If
Next i%

End Sub

Private Sub Command5_Click()
Dim i%, id$

i% = List2.ListIndex
If i% >= List2.ListCount - 1 Then Exit Sub

id$ = List2.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
form1.sqlqry ("update auftrittsfelder set position=" & i% + 2 & " where id='" & id$ + "'")
id$ = List2.List(i% + 1)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
form1.sqlqry ("update auftrittsfelder set position=" & i% + 1 & " where id='" & id$ + "'")
Call rlist2
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), "." + Text1.text) > 0 Or word1(List2.List(i%)) = Text1.text Then
    List2.ListIndex = i%
    Exit For
  End If
Next i%
End Sub

Private Sub Command6_Click()
Dim was$, i%, tb$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
tb$ = "auftrittstypen"
If i% < notb4% Then tb$ = "adresstypen"
was$ = List1.List(List1.ListIndex)
form1.sqlqry ("update " + tb$ + " set sortierung=sortierung - 15 where id='" + was$ + "'")
Call rlist1
If i% < List1.ListCount And i% > 0 Then
  List1.ListIndex = i% - 1
End If

End Sub

Private Sub Command7_Click()
Dim was$, i%, tb$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
tb$ = "auftrittstypen"
If i% < notb4% Then tb$ = "adresstypen"

was$ = List1.List(List1.ListIndex)
form1.sqlqry ("update " + tb$ + " set sortierung=sortierung+15 where id='" + was$ + "'")
Call rlist1
If i% < List1.ListCount - 1 And i% >= 0 Then
  List1.ListIndex = i% + 1
End If
End Sub

Private Sub Command8_Click()
Dim id$, z, fnam$, typ$, cmd$, p$, anz0, d0, d1, anz1, i%

id$ = Text3.text
If Len(id$) = 0 Then Exit Sub
MousePointer = 11: DoEvents
Text3.text = ""
z = Val(Text2.text)
fnam$ = Text1.text
typ$ = List1.List(List1.ListIndex)
If Combo1.text <> "" Then fnam$ = Combo1.text + "." + fnam$
If Combo2.text <> "" Then fnam$ = fnam$ + "." + Combo2.text
cmd$ = "update auftrittsfelder set feldname='" + fnam$ + "' where id='" & id$ & "'"
form1.sqlqry (cmd$)
If Check1.value <> 0 Then
  If Len(p$) > 0 And Len(prvt$) <> 0 Then
    form1.sqlqry ("update auftritthigru set feldname='" + p$ + "' where feldname='" + prvt$ + "' and auftrittstyp='" + typ$ + "'")
  End If
End If
If Not form1.isfieldmissing("auftrittsfelder", "opthordeshare") Then
  cmd$ = "update auftrittsfelder set opthordeshare="
  If Check2.value = 0 Then
    cmd$ = cmd$ + "0"
  Else
    cmd$ = cmd$ + "1"
  End If
  cmd$ = cmd$ + " where id='" & id$ + "'"
  form1.sqlqry (cmd$)
End If

If z <> 0 Then
  cmd$ = "update auftrittsfelder set zeilen=" + trm(z) + " where id='" & id$ + "'"
Else
  If z = 0 Then
    anz0 = 10
    If List1.ListIndex >= notb4% Then
      anz0 = sqla.TableDefs("usr_" & utabn(cid$)).Fields.Count
      Call form1.chgcreate("FLUSH TABLES;")
      Call form1.chgappend("FLUSH STATUS;")
      Call form1.chgappend("ALTER TABLE usr_" & utabn(cid$) + " drop " + fnam$ + ";")
      Call form1.chgappend("FLUSH TABLES;")
      Call form1.chgappend("FLUSH STATUS;")
      Call form1.crShell("sqlchg", False)
      d0 = Time
      Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86400 > 5
      sqla.Close
      Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
      anz1 = sqla.TableDefs("usr_" & utabn(cid$)).Fields.Count
    Else
      anz1 = anz0 - 1
    End If
    If anz1 = anz0 - 1 Then
      cmd$ = "delete from auftrittsfelder where id='" & id$ + "'"
      If Check1.value <> 0 Then
        Call form1.sqlqry(cmd$)
        cmd$ = "delete from auftritthigru where feldname='" + fnam$ + "' and auftrittstyp='" + typ$ + "'"
      End If
    Else
      cmd$ = ""
      MsgBox "fehlgeschlagen, starten sie zuerst den Server neu."
    End If
  End If
End If
If cmd$ <> "" Then form1.sqlqry (cmd$)
cmd$ = Text1.text
Text1.text = ""
Text2.text = ""
Text3.text = ""
Call rlist2
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), "." + cmd$ + ".") > 0 Or word1(List2.List(i%)) = cmd$ Then
    List2.ListIndex = i%
    Exit For
  End If
Next i%
MousePointer = 0
End Sub

Private Sub Command9_Click()

Text1.Enabled = True
Text1.SetFocus
prvt$ = Text1.text

End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Form_Load"
Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
nosv4 = 0
auftstart = 9999

Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
On Error Resume Next
Kill "sqlchg.txt"
On Error GoTo 0
Call Timer1_Timer
Dim usr$

Show
Combo1.Clear
Combo2.Clear
Combo1.AddItem "adrselect"
Combo1.AddItem "programm"
Combo1.AddItem "finanzen"
Combo1.AddItem "tabelle"
Combo1.AddItem "besetzung"
Combo1.AddItem "dekade"
If Not form1.isfieldmissing("opt_vnr", "id") Then Combo1.AddItem "Vertragsnummer"
Command8.Enabled = False
Text1.text = ""
Text2.text = ""
Call rlist1
If exist(form1.getmymysqld()) = 0 Then
  Load einstellungen
  einstellungen.SetFocus
  MsgBox form1.inmylanguage("Bitte tragen sie zuerst in Ihren Einstellungen mysql.exe und den Server ein." + vbCrLf + "Oder ändern Sie nichts an Terminen!")
End If
usr$ = LCase(form1.getuserid())
If usr$ = "www" Or usr$ = "administrator" Then
  Command3.Enabled = True
  Command16.Enabled = True
End If
If Not form1.isfieldmissing("auftrittsfelder", "opthordeshare") Then
  Check2.Enabled = True
End If
End Sub

Private Sub setcurrent(id$)
Dim i%
For i% = 0 To List1.ListCount - 1
  If List1.List(i%) = id$ Then
    List1.ListIndex = i%
    Call List1_Click
    i% = List1.ListCount
  End If
Next i%

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub

Private Sub kfrb_Change(Index As Integer)
Dim i%
If Len(kfrb(i%).text) > 3 Then kfrb(i%).text = "255"
For i% = 0 To 2
  On Error Resume Next
  kfarbe%(i%) = Val(kfrb(i%).text)
  On Error GoTo 0
  If kfarbe%(i%) < 0 Then kfarbe%(i%) = 0
  If kfarbe%(i%) > 255 Then kfarbe%(i%) = 255
Next i%
Label6.BackColor = RGB(kfarbe%(0), kfarbe%(1), kfarbe%(2))

End Sub

Private Sub kfrb_DblClick(Index As Integer)
Call Label6_Click
End Sub

Private Sub Label6_Click()
Load colorsel
colorsel.SetFocus
colorsel.updc (Label6.BackColor)
Timer2.Enabled = True
Timer2.Interval = 1000

End Sub

Private Sub List1_Click()

cid$ = List1.List(List1.ListIndex)

nosv4 = 1
Call rlist2
nosv4 = 0

End Sub

Private Sub List2_Click()
Dim asatz$, id$, fldnam$, z As Integer, fldfromtab$, fldfromfld$, p%, cmd$

Command8.Enabled = False
Combo1.text = ""
Combo2.text = ""

asatz$ = trm(List2.List(List2.ListIndex))
p% = InStr(asatz$, "(ID:")
id$ = Mid$(asatz$, p% + 4)
Text3.text = id$
asatz$ = trm(Left$(asatz$, p% - 1))
p% = InStr(asatz$, " Zeilen: ")
If p% > 0 Then
  z = Val(Mid$(asatz$, InStr(asatz$, " Zeilen: ") + 9))
  fldnam$ = trm(Left$(asatz$, p% - 1))
Else
  fldnam$ = ""
End If
If z < 1 Then z = 1
If z > 99 Then z = 99

Text2.text = z
p% = InStr(fldnam$, ".")
If p% > 0 Then
  Combo1.text = Left$(fldnam$, p% - 1)
  fldnam$ = Mid$(fldnam$, p% + 1)
End If
p% = InStr(fldnam$, ".")
If p% > 0 Then
  Combo2.text = Mid$(fldnam$, p% + 1)
  fldnam$ = Left$(fldnam$, p% - 1)
End If
Text1.text = fldnam$
Command9.Enabled = True
If Not form1.isfieldmissing("auftrittsfelder", "opthordeshare") Then
  cmd$ = "select opthordeshare as wert from auftrittsfelder where id='" + id$ + "'"
  Check2.value = 0
  If form1.get1erg(cmd$) = "1" Then Check2.value = 1
End If
End Sub

Private Sub Text1_Change()
Command8.Enabled = True
End Sub

Private Sub Text1_LostFocus()
Dim r As ADODB.Recordset, p$, typ$, d0, d1, i%
Dim d2infile As String, d2insub As String

d2infile = "auftrittshintergrund": d2insub = "Text1_LostFocus"
auftrittshintergrund.MousePointer = 11
Command9.Enabled = True
Text1.Enabled = False
p$ = Text1.text
If LCase(p$) <> "neuesfeld" Then
  typ$ = List1.List(List1.ListIndex)
  If p$ <> prvt$ Then
    If List1.ListIndex >= notb4% Then
      Call form1.chgcreate("FLUSH TABLES;")
      Call form1.chgappend("FLUSH STATUS;")
      Call form1.chgappend("ALTER TABLE usr_" & utabn(typ$) + " CHANGE " + prvt$ + " " + p$ + " LONGTEXT;")
      Call form1.chgappend("FLUSH TABLES;")
      Call form1.chgappend("FLUSH STATUS;")
      Call form1.crShell("sqlchg", False)
      d0 = Time
      Do: d1 = Time: DoEvents: Loop Until (d1 - d0) * 86200 > 4
      sqla.Close
      Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
      For i% = 0 To sqla.TableDefs("usr_" & utabn(typ$)).Fields.Count - 1
        If sqla.TableDefs("usr_" & utabn(typ$)).Fields(i%).name = p$ Then
          i% = sqla.TableDefs("usr_" & utabn(typ$)).Fields.Count + 100
        End If
      Next i%
      If i% > sqla.TableDefs("usr_" & utabn(typ$)).Fields.Count + 10 Then
        On Error Resume Next
        Kill "sqlchg.txt"
        On Error GoTo 0
        Call form1.sqlqry("update auftritthigru set feldname='" + p$ + "' where auftrittstyp='" + typ$ + "' and feldname='" + prvt$ + "'")
      Else
        MsgBox "fehlgeschlagen, starten sie zuerst den Server neu."
      End If
    Else
      Call form1.sqlqry("update auftritthigru set feldname='" + p$ + "' where auftrittstyp='" + typ$ + "' and feldname='" + prvt$ + "'")
    End If
  End If
End If
auftrittshintergrund.MousePointer = 0

End Sub

Private Sub Text2_Change()
Command8.Enabled = True
If Text2.text = "0" Then
  Combo1.text = ""
  Combo2.text = ""
End If
End Sub

Private Sub Text4_change()

If nosv4 = 1 Then Exit Sub
If trm(Text4.text) = "" Then Exit Sub
If List1.ListIndex < 0 Then Exit Sub

Call form1.sqlqry("update auftrittstypen set abkz='" & trm(Text4.text) & "' where id='" & List1.List(List1.ListIndex) & "'")
End Sub

Private Sub Text5_Change()
Dim i%

If Len(Text5.text) <> 6 Then Exit Sub
For i% = 0 To 2
  On Error Resume Next
  kfrb(i%).text = hex2dec(Mid$(Text5.text, i% * 2 + 1, 2))
  On Error GoTo 0
Next i%
Label6.BackColor = RGB(kfarbe%(0), kfarbe%(1), kfarbe%(2))

End Sub

Private Sub Timer1_Timer()
If exist("sqlupd.txt") = 1 Then
  Command15.Enabled = True
Else
  Command15.Enabled = False
End If
If exist("sqldel.txt") = 1 Then
  Command14.Enabled = True
Else
  Command14.Enabled = False
End If

End Sub


Private Sub Timer2_Timer()
Dim c As Long
Dim w As Long, r As Long, g As Long, b As Long

c = form1.getcolorselected()

If c < -10 Then Exit Sub
Timer2.Enabled = False
If c < 0 Then Exit Sub

b = c / 65536
w = c Mod 65536
g = w / 256
r = w Mod 256
kfrb(0).text = r
kfrb(1).text = g
kfrb(2).text = b
End Sub
