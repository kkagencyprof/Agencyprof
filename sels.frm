VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form sels 
   Caption         =   "Gespeicherte Selektionen"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "Vertrags- nummernliste"
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5953
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
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sels.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Selektion ausführen, Ergebnis per Email im Agencyprof-Format"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3840
      Picture         =   "sels.frx":00B2
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Selektion ausführen, CSV-Datei speichern"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox svas 
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00000000&
      Picture         =   "sels.frx":0565
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Selektion anzeigen"
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      MaskColor       =   &H00000000&
      Picture         =   "sels.frx":0A97
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Speichern"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Text            =   "sels.frx":0E3E
      Top             =   120
      Width           =   6735
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
      Height          =   495
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "sels.frx":0E44
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3600
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3315
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "speichern als"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "sels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub rlist1()
Dim tr As String

List1.Clear
Text1.text = ""
tr = Dir(form1.s0dir() & "\*.sqs")
While tr <> ""
  List1.AddItem tr
  tr = Dir
Wend
tr = Dir(form1.mydatadir() & "\*.sqs")
While tr <> ""
  List1.AddItem tr
  tr = Dir
Wend
tr = Dir(form1.vorlagendir() & "\Reportvorlage_*.ht*")
While tr <> ""
  List1.AddItem Mid(tr, 15)
  tr = Dir
Wend
tr = Dir(form1.vorlagendir() & "\*.sqs")
While tr <> ""
  List1.AddItem tr
  tr = Dir
Wend

End Sub

Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command19_Click()
Call form1.handbuchcall("04-Hauptformular.htm")

End Sub

Private Sub Command2_Click()
Dim c$, r As ADODB.Recordset, i%, o%, X, fn$, colHeader, lvitem, rrr
Dim d2infile As String, d2insub As String, fwert As Variant
Dim cnvlist(99) As String, sp%, Y%, res$, zellwert As String, ask%

d2infile = "sels": d2insub = "Command2_Click"
If gd1.Visible = False Then


c$ = Text1.text
If LCase(word1(c$)) = "delete" Then
  ask% = MsgBox("Well then, a chance to quit..." + vbCrLf + "Stop here!", vbYesNo + vbCritical + vbDefaultButton1, "Check that you have a VALID backup.")
  If ask% = vbNo Then
    ask% = MsgBox("You have been warned. I will do this:" + vbCrLf + c$, vbYesNo + vbCritical + vbDefaultButton2, "Check that you have a VALID backup.")
    If ask% = vbYes Then
      Call form1.sqlqry(c$)
      MsgBox "Done, have fun."
    End If
  End If
  Exit Sub
End If
If LCase(word1(c$)) <> "select" Then
  Call html_vorlage
  Exit Sub
End If
MousePointer = 11: DoEvents
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
MousePointer = 0: DoEvents
If rrr <> 0 Then
  MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr)
  Exit Sub
End If
gd1.Visible = True
gd1.View = lvwReport
MousePointer = 11: DoEvents
Y% = 0
For i% = 0 To r.Fields.Count - 1
  'Print #o%, """"; r.Fields(i%).Name; """;";
  Set colHeader = gd1.ColumnHeaders.add(, , r.Fields(i%).name, 800)
  cnvlist(i%) = "": If form1.getusersetting("selcnv" + trm(r.Fields(i%).name), "") = "Datum" Then cnvlist(i%) = "Datum"
  If form1.getusersetting("selcnv" + trm(r.Fields(i%).name), "") = "Uhrzeit" Then cnvlist(i%) = "Uhrzeit"
Next i%
While Not r.EOF
  Set lvitem = gd1.ListItems.add(, , trm(r.Fields(0).value))
  For i% = 1 To r.Fields.Count - 1
    fwert = trm(r.Fields(i%).value)
    If cnvlist(i%) = "Datum" Then
      fwert = datfromsql(fwert)
    End If
    If cnvlist(i%) = "Uhrzeit" Then
      fwert = trm(strrepl(abziffer(trm(fwert)), "Uhr", ""))
      fwert = fwert + ":00"
    End If
    lvitem.SubItems(i%) = fwert
  Next i%
  Y% = Y% + 1
  r.MoveNext
Wend
For i% = 0 To Y% - 1
  Set lvitem = gd1.ListItems(i% + 1)
  For sp% = 0 To r.Fields.Count - 1
    If sp% = 0 Then
      zellwert = lvitem.text
    Else
      zellwert = lvitem.SubItems(sp%)
    End If
    If InStr(zellwert, "select ") > 0 Then
      res$ = ""
      c$ = zellwert
      While c$ <> ""
        fn$ = cut_d1(c$, "|")
        c$ = cut_d2bis(c$, "|")
        If InStr(fn$, "select ") = 1 Then
          fn$ = form1.get1erg(fn$)
          If fn$ <> "" Then res$ = res$ + fn$ + " "
        Else
          res$ = res$ + trm(fn$) + " "
        End If
      Wend
      If sp% > 0 Then
        lvitem.SubItems(sp%) = res$
      Else
        lvitem.text = res$
      End If
      DoEvents
    End If
  Next sp%
Next i%
MousePointer = 0

Else
  gd1.ColumnHeaders.Clear
  gd1.ListItems.Clear
  gd1.Visible = False
End If

End Sub

Private Sub Command24_Click()
Dim i%, c$, tg$, p%, r As ADODB.Recordset, rrr, l$, tg0$, o%

Dim d2infile As String, d2insub As String

d2infile = "sels": d2insub = "Command24_Click"
i% = List1.ListIndex
If i% < 0 Then Exit Sub
c$ = Text1.text
If LCase(word1(c$)) <> "select" Then
  MsgBox "Sorry, nur ""SELECT"" ist möglich"
  Exit Sub
End If
MousePointer = 11: DoEvents
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
MousePointer = 0: DoEvents
If rrr <> 0 Then
  Close #o%
  MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr)
  Exit Sub
End If
If Not r.EOF Then
  MousePointer = 11: DoEvents
  On Error Resume Next
  Kill form1.mydatadir() & "\*.sql"
  On Error GoTo 0
  While Not r.EOF
    Call form1.sqlex_adresse("adresse", "id", r.Fields(0).value)
    r.MoveNext
  Wend
  Load smtp
  On Error Resume Next
  Call smtp.SetFocus
  On Error GoTo 0
  tg0$ = form1.mydatadir() & "\" & strrepl(basename(List1.List(i%), ".sqs"), " ", "_") & ".sql"
  o% = FreeFile
  Open tg0$ For Output As #o%
  tg$ = Dir(form1.mydatadir() & "\*.sql")
  While tg$ <> ""
    If form1.mydatadir() & "\" & tg$ <> tg0$ Then
    On Error Resume Next
    p% = FreeFile
    Open form1.mydatadir() & "\" & tg$ For Input As #p%
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      While Not EOF(p%)
        Line Input #p%, l$
        Print #o%, l$
      Wend
      Close #p%
      On Error Resume Next
      Kill form1.mydatadir() & "\" & tg$
      On Error GoTo 0
    End If
    End If
    tg$ = Dir
  Wend
  Close #o%
  smtp.txtMessageSubject = "Agencyprof Datenpakete aus einer Selektion: " & List1.List(i%)
  smtp.txtMessageText = "Speichern Sie die Attachments in Ihrem Agencyprof-Verzeichnis"
  tg$ = Dir(form1.mydatadir() & "\*.sql")
  While tg$ <> ""
    Call smtp.attachfile(form1.mydatadir() & "\" & tg$)
    tg$ = Dir
  Wend
  MousePointer = 0
End If

End Sub

Private Sub Command3_Click()
Dim c$, r As ADODB.Recordset, i%, o%, X, fn$, rrr, xld$, zelle$, res$, zz$
Dim cnvlist$(19)

c$ = Text1.text
If LCase(word1(c$)) <> "select" Then
  MsgBox transe("Sorry, nur ""SELECT"" ist möglich")
  Exit Sub
End If
xld$ = form1.getusersetting("exceldelimiter", ",")
o% = FreeFile
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) <> "" Then

MousePointer = 11: DoEvents
Open fn$ For Output As #o%
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, "", "")
MousePointer = 0: DoEvents
If rrr <> 0 Then
  Close #o%
  MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr)
  Exit Sub
End If
MousePointer = 11: DoEvents
For i% = 0 To r.Fields.Count - 1
  Print #o%, """"; r.Fields(i%).name; """" + xld$;
  cnvlist(i%) = "": If form1.getusersetting("selcnv" + trm(r.Fields(i%).name), "") = "Datum" Then cnvlist(i%) = "Datum"
  If form1.getusersetting("selcnv" + trm(r.Fields(i%).name), "") = "Uhrzeit" Then cnvlist(i%) = "Uhrzeit"
Next i%
Print #o%,
While Not r.EOF
  For i% = 0 To r.Fields.Count - 1
    zelle$ = trm(r.Fields(i%).value)
    If InStr(zelle$, "select ") > 0 Then
      res$ = ""
      c$ = zelle$
      While c$ <> ""
        zz$ = cut_d1(c$, "|")
        c$ = cut_d2bis(c$, "|")
        If InStr(zz$, "select ") = 1 Then
          zz$ = form1.get1erg(zz$)
          If zz$ <> "" Then res$ = res$ + zz$ + " "
        Else
          res$ = res$ + trm(zz$) + " "
        End If
      Wend
      zelle$ = res$
      DoEvents
    End If
    If cnvlist(i%) = "Datum" Then
      zelle$ = datfromsql(zelle$)
    End If
    If cnvlist(i%) = "Uhrzeit" Then
      zelle$ = strrepl(abziffer(trm(zelle$)), "Uhr", "")
      zelle$ = zelle$ + ":00"
    End If
    Print #o%, """"; zelle$; """" + xld$;
  Next i%
  Print #o%,
  r.MoveNext
Wend
Close #o%
X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)
MousePointer = 0
End If

End Sub

Private Sub Command4_Click()
Dim fn$, o%, i%, sva$

Dim d2infile As String, d2insub As String

d2infile = "sels": d2insub = "Command4_Click"
sva$ = trm(svas.text)
If InStr(LCase(sva$), ".sqs") = 0 Then sva$ = sva$ + ".sqs"
If sva$ <> "" Then
  o% = FreeFile
  fn$ = sva$
  Open fn$ For Output As #o%
  Print #o%, Text1.text;
  Close #o%
Else
  MsgBox transe("kein Dateiname!")
End If
BackColor = form1.cleancolor()
Command4.Enabled = False
Call rlist1
End Sub

Private Sub Command6_Click()
Dim c$, r As ADODB.Recordset, ra As ADODB.Recordset, rrr
Dim bez$, ort$, dtg$, doc$, erst$, ownr$, betr$, typ$, xld$, o%, fn$, pr As Boolean

xld$ = form1.getusersetting("exceldelimiter", ",")
o% = FreeFile
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) <> "" Then
MousePointer = 11: DoEvents
Open fn$ For Output As #o%
c$ = """Vertr.Nr.""" + xld$ + """Datum""" + xld$ + """Bezeichnung""" + xld$ + """Ort""" + xld$ + _
         """Dokument""" + xld$ + """erstellt""" + xld$ + _
         """Benutzer"""
Print #o%, c$

c$ = "select * from opt_vnr order by sortnr"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly)
If rrr = 0 Then
  While Not r.EOF
    bez$ = "fehlt"
    dtg$ = ""
    ort$ = ""
    typ$ = ""
    c$ = "select datum,ort,bezeichnung,auftrittstyp from auftritt where id='" + trm(r!aid) + "'"
    Set ra = New ADODB.Recordset
    ra.CursorLocation = adUseServer
    rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly)
    If rrr = 0 Then
      If Not ra.EOF Then
        bez$ = trm(ra!bezeichnung)
        dtg$ = trm(ra!datum)
        ort$ = trm(ra!ort)
        typ$ = trm(ra!auftrittstyp)
      End If
    End If
    doc$ = "unbekannt"
    erst$ = ""
    ownr$ = ""
    betr$ = ""
    pr = False
    c$ = "select erstellt,owner,betreff,docname from dochist where doctyp='Vertragsnummer " + trm(r!id) + "' order by erstellt"
    Set ra = New ADODB.Recordset
    ra.CursorLocation = adUseServer
    rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly)
    If rrr = 0 Then
      While Not ra.EOF
        doc$ = "unbekannt"
        erst$ = ""
        ownr$ = ""
        betr$ = ""
        pr = True
        erst$ = trm(ra!erstellt)
        ownr$ = trm(ra!Owner)
        betr$ = trm(ra!betreff)
        doc$ = trm(ra!docname)
        c$ = """" + trm(r!id) + """" + xld$ + """" + dtg$ + """" + xld$ + """" + bez$ + """" + xld$ + """" + ort$ + """" + xld$ + _
           """" + doc$ + """" + xld$ + """" + erst$ + """" + xld$ + _
           """" + ownr$ + """"
        Print #o%, c$
        ra.MoveNext
      Wend
    End If
    If Not pr Then
      c$ = """" + trm(r!id) + """" + xld$ + """" + dtg$ + """" + xld$ + """" + bez$ + """" + xld$ + """" + ort$ + """" + xld$ + _
         """" + doc$ + """" + xld$ + """" + erst$ + """" + xld$ + _
         """" + ownr$ + """"
      Print #o%, c$
    End If
    r.MoveNext
  Wend
End If
Close #o%
MousePointer = 0
o% = Int(Shell("notepad.exe " + fn$, vbNormalFocus))
End If
End Sub

Private Sub Form_Load()
Dim dbpara$

axsResizer1.SaveControlPositions
sels.Caption = transe("Gespeicherte Selektionen")
Command24.ToolTipText = transe("Selektion ausführen, Ergebnis per Email im Agencyprof-Format")
Command3.ToolTipText = transe("Selektion ausführen, CSV-Datei speichern")
Command2.ToolTipText = transe("Selektion anzeigen")
Command4.ToolTipText = transe("Speichern")
Command19.ToolTipText = transe("Hilfeseite öffnen")
Label1.Caption = transe("speichern als")

Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

gd1.Visible = False
If form1.isfieldmissing("opt_vnr", "id") Then Command6.Visible = False

BackColor = form1.cleancolor()
Call rlist1
BackColor = form1.cleancolor()

End Sub
Private Sub Form_Resize()

axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub

Private Sub List1_Click()
Dim i%, o%, fn$, c$, r As ADODB.Recordset, M$, rrr

Text1.text = ""
i% = List1.ListIndex
If i% < 0 Then Exit Sub
Label1.Caption = ""
fn$ = form1.mydatadir() & "\" & List1.List(i%)
If nexist(fn$) Then fn$ = form1.s0dir() & "\" & List1.List(i%)
If nexist(fn$) Then fn$ = form1.vorlagendir() & "\" & List1.List(i%)
If nexist(fn$) Then fn$ = form1.vorlagendir() & "\Reportvorlage_" & List1.List(i%)
svas.text = fn$
o% = FreeFile
Open fn$ For Input As #o%
While Not EOF(o%)
  Line Input #o%, fn$
  If Text1.text <> "" Then Text1.text = Text1.text & vbCrLf
  Text1.text = Text1.text & fn$
Wend
Close #o%

c$ = Text1.text
If LCase(word1(c$)) = "select" Then
  c$ = Mid$(c$, InStr(LCase(c$), "from "))
  c$ = "select count(*) as cnt " & c$
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly)
  If rrr = 0 Then
    M$ = r!cnt & " Sätze"
  Else
    M$ = "Anzahl nicht feststellbar"
  End If
  Label2.Caption = M$
End If
BackColor = form1.cleancolor()
Command4.Enabled = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, ask

If KeyCode = 8 Or KeyCode = 46 Then
  i% = List1.ListIndex
  If i% < 0 Then Exit Sub
  ask = MsgBox(transe("Datei löschen: ") + " " & List1.List(i%) + "?", vbYesNo + vbCritical + vbDefaultButton2, "Vorhandene Datei löschen?")
  If ask = vbNo Then Exit Sub
  On Error Resume Next
  Kill form1.mydatadir() & "\" & List1.List(i%)
  On Error GoTo 0
  List1.RemoveItem i%
  Text1.text = ""
  BackColor = form1.cleancolor()
  Command4.Enabled = False
End If

End Sub

Private Sub svas_Change()

Command4.Enabled = True
BackColor = form1.dirtycolor()

End Sub

Private Sub Text1_Change()

Command4.Enabled = True
BackColor = form1.dirtycolor()
End Sub

Private Sub html_vorlage()
Dim i As Integer, n As Integer, c$, tx$, sqlcoll As Boolean, sql$, rrr, l$
Dim fn$, o%, r As ADODB.Recordset, k As Integer, perc As Integer, p%, f$

fn$ = form1.myuniquedocname("", "htm")
sqlcoll = False
If fn$ = "" Then Exit Sub
svas.Visible = False
o% = FreeFile
Open fn$ For Output As #o%
tx$ = Text1.text
Text1.text = ""
svas.text = ""
n = linesof(tx$)
For i = 1 To n
  c$ = lineof(i, tx$)
  perc% = 100 * (i / n)
  Text1.text = trm(perc%) + "%" + vbCrLf + c$
  DoEvents
  If trm(c$) = "?>" Then
    sqlcoll = False
    c$ = ""
    sql$ = trm(sql$)
    If LCase(word1(sql$)) <> "select" Then
      f$ = form1.mydatadir() & "\" & sql$
      If nexist(f$) Then f$ = form1.s0dir() & "\" & sql$
      If nexist(f$) Then f$ = form1.vorlagendir() & "\" & sql$
      sql$ = f$
      If Not nexist(sql$) Then
        p% = FreeFile
        Open sql$ For Input As #p%
        sql$ = ""
        While Not EOF(p%)
          Line Input #p%, l$
          If trm(l$) <> "" Then sql$ = sql$ + " " + l$
        Wend
        Close #p%
        sql$ = trm(sql$)
      End If
    End If
    Text1.text = trm(perc%) + "%" + vbCrLf + sql$
    DoEvents
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, sql$, form1.adoc, dbOpenDynaset, dbReadOnly)
    If rrr <> 0 Then
      MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr) & vbCrLf & sql$
      Close #o%
      Exit Sub
    End If
    While Not r.EOF
      Print #o%, "<font size=""1"">"
      For k = 0 To r.Fields.Count - 1
        Print #o%, form1.repl1310htm(trm(r.Fields(k).value) + "&nbsp;")
      Next k
      Print #o%, "</font>"
      Print #o%, "<br>"
      r.MoveNext
    Wend
  End If
  If sqlcoll Then
    sql$ = sql$ + " " + c$
  End If
  If trm(c$) = "<?SQL" Then
    sqlcoll = True
    sql$ = ""
  End If
  If Not sqlcoll Then Print #o%, c$
Next i
Close #o%
svas.Visible = True
'Call form1.openthisdoc(fn$, "")
rrr = Shell("explorer.exe " & fn$, vbNormalFocus)

End Sub
