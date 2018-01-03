VERSION 5.00
Begin VB.Form verwalt_public 
   Caption         =   "Verwaltung"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox liblog 
      Caption         =   "Log"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   33
      Top             =   480
      Value           =   1  'Aktiviert
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "APLibTest"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "Open replication form"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H80000005&
      Caption         =   "Fehler!"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      ToolTipText     =   "Persönliche Alarme und Hintergrund-Datentypen einstellen"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "rebuild mailsafe"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   2400
      Left            =   5280
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Länderliste"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton rdwrkvz 
      Caption         =   "Werkabgleich"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command26 
      Caption         =   "dbstamp"
      Height          =   255
      Left            =   3360
      TabIndex        =   26
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "update Agencyprof"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command23 
      Caption         =   "changelog"
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "zu übersetzen"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Daten testen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Mailadressexport"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "Mailadresssen im CSV-Format exportiern"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox dbgopt 
      Caption         =   "Debug"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Export"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Kassenbuch"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Abos"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "refresh"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6960
      Top             =   120
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Übersetzungstabelle"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "... weiter"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Alles halt"
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Aktuelle Währungskurse anfordern"
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Übersetzungstabelle"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "leere Kontakte u. Hinweise löschen"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Datenimport"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Dump Database"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000005&
      Caption         =   "Ba&ustelle"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Persönliche Alarme und Hintergrund-Datentypen einstellen"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Währungen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&larmlisten"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Hintergrunddatentypen"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Left            =   2400
      Top             =   2640
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   840
      TabIndex        =   34
      ToolTipText     =   "temporary debug"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "online:"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Viewer"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3720
      Picture         =   "verwalt_public.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3120
      Picture         =   "verwalt_public.frx":0C42
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   3120
   End
End
Attribute VB_Name = "verwalt_public"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wrkJet As Workspace
Dim vncv$, vncs$, t1yp$(199)
Dim f_ldn$(199), mnams_engl$(1 To 12), libloglock As Boolean

Private Sub Command1_Click()
Load alarmlist
Call alarmlist.SetFocus
Command2_Click
End Sub

Private Sub Command11_Click()
Dim n$, wert$, p%, c$

wert$ = "neu=new"
wert$ = InputBox("Neue Übersetzung:", "Neuer Eintrag", wert$)
p% = InStr(wert$, "=")
If p% > 0 Then
  n$ = trm(Mid$(wert$, p% + 1))
  wert$ = trm(Left$(wert$, p% - 1))
  c$ = "insert into dictionary (id,translat) values('" & _
                 wert$ & "','" & _
                 n$ & "')"
  Call form1.sqlqry(c$)
  Call rlist1
End If

End Sub

Private Sub Command12_Click()
Load verwaltung
verwaltung.SetFocus

End Sub

Private Sub Command13_Click()
Dim o%

o% = FreeFile
Open form1.s0dir() & "\lock.lck" For Output As #o%
Print #o%, form1.computername
Close #o%

End Sub

Private Sub Command14_Click()
On Error Resume Next
Kill form1.s0dir() & "\lock.lck"
On Error GoTo 0

End Sub

Private Sub Command15_Click()
Dim tr

tr = Dir(form1.s0dir() & "\*.run")
While tr <> ""
  On Error Resume Next
  Kill form1.s0dir() & "\" & tr
  On Error GoTo 0
  tr = Dir
Wend
Call rlist2

End Sub

Private Sub Command16_Click()
Load abos
On Error Resume Next
Call abos.SetFocus
On Error GoTo 0

End Sub

Private Sub Command17_Click()
Load kbuch
    On Error Resume Next
    Call kbuch.SetFocus
    On Error GoTo 0

End Sub

Private Sub Command18_Click()
Dim rrr, c$, o%, hdr$, l$, hd%, frm$, cc$, an$, ccl$, anl$, sbj$, i%
Dim t As ADODB.Recordset, efn$, ut$, ut1$, ut2$, utyp As String

c$ = "select * from mailsafe order by erstellt"
Set t = New ADODB.Recordset
t.CursorLocation = adUseServer
t.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
While Not t.EOF
  c$ = word1(trm(t!erstellt))
  If Command18.Caption <> c$ Then Command18.Caption = c$
  efn$ = form1.composeemlname(trm(t!message))
  If efn$ = "" Then efn$ = trm(t!message)
  Debug.Print efn$
  If Not nexist(efn$) Then
    hdr$ = ""
    o% = FreeFile
    Open efn$ For Input As #o%
    hd% = 1
    While Not EOF(o%) And hd% = 1
      Line Input #o%, l$: hdr$ = hdr$ + vbCrLf + l$
      If l$ = "" And hd% = 1 Then
        hd% = 0
      End If
    Wend
    Close #o%
    frm$ = emailonly(GetHeaderValue(hdr$, "From"))
    If frm$ = "" Then frm$ = emailonly(GetHeaderValue(hdr$, "Return-Path"))
    sbj$ = GetHeaderValue(hdr$, "Subject")
    If InStr(LCase(sbj$), "utf") > 0 Then
      ut1$ = utf8sbjdecode(sbj$): If ut1$ <> "" Then sbj$ = ut1$
    End If
    If InStr(LCase(sbj$), "iso-8859") > 0 Then
      sbj$ = QuotedPrintableDecode(sbj$)
    End If
    If trm(sbj$) <> trm(t!Subject) Then
      Debug.Print sbj$
      c$ = "update mailsafe set Subject='" + strrepl(Left$(sbj$, 240), "'", "´") + "' where id='" + t!id + "'"
      Call form1.sqlqry(c$)
    End If
    cc$ = GetHeaderValue(hdr$, "CC"): ccl$ = ""
    While cc$ <> ""
      c$ = cut_d1(cc$, ","): cc$ = cut_d2bis(cc$, ",")
      If ccl$ <> "" And Right(trm$(ccl$), 1) <> "," Then
        ccl$ = ccl$ + ","
      End If
      ccl$ = ccl$ + emailonly(c$)
    Wend
    an$ = GetHeaderValue(hdr$, "To"): anl$ = ""
    While an$ <> ""
      c$ = cut_d1(an$, ","): an$ = cut_d2bis(an$, ",")
      If anl$ <> "" And Right(trm$(anl$), 1) <> "," Then
        anl$ = anl$ + ","
      End If
      anl$ = anl$ + emailonly(c$)
    Wend
    If frm$ <> t!frm And frm$ <> "" Then
      Debug.Print "From: " + frm$, t!frm
      c$ = "update mailsafe set frm='" + frm$ + "' where id='" + t!id + "'"
      Call form1.sqlqry(c$)
    End If
    If Not form1.isfieldmissing("mailsafe", "optcc") Then
      If ccl$ <> trm(t!optcc) Then
'        Debug.Print "CC: " + ccl$, t!optcc
        c$ = "update mailsafe set optcc='" + ccl$ + "' where id='" + t!id + "'"
        Call form1.sqlqry(c$)
      End If
    End If
    If Not form1.isfieldmissing("mailsafe", "optan") Then
      If anl$ <> trm(t!optan) Then
'        Debug.Print "To: " + an$, t!optan
        c$ = "update mailsafe set optan='" + Left(anl$, 240) + "' where id='" + t!id + "'"
        Call form1.sqlqry(c$)
      End If
    End If
  End If
  DoEvents
  t.MoveNext
Wend
Command18.Caption = "rebuild mailsafe"
End Sub

Private Sub Command19_Click()
Dim i As Integer, j As Integer, ask%, flst$, fnam
Dim stn$, s As Database, e As Recordset
Dim r As Recordset, t As TableDef, tn$, f As Field, dbname$
Dim idx As New Index, ftyp, o%, fn$, fsiz As Integer

List1.Clear
dbname$ = form1.getdbname()
If InStr(LCase(dbname$), ".mdb") > 0 Then
  Call mdb2sql
  Exit Sub
End If
fn$ = form1.s0dir() + "\" + dbname$ & ".mdb"
If exist(fn$) <> 0 Then
  ask% = MsgBox(fn$ & " existiert - überschreiben?", vbYesNo + vbCritical + vbDefaultButton1, "SQL-Datenbankexport")
  If ask% <> vbYes Then Exit Sub
  Kill fn$
End If
On Error Resume Next
Kill dbname$ & ".err"
On Error GoTo 0
Set s = wrkJet.CreateDatabase(fn$, dbLangGeneral)
s.Close
Set s = wrkJet.OpenDatabase(fn$, False, False)
For i = 0 To form1.sqla.TableDefs.Count - 1
  flst$ = ""
Debug.Print form1.sqla.TableDefs(i).name
  If Left$(LCase(form1.sqla.TableDefs(i).name), 4) <> "msys" _
     And Left$(LCase(form1.sqla.TableDefs(i).name), 6) <> "mysql." Then
    List1.AddItem form1.sqla.TableDefs(i).name
    List1.ListIndex = List1.ListCount - 1
    DoEvents
    Set t = s.CreateTableDef(form1.sqla.TableDefs(i).name)
    For j = 0 To 199: f_ldn$(j) = "": Next j
    For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
      ftyp = form1.sqla.TableDefs(i).Fields(j).Type
      fnam = form1.sqla.TableDefs(i).Fields(j).name
      If ftyp <> 12 Then
        fsiz = form1.sqla.TableDefs(i).Fields(j).Size
      Else
        fsiz = 250
      End If
      If flst$ <> "" Then flst$ = flst$ & ","
      flst$ = flst$ & fnam
      f_ldn$(j) = fnam
      If ftyp = 0 Then ftyp = 3
      t1yp$(j) = trm(ftyp)
      Set f = t.CreateField(fnam, ftyp)
      f.Size = fsiz
      f.Required = False
      t.Fields.Append f
    Next j
    For j = 0 To form1.sqla.TableDefs(i).Indexes.Count - 1
        Set idx = form1.sqla.TableDefs(i).Indexes(j)
        ' *** Create the index
        Set indIndexObj = t.CreateIndex(idx.name)
        indIndexObj.Fields = idx.Fields
        indIndexObj.Unique = idx.Unique
        If idx.name = "PRIMARY" Then indIndexObj.Primary = True
        ' *** Add this index
        t.Indexes.Append indIndexObj
    Next j
    ' *** Append this new table
    s.TableDefs.Append t
    c$ = "select * from " & form1.sqla.TableDefs(i).name
    Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    Set e = s.OpenRecordset(form1.sqla.TableDefs(i).name)
    While Not r.EOF
      e.AddNew
      flst$ = ""
      werte$ = ""
      For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
        If trm(r.Fields(j).value) <> "" Then
        ff$ = "'": If t1yp$(j) = "3" Or t1yp$(j) = "4" Or t1yp$(j) = "7" Then ff$ = ""
        flst$ = flst$ & f_ldn$(j) & ","
        If ff$ = "" Then
          werte$ = werte$ & strrepl(trm(r.Fields(j).value), ",", ".") & ","
        Else
          werte$ = werte$ & ff$ & r.Fields(j).value & ff$ & ","
        End If
        e.Fields(j).value = r.Fields(j).value
        End If
      Next j
      If Len(werte$) > 1 Then
        werte$ = Left$(werte$, Len(werte$) - 1)
        flst$ = Left$(flst$, Len(flst$) - 1)
        c$ = "insert into " & form1.sqla.TableDefs(i).name & " (" & flst$ & ") values(" & werte$ & ");"
      End If
      On Error Resume Next
      e.Update
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        o% = FreeFile
        Open dbname$ & ".err" For Append As #o%
        Print #o%, Error$(rrr); ": "; vbCrLf; c$
        Close #o%
      End If
      r.MoveNext
    Wend
  End If
Next i
If exist(dbname$ & ".err") <> 0 Then
  X = Shell("notepad.exe " & dbname$ & ".err", 1)
  DoEvents
  On Error Resume Next
  Kill dbname$ & ".err"
  On Error GoTo 0
End If
Call rlist1

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command21_Click()
Dim c$, r As Recordset, i%, o%, X, fn$, rrr

o% = FreeFile
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) <> "" Then
Open fn$ For Output As #o%
c$ = "select name, email from adresse where trim(email)<>''"
On Error Resume Next
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Close #o%
  MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr)
  Exit Sub
End If
MousePointer = 11: DoEvents
For i% = 0 To r.Fields.Count - 1
  Print #o%, """"; r.Fields(i%).name; """,";
Next i%
Print #o%,
While Not r.EOF
  For i% = 0 To r.Fields.Count - 1
    Print #o%, """"; trm(r.Fields(i%).value); """,";
  Next i%
  Print #o%,
  r.MoveNext
Wend
c$ = "select name, email from kontakt where trim(email)<>''"
On Error Resume Next
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Close #o%
  MsgBox "Fehler " & rrr & " in der SQL - Abfrage:" & vbCrLf & Error$(rrr)
  Exit Sub
End If
While Not r.EOF
  For i% = 0 To r.Fields.Count - 1
    Print #o%, """"; trm(r.Fields(i%).value); """,";
  Next i%
  Print #o%,
  r.MoveNext
Wend
Close #o%
X = Shell("notepad.exe " & fn$, 1)
MousePointer = 0
End If

End Sub

Private Sub Command22_Click()
Dim r As Recordset, c$, o%, s As Recordset, prv$, ro%, X, i%
Dim ttt As ADODB.Recordset, aid$, kid$

MousePointer = 11: DoEvents
List3.Clear
List3.Visible = True
List3.AddItem "adresstyp": DoEvents
o% = FreeFile
Open "dcheck.txt" For Output As #o%
ro% = FreeFile
Open "checkremarks.txt" For Output As #ro%


If Not form1.isfieldmissing("auftritthigru", "opt_kid") Then
cmd$ = "SELECT auftritthigru.id, auftritthigru.auftrittsid, auftritthigru.FeldName, auftritthigru.FeldDaten, auftritthigru.opt_kid "
cmd$ = cmd$ + "FROM auftritthigru INNER JOIN adresstypen ON auftritthigru.auftrittstyp = adresstypen.id where auftritthigru.auftrittsid not in (select id from adresse where id=auftritthigru.auftrittsid);"
Set t = New ADODB.Recordset
t.CursorLocation = adUseServer
t.Open cmd$, form1.adoc, adOpenDynamic, adLockReadOnly
While Not t.EOF
  If trm(t!opt_kid) = "" Then
Debug.Print t!auftrittsid; " "; t!feldname; " "; t!felddaten
    aid$ = trm(t!auftrittsid)
    For i% = 1 To Len(aid$)
      z$ = Left(aid$, i%)
      cmd$ = "select ID as wert from adresse where id='" + z$ + "'"
      If form1.get1erg(cmd$) <> "" Then
        kid$ = Mid(aid$, i% + 1)
        cmd$ = "update auftritthigru set opt_kid='" + kid$ + "' where id='" + t!id + "'"
        Debug.Print aid$; ": "; cmd$
        Call form1.sqlqry(cmd$)
        Exit For
      End If
    Next i%
  End If
  DoEvents
  t.MoveNext
Wend
End If

If Not form1.isfieldmissing("opt_repertoire", "id") Then
List3.AddItem "repertoire": DoEvents
c$ = "select * from opt_repertoire"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  c$ = "select * from w_loc where id='" & r!wid & "'"
Debug.Print c$
  Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If s.EOF Then
    DoEvents
    Debug.Print "Werk " + trm(r!wid) + " fehlt (Repertoire): Artist=" + trm(r!vid)
    Print #ro%, "Werk " + trm(r!wid) + " fehlt (Repertoire): Artist=" + trm(r!vid)
    Print #o%, "delete from opt_repertoire where id='" & r!id & "';"
    c$ = "select * from programmliste where WerkID='" & r!wid & "'"
Debug.Print c$
    Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    While Not s.EOF
      DoEvents
      Debug.Print "Werk " + trm(r!wid) + " fehlt (Programmliste) Prog=" + trm(s!programmid)
      Print #ro%, "Werk " + trm(r!wid) + " fehlt (Programmliste) Prog=" + trm(s!programmid)
      Print #o%, "delete from programmliste where id='" & r!id & "';"
      s.MoveNext
    Wend
    End If
  r.MoveNext
Wend
End If
Close ro%
X = Shell("notepad.exe checkremarks.txt", 1)
DoEvents
List3.AddItem "kontakt": DoEvents
c$ = "SELECT * from adresstyp ORDER BY vid, kid, typ;"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
prv$ = "-1"
While Not r.EOF
  c$ = r!vid & "|" & r!typ & "|" & r!wert & "|" & r!kid
  If prv$ = c$ Then
    Print #o%, "delete from adresstyp where id='" & r!id & "';"
  End If
  prv$ = c$
  r.MoveNext
Wend
DoEvents
List3.AddItem "kontakt": DoEvents
c$ = "select * from kontakt"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  c$ = "select id from adresse where id='" & r!vid & "'"
  Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If s.EOF Then
'Print #o%, trm("--" & r!vid & " - " & r!Name & " " & r!telfaxhandy)
    DoEvents
    Print #o%, "delete from kontakt where id='" & r!id & "';"
  End If
  r.MoveNext
Wend
List3.AddItem "auftritthigru": DoEvents
prv$ = "--1--0--"
c$ = "SELECT auftritthigru.id, auftritthigru.auftrittsid, auftritthigru.auftrittstyp, auftritthigru.FeldName, auftritthigru.FeldDaten "
c$ = c$ + "FROM auftritthigru INNER JOIN auftrittstypen ON auftritthigru.auftrittstyp = auftrittstypen.id"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If prv$ <> trm(r!auftrittsid) Then
    prv$ = trm(r!auftrittsid)
    c$ = "select id from auftritt where id='" & r!auftrittsid & "'"
    Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    If s.EOF Then
      Print #o%, "delete from auftritthigru where auftrittsid='" & r!auftrittsid & "';"
      Print #o%, "delete from usr_" & utabn(r!auftrittstyp) & " where id='" & r!auftrittsid & "';"
    End If
  End If
  DoEvents
  r.MoveNext
Wend
Close #o%
List3.Clear
List3.Visible = False
MousePointer = 0
X = Shell("notepad.exe dcheck.txt", 1)
End Sub

Private Sub Command23_Click()
Dim s%, t%, s0d$

t% = FreeFile
s0d$ = form1.s0dir() & "\"
Open s0d$ & "changelog.new" For Output As #t%
Print #t%, App.Major & "." & App.Minor & "-" & App.Revision; "-"; datum2sql(Date); ": "
Close #t%
If exist(s0d$ & "changelog.txt") <> 0 Then
  t% = FreeFile
  Open s0d$ & "changelog.new" For Append As #t%
  s% = FreeFile
  Open s0d$ & "changelog.txt" For Input As #s%
  While Not EOF(s%)
    Line Input #s%, l$
    Print #t%, l$
  Wend
  Close #s%
  Close #t%
  Kill s0d$ & "changelog.txt"
End If
Name s0d$ & "changelog.new" As s0d$ & "changelog.txt"
X = Shell("notepad.exe " & s0d$ & "changelog.txt", 1)

End Sub

Private Sub Command24_Click()
List1.Clear
tr = Dir("*.frm")
While tr <> ""
  List1.AddItem tr
  tr = Dir
Wend
End Sub

Private Sub Command25_Click()
Call form1.updateme
End Sub

Public Sub Command26_Click()
Dim i%, j%, o%, dbx$, apv As Long, aplv As Long

MousePointer = 11
DoEvents
apv = App.Revision
aplv = hexstring2dec(bas_getAPLibVersion())
dbx$ = "": If InStr(form1.getdbname(), ".mdb") > 0 Then dbx$ = ".mdb"
fn$ = form1.s0dir() & "\" & App.Major & "-" & App.Minor & "-Agencyprof1.ver"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, apv
Print #o%, aplv
Close #o%
fn$ = App.Major & "-" & App.Minor & dbx$ & ".dbini"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, App.Revision
For i% = 0 To form1.sqla.TableDefs.Count - 1
  If Left$(LCase(form1.sqla.TableDefs(i%).name), 4) <> "msys" Then
    t$ = form1.sqla.TableDefs(i%).name
    tx$ = Left$(t$, 4)
    If tx$ <> "tmp_" And tx$ <> "usr_" And tx$ <> "opt_" Then
      For j% = 0 To form1.sqla.TableDefs(i%).Fields.Count - 1
        f$ = form1.sqla.TableDefs(i%).Fields(j%).name
        ty = form1.sqla.TableDefs(i%).Fields(j%).Type
        le = form1.sqla.TableDefs(i%).Fields(j%).Size
        Print #o%, t$; "."; f$; "."; trm(ty); "."; trm(le)
      Next j%
    End If
  End If
Next i%
Close #o%
MousePointer = 0
X = Shell("notepad.exe " & fn$, 1)

End Sub

Private Sub Command27_Click()
Dim t As ADODB.Recordset
Dim s As ADODB.Recordset

MousePointer = 11: DoEvents
cmd$ = "delete from sysvars where left(owner,28)='sysvar_system_landeskennung_';": Call form1.sqlqry(cmd$)
Set t = New ADODB.Recordset
t.CursorLocation = adUseServer
t.Open "SELECT land FROM adresse order by land desc", form1.adoc, adOpenDynamic, adLockReadOnly
While Not t.EOF And trm(t!land) <> ""
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  s.Open "SELECT wert FROM sysvars where owner='sysvar_system_landeskennung_" & trm(t!land) & "'", form1.adoc, adOpenDynamic, adLockReadOnly
  If s.EOF Then
    c$ = "insert into sysvars (id,owner,wert) values('" & form1.newid("sysvars", "id", 22) & "','sysvar_system_landeskennung_" & trm(t!land) & "','" & trm(t!land) & "')": Call form1.sqlqry(c$)
  End If
  t.MoveNext
Wend
MousePointer = 0
End Sub


Private Sub Command28_Click()

MsgBox trm(1 / 0)
End Sub

Private Sub Command29_Click()
Dim aplv As String, V, ge As String, o%, ierg As Integer, lerg As Long
Dim testtext As String, liblevel As Long, rrr, echodas As String

Load dbupgrade
On Error Resume Next
dbupgrade.SetFocus
On Error GoTo 0
dbupgrade.Caption = "testing agencyproflib.dll"
dbupgrade.List1.Clear
'0x10001
aplv = "not found"
testtext = "qwertzuiopasdfghjklyxcvbnm1234567890"
On Error Resume Next
V = getAPLibVersion
rrr = Err
On Error GoTo 0
If rrr = 0 Then aplv = "0x" + dec2hex(trm(V))
dbupgrade.List1.AddItem "Version: " + aplv
If aplv = "not found" Then Exit Sub
liblevel = hexstring2dec(aplv)
echodas = "0123456789": ge = getAPLibEcho2(ByVal echodas)
If echodas = ge Then
  ge = "ok"
Else
  ge = "failed"
End If
dbupgrade.List1.AddItem "Echotest: " + ge
dbupgrade.List1.AddItem "DateLong(): " + trm(APLibDateLong())
dbupgrade.List1.AddItem "TimeLong(): " + trm(APLibTimeLong()) + " (UTC)"
o% = FreeFile
Open "__b64testfile__.tmp" For Output As #o%
Print #o%, testtext
Close #o%
ierg = APLibEncodeFileB64(ByVal "__b64testfile__.tmp", ByVal "__b64testfile__.tmp.b64")
ge = "ok": If ierg <> 0 Then ge = "failed: " + trm(ierg)
dbupgrade.List1.AddItem "EncodeFileB64: " + ge
ierg = APLibDecodeFileB64(ByVal "__b64testfile__.tmp.b64", ByVal "__b64testfile__.tmp.dec")
ge = "ok": If ierg <> 0 Then ge = "failed: " + trm(ierg)
dbupgrade.List1.AddItem "DecodeFileB64: " + ge
o% = FreeFile
Open "__b64testfile__.tmp.dec" For Input As #o%
Line Input #o%, ge
Close #o%
If ge <> testtext Then
  ge = "failed"
Else
  ge = "ok"
End If
dbupgrade.List1.AddItem "compared files: " + ge
On Error Resume Next
Kill "__b64testfile__.tmp"
Kill "__b64testfile__.tmp.b64"
Kill "__b64testfile__.tmp.dec"
On Error GoTo 0
ierg = APLibIntTest(ByVal "")
dbupgrade.List1.AddItem "IntTest (expecting '1'): '" + trm(ierg) + "'"
testtext = datum2sql(trm(Date)) + " " + trm(Time)
lerg = APLibTimeFromString(ByVal testtext)
dbupgrade.List1.AddItem "TimeFromString(" + testtext + "): " + trm(lerg)
'Debug.Print APLibIstSommerzeit(ByVal dtg$)
'0x10002
If liblevel < 2 Then
  dbupgrade.List1.AddItem "cannot run 0x10002 tests on " + aplv + " (this version), lib is too old."
  Call dbupgrade.addline("Test aborted, update agencyproflib.dll.")
  Exit Sub
End If
ierg = bas_APLibWriteLog(-1)
dbupgrade.List1.AddItem "WriteLogMode is " + trm(ierg)
End Sub

Private Sub Command4_Click()

Load waehrung
Call waehrung.SetFocus
End Sub

Private Sub Command6_Click()
Dim o%, X, r As Recordset, sqla As Database, wrkJet As Workspace
Dim dbpara$

Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
dbpara$ = form1.getconnstr()
If dbpara$ <> "msaccessmdb" Then
  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
Else
  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
End If

For i% = 0 To form1.sqla.TableDefs.Count - 1
  If Left$(LCase(form1.sqla.TableDefs(i%).name), 4) <> "msys" Then
    Call form1.ExportOneTableToExcel(form1.sqla.TableDefs(i%).name)
  End If
Next i%

o% = FreeFile
fn$ = "dump"
Open fn$ + ".bat" For Output As #o%
Print #o%, DirName(form1.getmymysqld()) + "\mysqldump -h " + form1.getmymysqlhost() + " -u root " + form1.getdbname() + " >" + form1.getdbname() + ".txt"
Close #o%
'x = Shell("notepad.exe dump.bat", 1)
X = Shell(fn$ + ".bat", 1)

End Sub

Private Sub Command7_Click()
Load auftrittshintergrund
Call auftrittshintergrund.SetFocus
Call Command2_Click
End Sub

Private Sub Command8_Click()
Load import
On Error Resume Next
Call import.SetFocus
On Error GoTo 0

End Sub

Private Sub Command9_Click()
MousePointer = 11
DoEvents
form1.sqlqry ("delete from kontakt WHERE (((Trim(name))='') AND ((Trim(tel))='') AND ((Trim(fax))='') AND ((Trim(email))='') AND ((kontakt.handy)='0'));")
DoEvents
form1.sqlqry ("delete from auftritthigru where ( Trim(felddaten)='' )")
MousePointer = 0

End Sub
Sub rlist1()
Dim r As ADODB.Recordset
Dim cmd$

List1.Clear

cmd$ = "select * from dictionary"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open cmd$, form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  List1.AddItem trm(r!id) & "=" & trm(r!translat)
  r.MoveNext
Wend

End Sub


Private Sub dbgopt_Click()
form1.dbg2file% = dbgopt.value
End Sub

Private Sub Form_Load()
Dim dbname$, cf$, c$, i%

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Me.Caption = "Agencyprof" & trm(App.Major) & "." & trm(App.Minor) & " Build #" & trm(App.Revision) & " APLib: " & bas_getAPLibVersion
vncv$ = LCase(form1.getuserid())
If vncv$ = "www" Or vncv$ = "administrator" Or vncv$ = "kurse" Or vncv$ = "kk" Then
  Command8.Enabled = True
  Command12.Enabled = True
End If
vncv$ = ""
c$ = form1.UseBrowser()
If c$ <> "" And Not nexist(c$) Then Command25.Enabled = True
For i% = 1 To 12: mnams_engl$(i%) = form1.dictionarylookupmonth(i%): Debug.Print mnams_engl$(i%):  Next i%
If Not nexist(form1.s0dir() & "\werkvz.mdb") Then rdwrkvz.Visible = True
Label1.Visible = False
Label2.Visible = False
Image1.Visible = False
Image2.Visible = False
If Not nexist("access_log") Then Command28.Visible = True
vncv$ = form1.getusersetting("vncviewer")
vncs$ = form1.getusersetting("vncserver")
If vncv$ <> "" Then
  Label1.Visible = True
  Image1.Visible = True
End If
If vncs$ <> "" Then
  Label2.Visible = True
  Image2.Visible = True
End If
Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
dbname$ = form1.getdbname()
If InStr(LCase(dbname$), ".mdb") > 0 Then
  Command19.Caption = "Export als SQL"
Else
  Command19.Caption = dbname$ & ".mdb erstellen"
End If
Command19.Enabled = True

List1.Clear
dbgopt.value = form1.dbg2file%
Show
Call rlist1
Call rlist2
Timer1.Interval = 10000
Timer1.Enabled = True
libloglock = True
If form1.libist > 1 Then
  liblog.Enabled = True
  liblog.value = imax(0, bas_APLibWriteLog(-1))
End If
libloglock = False
If InStr(LCase(App.EXEName), "apadmin") = 0 And InStr(LCase(App.EXEName), "projekt1") = 0 Then
  Command12.Visible = False
  Command8.Visible = False
  Command21.Visible = False
  Command19.Visible = False
  Command9.Visible = False
  Command6.Visible = False
  Command24.Visible = False
End If
c$ = form1.get1erg("select count(*) as wert from sysvars where owner like'sysvar_system_tlnk_%'")
Label4.Caption = c$ + " Topic Links"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub

Private Sub Image1_Click()
Dim X

X = Shell(vncv$, 1)

End Sub

Private Sub Image2_Click()
Dim X

X = Shell(vncs$, 1)

End Sub

Private Sub Label1_Click()
Call Image1_Click
End Sub

Private Sub Label2_Click()
Call Image2_Click
Unload frmBrowser
DoEvents
frmBrowser.StartingAddress = "http://www.agencyprof.de:5080/cgi-bin/sys/myip"
Load frmBrowser

End Sub

Private Sub liblog_Click()
If libloglock Then Exit Sub
Call bas_APLibWriteLog(liblog.value)

End Sub

Private Sub List1_DblClick()
Dim i%, brk As Boolean, ctch As Boolean
Dim n$, wert$, p%, c$, cob$

i% = List1.ListIndex
If i% >= 0 Then
wert$ = List1.List(i%)
p% = InStr(wert$, "=")
If p% > 0 Then
  n$ = trm(Mid$(wert$, p% + 1))
  wert$ = trm(Left$(wert$, p% - 1))
  n$ = InputBox("Übersetzung:" & vbCrLf & wert$, "Übersetzung", n$)
  c$ = "update dictionary set translat='" & n$ & "' where id='" & wert$ & "'"
  Call form1.sqlqry(c$)
  Call rlist1
Else
  If InStr(LCase(wert$), ".frm") > 0 Then
    o% = FreeFile
    Open wert$ For Input As #o%
    p% = FreeFile
    Open wert$ & ".add" For Output As #p%
    q% = FreeFile
    Open wert$ & ".tab" For Output As #q%
    brk = False: cob$ = ""
    While Not EOF(o%) And Not brk
      Line Input #o%, l$: l$ = trm(l$)
      If InStr(l$, "Attribute") = 1 Then
        brk = True
      Else
        If InStr(trm(l$), "Begin") = 1 And InStr(trm(l$), "Font") = 0 Then
          l$ = trm(strrepl(l$, " ", vbCrLf))
          cob$ = lastlineof(l$)
          Debug.Print cob$
        Else
          If cob$ <> "" Then
            zwort$ = word1(l$)
            Select Case zwort$:
              Case "End":
                cob$ = ""
              Case "Caption":
                ctch = True
              Case "ToolTipText":
                ctch = True
              Case "Text":
                ctch = True
              Case Else:
                Debug.Print l$
                ctch = False
            End Select
            If ctch Then
              l1$ = strrepl(l$, " ", vbCrLf)
              r% = InStr(l$, Chr$(34))
              cb$ = ""
              If r% > 0 Then
                cb$ = Mid$(l$, r%)
              End If
              outl$ = cob$ & "." & zwort$ & "=form1.inmylanguage(" & cb$ & ")"
              Debug.Print outl$
              Print #p%, outl$
              Print #q%, strrepl(cb$, """", "") & "|"
            End If
          End If
        End If
      End If
    Wend
    Close #o%
    Close #p%
    Close #q%
    X = Shell("notepad.exe " & wert$ & ".tab", 1)
    X = Shell("notepad.exe " & wert$ & ".add", 1)
  End If
  Call rlist1
End If
End If


End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%
Dim n$, wert$, p%, c$

If KeyCode = 8 Or KeyCode = 46 Then
  i% = List1.ListIndex
  If i% >= 0 Then
    wert$ = List1.List(i%)
    p% = InStr(wert$, "=")
    If p% > 0 Then
      wert$ = trm(Left$(wert$, p% - 1))
      c$ = "delete from dictionary where id='" & wert$ & "'"
      Call form1.sqlqry(c$)
      List1.RemoveItem i%
    End If
  End If
End If


End Sub
Sub rlist2()
Dim tr

List2.Clear

tr = Dir(form1.s0dir() & "\*.run")
While tr <> ""
  List2.AddItem basename(trm(tr), ".run")
  tr = Dir
Wend

End Sub




Private Sub rdwrkvz_Click()
Dim s As Recordset, t As ADODB.Recordset, r As Recordset
Dim acc As Database, tid$, d0$
Dim datwert As String
Dim st$, tt$, V$, b$, d$, p%

MousePointer = 11: DoEvents
If nexist(form1.s0dir() & "\werkvz.mdb") Then Exit Sub
List3.Visible = True
cmd$ = "delete from k_loc where name<>'Pause' and vornamen<>'Pause' name<>'oder' and vornamen<>'oder'": List3.AddItem cmd$: Call form1.sqlqry(cmd$)
cmd$ = "delete from w_loc where name<>'Pause' and name<>'oder'": List3.AddItem cmd$: Call form1.sqlqry(cmd$)
cmd$ = "delete from sbz_loc": List3.AddItem cmd$: Call form1.sqlqry(cmd$)
Set acc = wrkJet.OpenDatabase(form1.s0dir() & "\werkvz.mdb", False, True)
cmd$ = "select * from komponis order by kompnr desc"
Set s = acc.OpenRecordset(cmd$, dbOpenDynaset, dbOpenDynaset)
List3.Clear
List3.AddItem "Komponisten werden gelesen ..."
DoEvents
While Not s.EOF
  st$ = strrepl(trm(s!name) & "-" & trm(s!vorname) & trm(s!daten) & trm(s!altnam) & datum2sql(s!datum), "'", "´")
  If st$ <> "-" Then
    tt$ = ""
    tid = s!kompnr
    Set t = New ADODB.Recordset
    t.CursorLocation = adUseServer
    t.Open "SELECT * FROM k_loc where id='" & s!kompnr & "'", form1.adoc, adOpenDynamic, adLockReadOnly
    If Not t.EOF Then
      tt$ = trm(t!name) & "-" & trm(t!vornamen) & trm(t!daten) & trm(t!Alternativschreibweisen) & trm(t!stand)
      tid = t!id
    End If
    If tt$ <> st$ Then
      List3.AddItem s!kompnr & " " & s!name & ", " & s!vorname
      List3.ListIndex = List3.ListCount - 1
      DoEvents
      If tt$ = "" Then
        cmd$ = "insert into k_loc (id) values('" & s!kompnr & "')": Call form1.sqlqry(cmd$)
      End If
      cmd$ = "update k_loc set name='" & strrepl(trm(s!name), "'", "´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update k_loc set vornamen='" & strrepl(trm(s!vorname), "'", "´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update k_loc set daten='" & trm(s!daten) & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      d$ = trm(s!daten)
      p% = InStr(d$, "-")
      If p% > 0 Then
        V$ = trm(Left(d$, p% - 1))
        b$ = trm(Mid(d$, p% + 1))
        If V$ <> "" Then Call form1.sqlqry("update k_loc set von='" & V$ & "' where id='" & tid & "'")
        If b$ <> "" Then Call form1.sqlqry("update k_loc set bis='" & b$ & "' where id='" & tid & "'")
      End If
      cmd$ = "update k_loc set alternativschreibweisen='" & strrepl(trm(s!altnam), "'", "´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update k_loc set stand='" & datum2sql(s!datum) & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
    End If
    t.Close
  End If
  s.MoveNext
Wend
List3.AddItem "Werke werden gelesen ..."
List3.ListIndex = List3.ListCount - 1
DoEvents
cmd$ = "select * from kompos order by opcode desc;"
Set s = acc.OpenRecordset(cmd$, dbOpenDynaset, dbOpenDynaset)
While Not s.EOF
  If IsNull(s!opcode) Or InStr(LCase(s!opname1), "?") = 1 Then
    Debug.Print s!opname1; " "; s!opname2; " nicht importiert"
  Else
  dau$ = trm(s!Dauer): If dau$ = "" Then dau$ = " "
  st$ = strrepl(strrepl(dau$ & _
        trm(s!opname1) & "-" & _
        trm(s!opname2) & _
        trm(s!Tonart) & _
        trm(s!nummer) & _
        trm(s!opusbez) & _
        trm(s!opusnr) & _
        trm(s!opjahr) & _
        trm(s!opjanf) & _
        trm(s!opjend) & _
        trm(s!satz1) & _
        trm(s!satz2) & _
        trm(s!satz3) & _
        trm(s!satz4) & _
        trm(s!satz5) & _
        trm(s!satz6) & _
        trm(s!satz7) & _
        trm(s!satz8) & _
        trm(s!satz9) & _
        trm(s!satz10) & _
        trm(s!satz11) & _
        trm(s!satz12) & _
        trm(s!satz13) & _
        trm(s!satz14) & datum2sql(s!datum) & _
        trm(s!bem), "'", "´"), """", "´´")
  If st$ <> "-" Then
    tt$ = ""
    tid = s!opcode
    Set t = New ADODB.Recordset
    t.CursorLocation = adUseServer
    t.Open "SELECT * FROM w_loc where id='" & tid & "'", form1.adoc, adOpenDynamic, adLockReadOnly
    If Not t.EOF Then
      d1$ = trm(t!Dauer): If d1$ = "" Then d1$ = " "
      tt$ = d1$ & _
        trm(t!Opusname1) & "-" & _
        trm(t!Opusname2) & _
        trm(t!Tonart) & _
        trm(t!nummer) & _
        trm(t!Opusbezeichnung) & _
        trm(t!OpusNummer) & _
        trm(t!Opusjahr) & _
        trm(t!Opusjahr_von) & _
        trm(t!Opusjahr_bis) & _
        trm(t!s1) & _
        trm(t!s2) & _
        trm(t!s3) & _
        trm(t!s4) & _
        trm(t!s5) & _
        trm(t!s6) & _
        trm(t!s7) & _
        trm(t!s8) & _
        trm(t!s9) & _
        trm(t!s10) & _
        trm(t!s11) & _
        trm(t!s12) & _
        trm(t!s13) & _
        trm(t!s14) & trm(t!stand) & _
        trm(t!Bemerkung)
      tid = t!id
    End If
    If tt$ <> st$ Then
Debug.Print st$ & vbCrLf & tt$
      If tt$ = "" Then
        cmd$ = "insert into w_loc (id,KomponistenNummer) values('" & tid & "','" & Left(tid, 4) & "')": Call form1.sqlqry(cmd$)
      End If
      List3.AddItem s!opcode & " " & s!opname1 & " (" & form1.getkompnamebywerkid(tid) & ")"
      List3.ListIndex = List3.ListCount - 1
      DoEvents
      cmd$ = "update w_loc set opusname1='" & strrepl(strrepl(trm(s!opname1), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set opusname2='" & strrepl(strrepl(trm(s!opname2), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set dauer='" & dau$ & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set tonart='" & strrepl(strrepl(trm(s!Tonart), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set nummer='" & strrepl(strrepl(trm(s!nummer), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set opusbezeichnung='" & strrepl(strrepl(trm(s!opusbez), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      If Not IsNull(s!opusnr) Then
        cmd$ = "update w_loc set opusnummer='" & strrepl(strrepl(s!opusnr, "'", "´"), """", "´´") & "' where id='" & tid & "'"
        Call form1.sqlqry(cmd$)
      End If
      cmd$ = "update w_loc set opusjahr='" & strrepl(strrepl(trm(s!opjahr), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set opusjahr_von='" & strrepl(strrepl(trm(s!opjanf), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set opusjahr_bis='" & strrepl(strrepl(trm(s!opjend), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s1='" & strrepl(strrepl(trm(s!satz1), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s2='" & strrepl(strrepl(trm(s!satz2), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s3='" & strrepl(strrepl(trm(s!satz3), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s4='" & strrepl(strrepl(trm(s!satz4), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s5='" & strrepl(strrepl(trm(s!satz5), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s6='" & strrepl(strrepl(trm(s!satz6), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s7='" & strrepl(strrepl(trm(s!satz7), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s8='" & strrepl(strrepl(trm(s!satz8), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s9='" & strrepl(strrepl(trm(s!satz9), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s10='" & strrepl(strrepl(trm(s!satz10), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s11='" & strrepl(strrepl(trm(s!satz11), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s12='" & strrepl(strrepl(trm(s!satz12), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s13='" & strrepl(strrepl(trm(s!satz13), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set s14='" & strrepl(strrepl(trm(s!satz14), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set stand='" & datum2sql(s!datum) & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
      cmd$ = "update w_loc set bemerkung='" & strrepl(strrepl(trm(s!bem), "'", "´"), """", "´´") & "' where id='" & tid & "'": Call form1.sqlqry(cmd$)
    
  knr$ = strrepl(trm(s!opname1), """", "´´")
  wert$ = trm(" " & s!nummer)
  If wert$ <> "" Then
    knr$ = knr$ & " Nr. " & wert$
  End If
  wert$ = trm(" " & s!Tonart): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = trm(" " & s!opusbez): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = trm(" " & s!opusnr): If wert$ <> "" Then knr$ = knr$ & " " & s!opusnr
  wert$ = strrepl(trm(" " & s!opname2), """", "´´"): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  Call form1.sqlqry("update w_loc set name='" & strrepl(strrepl(knr$, """", "´´"), "'", "´") & "' where id='" & tid & "'")
    
      cmd$ = "delete from sbz_loc where wid='" & tid & "'": Call form1.sqlqry(cmd$)
      cnt = 0
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz1), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz2), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz3), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz4), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz5), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz6), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz7), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz8), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz9), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz10), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz11), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz12), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz13), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
      cnt = cnt + 1: knr$ = strrepl(strrepl(trm(s!satz14), "'", "´"), """", "´´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & tid & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
    End If
    t.Close
  End If
  End If
  s.MoveNext
Wend
d0$ = form1.s0dir()
fn$ = d0$ & "\xp_k_loc.txt": List3.AddItem "Komponisten werden exportiert: " & fn$: List3.ListIndex = List3.ListCount - 1: DoEvents
Call form1.pg_xp("k_loc", fn$)
fn$ = d0$ & "\xp_w_loc.txt": List3.AddItem "Werke werden exportiert: " & fn$: List3.ListIndex = List3.ListCount - 1: DoEvents
Call form1.pg_xp("w_loc", fn$)
fn$ = d0$ & "\xp_sbz_loc.txt": List3.AddItem "Satzbezeichnungen werden exportiert: " & fn$: List3.ListIndex = List3.ListCount - 1: DoEvents
Call form1.pg_xp("sbz_loc", fn$)
X = Shell("explorer.exe " & d0$, vbNormalFocus)
List3.Visible = False
MousePointer = 0

End Sub

Private Sub Timer1_Timer()
Call rlist2
End Sub

Sub mdb2sql()
List1.Clear
dbname$ = form1.getdbname()
If InStr(LCase(dbname$), ".mdb") = 0 Then Exit Sub
fn$ = form1.mydatadir() & "\" & dbname$ & "-sql.txt"
If exist(fn$) <> 0 Then
  ask% = MsgBox(fn$ & " existiert - überschreiben?", vbYesNo + vbCritical + vbDefaultButton1, "SQL-Datenbankexport")
  If ask% <> vbYes Then Exit Sub
  Kill fn$
End If
o% = FreeFile
Open fn$ For Output As #o%
Debug.Print form1.sqla.name; " mit"; form1.sqla.TableDefs.Count - 1; " Tabellen"
For i = 0 To form1.sqla.TableDefs.Count - 1
  flst$ = ""
  If Left$(LCase(form1.sqla.TableDefs(i).name), 4) <> "msys" Then
    Debug.Print form1.sqla.TableDefs(i).name
    List1.AddItem form1.sqla.TableDefs(i).name
    List1.ListIndex = List1.ListCount - 1
    DoEvents
    For j = 0 To 99: f_ldn$(j) = "": Next j
    c$ = ""
    adl$ = ""
    For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
      ftyp = form1.sqla.TableDefs(i).Fields(j).Type
      fnam = form1.sqla.TableDefs(i).Fields(j).name
      isid = False
      If InStr(fnam, "Konzertklei") = 0 Then
        If LCase(fnam) = "id" Then
          adl$ = adl$ + vbCrLf + "alter table " + form1.sqla.TableDefs(i).name + " ADD PRIMARY KEY (" + fnam + ");"
          isid = True
        Else
          If InStr(LCase(fnam), "id") > 0 Then
            adl$ = adl$ + vbCrLf + "alter table " + form1.sqla.TableDefs(i).name + " ADD INDEX " + fnam + " (" + fnam + ");"
          End If
        End If
      End If
      fsiz = form1.sqla.TableDefs(i).Fields(j).Size
      If flst$ <> "" Then flst$ = flst$ & ","
      flst$ = flst$ & fnam
      f_ldn$(j) = fnam
      Debug.Print ftyp; "; "; fnam; "; s="; fsiz
      If ftyp = 0 Then ftyp = 3
      t1yp$(j) = trm(ftyp)
      sqlt$ = ""
      Select Case ftyp
        Case 20: sqlt$ = "bigint"
        Case 10: sqlt$ = "varchar"
        Case 12: sqlt$ = "longtext"
        Case 8: sqlt$ = "timestamp": fsiz = 14
        Case 3: sqlt$ = "smallint": fsiz = 14
        Case 7: sqlt$ = "double": fsiz = 0
        Case 5: sqlt$ = "int"
        Case 4: sqlt$ = "int"
        Case 2: sqlt$ = "int"
        Case 1: sqlt$ = "int"
        Case Else
      End Select
      If sqlt$ = "" Then
        MsgBox "unbekannter typ: " & ftyp & "; " & fnam & "; s=" & fsiz
        End
      Else
        If c$ <> "" Then c$ = c$ & ", "
        c$ = c$ & fnam & " " & sqlt$
        If fsiz <> 0 Then
          c$ = c$ & "(" & trm(fsiz) & ")"
          If isid Then c$ = c$ + " NOT NULL"
        End If
      End If
    Next j
    c1$ = "create table " & form1.sqla.TableDefs(i).name & " ( " & c$ & " ) TYPE = MYISAM;"
    Debug.Print c1$
    Print #o%, c1$
    If adl$ <> "" Then
      Debug.Print adl$
      Print #o%, adl$
    End If
    c$ = "select * from " & form1.sqla.TableDefs(i).name
    Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    While Not r.EOF
      werte$ = ""
      For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
        'If trm("" & r.Fields(j).value) <> "" Then
        'ff$ = "'": If t1yp$(j) = "3" Or t1yp$(j) = "4" Or t1yp$(j) = "7" Then ff$ = ""
        ff$ = "'"
        If ff$ = "" Then
          werte$ = werte$ & strrepl(trm("" & r.Fields(j).value), ",", ".") & ","
        Else
          werte$ = werte$ & ff$ & r.Fields(j).value & ff$ & ","
        End If
        'End If
      Next j
      werte$ = trm(Left$(werte$, Len(werte$) - 1))
      While Right$(werte$, 1) = ",": werte$ = Left$(werte$, Len(werte$) - 1): Wend
      werte$ = strrepl(werte$, "\", "||--||")
      werte$ = strrepl(werte$, "||--||", "\\")
      c$ = "insert into " & form1.sqla.TableDefs(i).name & " (" & flst$ & ") values(" & werte$ & ");"
'      Debug.Print c$
      Print #o%, c$
      r.MoveNext
    Wend
  End If
Next i
Close #o%
X = Shell("notepad.exe " & fn$, 1)
Call rlist1

End Sub
