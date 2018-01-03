VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form import 
   Caption         =   "Data-Import"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton kompos 
      Caption         =   "kompos.dbf"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton komponis 
      Caption         =   "komponis.dbf"
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   29
      Text            =   "200"
      ToolTipText     =   "# of records to import (-1 = all)"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "SQL"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "alle bis vdx"
      Height          =   255
      Left            =   7800
      TabIndex        =   27
      Top             =   5400
      Width           =   2175
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ALTID löschen"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Programme(KDS2)"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Werke (KDS2)"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Komponisten (KDS2)"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7800
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4800
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   4680
      Width           =   5175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LOS"
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ListBox List7 
      Height          =   1425
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "la&den"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "speichern"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.ListBox List6 
      Height          =   1425
      Left            =   6000
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&clear"
      Height          =   255
      Left            =   9360
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&ok"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox List5 
      Height          =   1425
      Left            =   4800
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Height          =   2400
      Left            =   7440
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox List3 
      Height          =   2400
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-->"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&lies"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "import.mdb"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label3"
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "MS-Access"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wrkJet As Workspace
'Dim sqla As Database
Sub rlist7()

tr$ = Dir("*.sqi")
List7.Clear
While tr$ <> ""
  List7.AddItem tr$
  tr$ = Dir
Wend
For i% = 0 To List7.ListCount - 1
  If List7.List(i%) = Text2.Text Then
    List7.ListIndex = i%
    i% = List7.ListCount
  End If
Next i%
import.Caption = List7.ListCount + " Importdefinitionen vorhanden"

End Sub
Private Sub Command1_Click()
Dim acc As Database

If exist(Text1.Text) = 0 Then Exit Sub
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)
List2.Clear
For i% = 0 To acc.TableDefs.Count - 1
  If Left$(LCase$(acc.TableDefs(i%).name), 4) <> "msys" And InStr(acc.TableDefs(i%).name, "(") = 0 Then
    Set r = acc.OpenRecordset( _
      "SELECT count(*) as cnt FROM " + acc.TableDefs(i%).name, dbOpenDynaset, dbReadOnly)

    ad$ = acc.TableDefs(i%).name + " " + r!cnt + " recs"
    List2.AddItem ad$
  End If
Next i%
acc.Close

End Sub

Private Sub Command10_Click()
Dim s As Recordset, t As Recordset
Dim acc As Database, c$, cmd$, sid As String, stx As String, nkid$

If exist(Text1.Text) = 0 Then Exit Sub
MousePointer = 11: DoEvents
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)

cmd$ = "select * from komponisten"
Set s = acc.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not s.EOF
  Command2.Caption = s!id
  DoEvents
  sid = strrepl(s!id, "'", "´")
  stx = ""
  If Not IsNull(s!Text) Then
    stx = strrepl(s!Text, "'", "´")
  End If
  nn$ = word1(sid): If Right$(nn$, 1) = "," Then nn$ = Left$(nn$, Len(nn$) - 1)
  vn$ = ""
  If stx <> "" Then
    If InStr(stx, nn$) > 1 Then vn$ = trm(Left(stx, InStr(stx, nn$) - 1))
  End If
  c$ = "select id,name,vornamen,Alternativschreibweisen as altn from k_loc where name='" + nn$ + "' and vornamen='" + vn$ + "'"
  Set t = form1.pub_sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If Not t.EOF Then         'schon da
    c$ = "update k_loc set Alternativschreibweisen='ALTID:" + sid + vbCrLf + t!altn + "'  where id='" + t!id + "'"
    Call form1.sqlqry(c$)
  Else              'neu
    dat$ = ""
    If InStr(stx, vbCrLf) > 0 Then dat$ = Mid$(stx, InStr(stx, vbCrLf) + 2)
    von$ = dat$: bis$ = dat$
    If InStr(von$, "-") Then
      von$ = trm(Left(von$, InStr(von$, "-") - 1))
      While Left$(von$, 1) = "("
        von$ = Mid$(von$, 2)
      Wend
    End If
    If InStr(bis$, "-") Then
      bis$ = trm(Mid(bis$, InStr(bis$, "-") + 1))
      While Right$(bis$, 1) = ")"
        bis$ = Left(bis$, Len(bis$) - 1)
      Wend
    End If
    nkid$ = s!id
    nknr$ = form1.newid("k_loc", "id", 4)
    c$ = "insert into k_loc (id,name,vornamen,daten,von,bis,kompnr,Alternativschreibweisen) values('" + _
        nknr$ + "','" + nn$ + "','" + vn$ + "','" + dat$ + "','" + von$ + "','" + bis$ + "','" + nknr$ + "','ALTID:" + sid + "')"
    Call form1.sqlqry(c$)
  End If
  s.MoveNext
Wend
MousePointer = 0
Command2.Caption = "&Schliessen"

End Sub

Private Sub Command11_Click()
Dim s As Recordset, r As Recordset
Dim acc As Database, sid As String

If exist(Text1.Text) = 0 Then Exit Sub
MousePointer = 11: DoEvents
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)

cmd$ = "select * from werke"
Set s = acc.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not s.EOF
  sid = strrepl(s!id, "'", "´")
  Command2.Caption = sid: DoEvents
  If InStr(sid, ":") > 0 Then
    k$ = trm(Left$(sid, InStr(sid, ":") - 1))
    kn$ = k$
    n$ = strrepl(trm(Mid$(sid, InStr(sid, ":") + 1)), "'", "")
    'p% = InStr(k$, ",")
    'k1$ = Left$(k$, p% - 1)
    'k2$ = trm(Mid$(k$, p% + 1))
    'k2$ = Left(k2$, Len(k2$) - 1)
    c$ = "select * from k_loc where id='" + k$ + "'"
    kid$ = ""
    Set r = form1.pub_sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    If r.EOF Then
      c$ = "select * from k_loc where (lcase(Alternativschreibweisen) like 'ALTID:" + k$ + "*')"
      Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
      If Not r.EOF Then
        p% = InStr(r!Alternativschreibweisen, "ALTID:")
        If p% > 0 Then
          kid$ = Mid(r!Alternativschreibweisen, p% + 6)
          If InStr(kid$, vbCrLf) > 0 Then kid$ = Left(kid$, InStr(kid$, vbCrLf) - 1)
          k$ = trm(r!kompnr)
        End If
      End If
      If kid$ = "" Then
        kid$ = form1.newid("k_loc", "id", 4)
        k$ = kid$
        c$ = "insert into k_loc (id,name,kompnr) values('" + kid$ + "','" + kn$ + "','" + k$ + "')"
        Call form1.sqlqry(c$)
      End If
    Else
      kid$ = r!id
      k$ = trm(r!kompnr)
    End If
    If Len(k$) > 2 Then k$ = Left(k$, 4)
    id$ = ""
    c$ = "select * from w_loc where (name='" + n$ + "')"
    Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    While Not r.EOF And id$ = ""
      If r!KomponistenNummer = k$ Then
        id$ = r!id
        r.MoveLast
      End If
      r.MoveNext
    Wend
    If id$ = "" Then
      id$ = form1.newid("w_loc", "id", 8)
      d$ = "0": If Not IsNull(s!länge) Then d$ = s!länge
      c$ = "insert into w_loc (id,name,komponistennummer,dauer) values('" + _
                          id$ + "','" + n$ + "','" + k$ + "','" + d$ + "')"
        Call form1.sqlqry(c$)
        scnt% = 0
      l$ = trm(s!Text)
      While Len(l$) > 0
        p% = InStr(l$, vbCrLf)
        If p% = 0 Then
          l1$ = l$
          l$ = ""
        Else
          l1$ = trm(Left$(l$, p% - 1))
          l$ = trm(Mid$(l$, p% + 2))
        End If
        i1$ = form1.newid("sbz_loc", "id", 10)
        scnt% = scnt% + 10
        c$ = "insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" + _
                        i1$ + "','" + id$ + "','" + l1$ + "'," + trm(scnt%) + ")"
        Call form1.sqlqry(c$)
      Wend
    End If
  End If
  s.MoveNext
Wend
MousePointer = 0
Command2.Caption = "&Schliessen"
End Sub

Private Sub Command12_Click()
Dim r As Recordset
Dim acc As Database

If exist(Text1.Text) = 0 Then Exit Sub
MousePointer = 11: DoEvents
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)

cmd$ = "select * from programm"
Set r = acc.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  Command2.Caption = r!id: DoEvents
  id$ = r!id
  c$ = "insert into programm (programmid) values('" + id$ + "')"
  Call form1.sqlqry(c$)
  pcnt% = 0
  pcnt% = pcnt% + 10: w$ = trm(r!werkid1): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid2): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid3): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid4): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid5): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid6): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid7): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid8): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid9): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid10): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid11): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid12): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid13): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid14): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid15): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid16): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid17): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid18): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid19): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid20): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid21): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid22): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid23): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  pcnt% = pcnt% + 10: w$ = trm(r!werkid24): If w$ <> "" Then Call werkinprg(w$, id$, pcnt%)
  r.MoveNext
Wend
MousePointer = 0
Command2.Caption = "&Schliessen"
End Sub

Sub werkinprg(werk$, prgid$, pcnt%)
Dim s As Recordset, r As Recordset

p% = InStr(werk$, ":")
If p% = 0 Then Exit Sub
w$ = strrepl(trm(Mid$(werk$, p% + 1)), "'", "´")
k$ = strrepl(trm(Left$(werk$, p% - 1)), "'", "´")
knr$ = "": wid$ = ""
c$ = "select id from w_loc where name='" + w$ + "' and KomponistenNummer='" + k$ + "'"
Set s = form1.pub_sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not s.EOF Then
  w$ = s!id
Else
  c$ = "select * from k_loc where name='" + k$ + "'"
  Set r = form1.pub_sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If r.EOF Then
    c$ = "select * from k_loc where (lcase(Alternativschreibweisen) like 'ALTID:" + k$ + "*')"
    Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    If Not r.EOF Then
      p% = InStr(r!Alternativschreibweisen, "ALTID:")
      If p% > 0 Then
        kid$ = Mid(r!Alternativschreibweisen, p% + 6)
        If InStr(kid$, vbCrLf) > 0 Then kid$ = Left(kid$, InStr(kid$, vbCrLf) - 1)
        id$ = r!id
        knr$ = trm(r!kompnr)
      End If
    Else
      id$ = form1.newid("k_loc", "id", 8)
      knr$ = mkkey(4)
      c$ = "insert into k_loc (id,name,kompnr) values('" + id$ + "','" + k$ + "','" + knr$ + "')"
      Call form1.sqlqry(c$)
    End If
  Else
    id$ = r!id
    knr$ = trm(r!kompnr)
  End If
  If knr$ <> "" Then
    c$ = "select id from w_loc where name='" + w$ + "' and KomponistenNummer='" + knr$ + "'"
    Set s = form1.pub_sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    If Not s.EOF Then
      wid$ = s!id
      w$ = s!id
    End If
  End If
  If wid$ = "" Then
    wid$ = form1.newid("w_loc", "id", 8)
    c$ = "insert into w_loc (id,name,komponistennummer) values('" + _
          wid$ + "','" + w$ + "','" + knr$ + "')"
    Call form1.sqlqry(c$)
  End If
  w$ = wid$
End If
cmd$ = "insert into programmliste (id,programmid,werkid,position) values('" + _
    form1.newid("programmliste", "id", 10) + "','" + prgid$ + "','" + w$ + "'," + trm(pcnt%) + ")"
Call form1.sqlqry(cmd$)

End Sub

Private Sub Command13_Click()
Dim s As Recordset, t As Recordset
Dim acc As Database, c$, cmd$

MousePointer = 11
cmd$ = "select * from k_loc where (lcase(Alternativschreibweisen) like 'ALTID:*')"
Set s = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not s.EOF
  p% = InStr(s!Alternativschreibweisen, "ALTID:")
  If p% > 0 Then
    p% = InStr(s!Alternativschreibweisen, vbCrLf)
    If p% > 0 Then
      aw$ = trm(Mid(s!Alternativschreibweisen, p% + 3))
    Else
      aw$ = ""
    End If
    c$ = "update k_loc set Alternativschreibweisen='" + aw$ + "' where id='" + s!id + "'"
    Command2.Caption = c$
    DoEvents
    Call form1.sqlqry(c$)
  End If
  s.MoveNext
Wend
MousePointer = 0
Command2.Caption = "&Schliessen"

End Sub

Private Sub Command14_Click()

While InStr(LCase(List7.List(List7.ListIndex)), "vdx") = 0
  DoEvents
  Call Command8_Click
Wend
End Sub

Private Sub Command15_Click()
Dim i%, u%, tn$, fn$, j%, ftyp, ft$, fs$
Dim acc As Database, r As Recordset, c$, cnt As Long, nrcs$

u% = FreeFile
List3.Clear
List4.Clear
List5.Clear
List6.Clear
If exist(Text1.Text) = 0 Then Exit Sub
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)

u% = FreeFile
nrcs$ = trm(Text4.Text)
If nrcs$ = "-1" Then
  nrcs$ = ""
Else
  nrcs$ = "top " + nrcs$ + " "
End If
fn$ = form1.mydatadir() + "\import_struct.sql"
Open fn$ For Output As #u%
Print #u%, "drop database import;"
Print #u%, "create database import;"
Print #u%, "use import;"
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  tn$ = word1(List2.List(i%))
  Print #u%, "CREATE TABLE " + tn$ + " (imp_id varchar (40) default '0' not null, ";
  cmd$ = ""
  For j% = 0 To acc.TableDefs(tn$).Fields.Count - 1
    ftyp = acc.TableDefs(tn$).Fields(j).Type
    fs$ = ""
    Select Case ftyp
      Case 3: ft$ = "int"
      Case 4: ft$ = "bigint"
      Case 5: ft$ = "char(20)"          ' currency not implemented
      Case 7: ft$ = "double"
      Case 8: ft$ = "datetime"
      Case 10: ft$ = "char": fs$ = " (" + trm(acc.TableDefs(tn$).Fields(j%).Size) + ")"
      Case 12: ft$ = "longtext"
      Case Else: ft$ = ""
    End Select
    cmd$ = cmd$ & acc.TableDefs(tn$).Fields(j%).name & " " & ft$ & fs$ & ", "
  Next j%
  Print #u%, cmd$;
  Print #u%, "primary key(imp_id));"
Next i%
Close #u%
X = Shell("notepad.exe " & fn$, 1)
fn$ = form1.mydatadir() & "\import_data.sql"
Open fn$ For Output As #u%
Print #u%, "use import;"
For i% = 0 To List2.ListCount - 1
  List2.ListIndex = i%
  DoEvents
  tn$ = word1(List2.List(i%))
  cmd$ = "insert into " & tn$ & " (imp_id,"
  For j% = 0 To acc.TableDefs(tn$).Fields.Count - 1
    cmd$ = cmd$ & acc.TableDefs(tn$).Fields(j%).name & ","
  Next j%
  cmd$ = Left$(cmd$, Len(cmd$) - 1) & ") values('"
  cnt = 0
  c$ = "select " & nrcs$ & "* from " & tn$
  Set r = acc.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  While Not r.EOF
    cnt = cnt + 1
    If (cnt Mod 100) = 0 Then
      Command8.Caption = cnt
      DoEvents
    End If
    c$ = ""
    For j% = 0 To acc.TableDefs(tn$).Fields.Count - 1
      c$ = c$ & strrepl("" & r.Fields(j%).value, "'", "´") & "','"
    Next j%
    c$ = Left$(c$, Len(c$) - 2) & ");"
    Print #u%, cmd$ & GUID() & "','" & c$
    r.MoveNext
  Wend
Next i%
Close #u%

acc.Close

End Sub

Private Sub Command2_Click()
Width = 4920
Unload import
End Sub

Private Sub Command3_Click()
Dim acc As Database

List3.Clear
List4.Clear
List5.Clear
List6.Clear
If exist(Text1.Text) = 0 Then Exit Sub

Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)
er% = List2.ListIndex
ich% = List1.ListIndex
If ich% >= 0 And er% >= 0 Then
  Width = 10065
  t_e$ = cut_d1(List2.List(er%) + " ", " ")
  t_ic$ = cut_d1(List1.List(ich%) + " ", " ")
  Label3.Caption = t_e$
  Label4.Caption = t_ic$
  List3.Clear
  List4.Clear
  For i% = 0 To acc.TableDefs(t_e$).Fields.Count - 1
    List3.AddItem acc.TableDefs(t_e$).Fields(i%).name
  Next i%
  For i% = 0 To form1.sqla.TableDefs(t_ic$).Fields.Count - 1
    List4.AddItem form1.sqla.TableDefs(t_ic$).Fields(i%).name
  Next i%
End If

acc.Close
Text2.Text = Label3.Caption + "-" + Label4.Caption + ".sqi"

End Sub

Private Sub Command4_Click()

er% = List2.ListIndex
ich% = List1.ListIndex
c_er% = List3.ListIndex
c_ich% = List4.ListIndex
If ich% >= 0 And er% >= 0 Then
  If c_ich% >= 0 And c_er% >= 0 Then
    t_e$ = cut_d1(List2.List(er%), " ")
    t_ic$ = cut_d1(List1.List(ich%), " ")
    c_e$ = List3.List(c_er%)
    c_ic$ = List4.List(c_ich%)
    List5.AddItem c_e$
    List6.AddItem c_ic$
    List5.ListIndex = List5.ListCount - 1
    List6.ListIndex = List6.ListCount - 1
  Else
    If c_ich% >= 0 Then
      c_ic$ = List4.List(c_ich%)
      List5.AddItem "NULL"
      List6.AddItem c_ic$
      List5.ListIndex = List5.ListCount - 1
      List6.ListIndex = List6.ListCount - 1
    End If
  End If
End If
End Sub

Private Sub Command5_Click()
List5.Clear
List6.Clear
End Sub

Private Sub Command6_Click()

o% = FreeFile
fn$ = Label3.Caption + "-" + Label4.Caption + ".sqi"
If Text2.Text <> "" Then fn$ = Text2.Text
Open fn$ For Output As #o%
Print #o%, Label3.Caption
Print #o%, Label4.Caption
Print #o%, List5.ListCount
For i% = 0 To List5.ListCount - 1
  Print #o%, List5.List(i%)
  Print #o%, List6.List(i%)
Next i%
Close #o%
Call rlist7

End Sub

Private Sub Command7_Click()
Call List7_DblClick

End Sub

Private Sub Command8_Click()
Dim s As Recordset, s1 As Recordset
Dim acc As Database
Dim datwert As String
Dim cnt


If exist(Text1.Text) = 0 Then Exit Sub
nnn% = Val(Text3.Text)
While nnn% > 0
 Text3.Text = nnn%
Set acc = wrkJet.OpenDatabase(Text1.Text, False, True)
cmd$ = "select "
cmd1$ = List5.List(0)
If InStr(cmd1$, "__") > 0 Then cmd1$ = Mid$(cmd1$, InStr(cmd1$, "__") + 2)
If InStr(cmd1$, "+") > 0 Then cmd1$ = Left$(cmd1$, InStr(cmd1$, "+") - 1)
cmd$ = cmd$ + cmd1$
For i% = 1 To List5.ListCount - 1
  If Left$(List5.List(i%), 1) <> "=" Then
    If Left$(List5.List(i%), 5) <> "date:" Then
      l$ = List5.List(i%)
      If InStr(l$, "+") > 0 Then
        l$ = Left$(l$, InStr(l$, "+") - 1)
      End If
      cmd$ = cmd$ + ", " + l$
    Else
      cmd$ = cmd$ + ", " + Mid$(List5.List(i%), InStr(List5.List(i%), ":") + 1)
    End If
  End If
Next i%
cmd$ = cmd$ + " from " + Label3.Caption
Debug.Print cmd$
Set s = acc.OpenRecordset(cmd$, dbOpenDynaset, dbOpenDynaset)
cnt = 0
tnum% = -1
fnum% = -1
hnum% = -1
For i% = 1 To List6.ListCount - 1
  'cmd$ = cmd$ + ", " + List6.List(i%)
  Select Case LCase(List6.List(i%))
    Case "tel": tnum% = i%
    Case "fax": fnum% = i%
    Case "handy": hnum% = i%
    Case Default:
  End Select
Next i%

While Not s.EOF
  cnt = cnt + 1
  Command8.Caption = "" & cnt
  DoEvents
  cmd$ = "insert into " + Label4.Caption + " ("
  cmd$ = cmd$ + List6.List(0)
  voni% = 1: If List6.List(0) = List6.List(1) Then voni% = 2
  For i% = voni% To List6.ListCount - 1
    If InStr(List6.List(i%), ":") = 0 Then
      cmd$ = cmd$ + ", " + List6.List(i%)
      While List6.List(i%) = List6.List(i% + 1)
        i% = i% + 1
      Wend
    End If
  Next i%
  cid$ = ""
  If IsNull(s.Fields(0).value) Then
    c0id$ = form1.newid(Label4.Caption, List6.List(0), 10)
    cmd$ = cmd$ + ") values('" + c0id$
  Else
    If LCase(Label4.Caption) = "kontakt" Then
      datwert = form1.newid("kontakt", "id", 8)
      cmd$ = cmd$ + ") values('" & datwert
      cid$ = datwert
    Else
      datwert = strrepl("" & s.Fields(0).value, "'", " ")
      datwert = strrepl(datwert, """", "´´")
      cmd$ = cmd$ + ") values('" & datwert
      cid$ = datwert
    End If
  End If
  If InStr(List5.List(0), "+") > 0 Then
    cmd$ = cmd$ + Mid$(List5.List(0), InStr(List5.List(0), "+") + 1)
  End If
  skp% = 0
  For i% = 1 To List6.ListCount - 1
    If Left$(List5.List(i%), 1) <> "=" Then
      If Left$(List5.List(i%), 5) <> "date:" Then
        If InStr(List6.List(i%), ":") <> 0 Then
          twrt$ = trm(strrepl("" & s.Fields(i% - skp%).value, "'", " "))
          twrt$ = strrepl(twrt$, """", "´´")
          If twrt$ <> "" Then
            If cid$ = "" Then
              cxid$ = s.Fields(1).value
              cx2id$ = c0id$
            Else
              If LCase(Label4.Caption) = "kontakt" Then
                cxid = s.Fields(0).value
                cx2id$ = cid$
              Else
                cxid$ = cid$
                cx2id$ = "-1"
              End If
            End If
            xid$ = form1.newid("adresstyp", "id", 8)
            If List5.List(i%) = "Vorname" Then
              cmdx$ = "insert into adresstyp (id,vid,typ,kid) values('" & xid$ & "','" & _
                cxid$ & "','Person','" & cid$ & "')"
              Call form1.sqlqry(cmdx$)
              c1md$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 8) & "','" & _
                  cxid$ & cid$ & "','Person','Vorname','" & _
                  twrt$ & "')"
              Call form1.sqlqry(c1md$)
            Else
              cmdx$ = "insert into adresstyp (id,vid,typ,wert,kid) values('" & xid$ & "','" & _
                cxid$ & "','" & _
                Mid$(List6.List(i%), InStr(List6.List(i%), ":") + 1) & "','" & _
                twrt$ & "','" & cx2id & "')"
              Call form1.sqlqry(cmdx$)
              c1md$ = "select id from adresstypen where id='" & twrt$ & "'"
              Set s1 = acc.OpenRecordset(c1md$, dbOpenDynaset, dbOpenDynaset)
              If Not s1.EOF Then
                c1md$ = "insert into adresstyp (id,vid,typ,kid) values('" & form1.newid("adresstyp", "id", 8) & "','" & _
                  cxid$ & "','" & _
                  twrt$ & "','" & _
                  cx2id & "')"
                Call form1.sqlqry(c1md$)
              End If
            End If
          End If
        Else
          If List6.List(i%) <> "telfaxhandy" Then
            If List6.List(i%) <> List6.List(i% - 1) Then
'Debug.Print s.Fields(i% - skp%).Name; "='"; s.Fields(i% - skp%).value; "'"
              cmd$ = cmd$ & "','" & trm(strrepl("" & s.Fields(i% - skp%).value, "'", " "))
            Else
              cmd$ = cmd$ & " " & trm(strrepl("" & s.Fields(i% - skp%).value, "'", " "))
            End If
            If InStr(List5.List(i%), "+") > 0 Then
              cmd$ = cmd$ + Mid$(List5.List(i%), InStr(List5.List(i%), "+") + 1)
            End If
          Else
            t_n$ = "": On Error Resume Next: t_n$ = onlynums("" & s.Fields(tnum%).value): On Error GoTo 0
            f_n$ = "": On Error Resume Next: f_n$ = onlynums("" & s.Fields(fnum%).value): On Error GoTo 0
            h_n$ = "": On Error Resume Next: h_n$ = onlynums("" & s.Fields(hnum%).value): On Error GoTo 0
            cmd$ = cmd$ & "','" & t_n$ & " " & f_n$ & " " & h_n$
          End If
        End If
      Else
        If List6.List(i%) = List6.List(i% - 1) Then
          If LCase$(List6.List(i%)) = "id" Then
            cmd$ = cmd$ & " " & datum2sql("" & s.Fields(i% - skp%).value)
          Else
            MsgBox 1 / 0
          End If
        Else
          cmd$ = cmd$ & "','" & datum2sql("" & s.Fields(i% - skp%).value)
        End If
      End If
    Else
      If Left$(List5.List(i%), 3) <> "=f:" Then
        cmd$ = cmd$ & "','" & Mid$(List5.List(i%), 2)
        skp% = skp% + 1
      Else
        ccmd$ = Mid$(List5.List(i%), 4)
        csrc$ = Mid$(ccmd$, InStr(ccmd$, "__") + 2)
        ccmd = Left$(ccmd$, InStr(ccmd$, "__") - 1)
        
        Select Case LCase(ccmd$)
          Case "penkstlid": w$ = trm(penkstlid(s.Fields(csrc$).value))
                            cmd$ = cmd$ & "','" & w$
          Case "pendatum": w$ = pendatum(s.Fields(csrc$).value)
                           cmd$ = cmd$ & "','" & w$
          Case Default:
        End Select
        skp% = skp% + 1
      End If
    End If
    While List6.List(i%) = List6.List(i% + 1)
      i% = i% + 1
      adw1$ = "": If List5.List(i%) = "PrivatTel" Then adw1$ = "PrivatTel: "
      adw1$ = "": If List5.List(i%) = "Postfach" Then adw1$ = "Postfach: "
      adw1$ = "": If List5.List(i%) = "PLZPostfach" Then adw1$ = "PLZPostfach: "
      If Not IsNull(s.Fields(i% - skp%).value) Then cmd$ = cmd$ & vbCrLf$ & adw1$ & trm(strrepl("" & s.Fields(i% - skp%).value, "'", " "))
    Wend
  Next i%
  cmd$ = cmd$ + "')"
  Debug.Print cmd$
  form1.sqlqry (cmd$)
  s.MoveNext
Wend
If List7.ListIndex < List7.ListCount - 1 Then List7.ListIndex = List7.ListIndex + 1
Call List7_DblClick
nnn% = nnn% - 1
Wend

End Sub

Private Sub Command9_Click()

i% = List5.ListIndex
If i% < 0 Then Exit Sub
If InStr(List5.List(i%), "date:") > 0 Then
  List5.List(i%) = Mid$(List5.List(i%), InStr(List5.List(i%), ":") + 1)
Else
  List5.List(i%) = "date:" + List5.List(i%)
End If

End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions

Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)


'Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, form1.getconnstr())
Label1.Caption = App.EXEName & " " & App.Major & "." & App.Minor & " - Build #" & App.Revision & Chr$(13) & Chr$(10) & Label1.Caption

List1.Clear
For i% = 0 To form1.sqla.TableDefs.Count - 1
  If InStr(form1.sqla.TableDefs(i%).name, "mysql") = 0 Then
    Set r = form1.sqla.OpenRecordset( _
      "SELECT count(*) as cnt FROM " + form1.sqla.TableDefs(i%).name, dbOpenDynaset, dbReadOnly)

    ad$ = form1.sqla.TableDefs(i%).name & " " & r!cnt & " recs"
    List1.AddItem ad$
  End If
Next i%

Show
Call Text1_Change
Call rlist7
Call Command1_Click
Text3.Text = "1"
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

Private Sub komponis_Click()

c$ = "SELECT id, Daten FROM k_loc;"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbOpenDynaset)
While Not r.EOF
  d$ = trm(r!daten)
  p% = InStr(d$, "-")
  If p% > 0 Then
    V$ = trm(Left(d$, p% - 1))
    b$ = trm(Mid(d$, p% + 1))
    If V$ <> "" Then Call form1.sqlqry("update k_loc set von='" & V$ & "' where id='" & r!id & "'")
    If b$ <> "" Then Call form1.sqlqry("update k_loc set bis='" & b$ & "' where id='" & r!id & "'")
  End If
  r.MoveNext
Wend

End Sub

Private Sub kompos_Click()
c$ = "SELECT * FROM w_loc;"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbOpenDynaset)
While Not r.EOF
  knr$ = Left(trm(r!id), 3)
  Call form1.sqlqry("update w_loc set komponistennummer='" & knr$ & "' where id='" & r!id & "'")
  knr$ = strrepl(trm(r!Opusname1), """", "´´")
  wert$ = trm(" " & r!nummer)
  If wert$ <> "" Then
    knr$ = knr$ & " Nr. " & wert$
  End If
  wert$ = trm(" " & r!Tonart): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = trm(" " & r!Opusbezeichnung): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = trm(" " & r!OpusNummer): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = strrepl(trm(" " & r!Opusbezeichnung), """", "´´"): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  wert$ = strrepl(trm(" " & r!Opusname2), """", "´´"): If wert$ <> "" Then knr$ = knr$ & " " & wert$
  Call form1.sqlqry("update w_loc set name='" & strrepl(knr$, """", "´´") & "' where id='" & r!id & "'")
  cnt = 0
  cnt = cnt + 1: knr$ = strrepl(trm(r!s1), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s2), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s3), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s4), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s5), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s6), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s7), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s8), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s9), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s10), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s11), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s12), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s13), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  cnt = cnt + 1: knr$ = strrepl(trm(r!s14), "'", "´"): If knr$ <> "" Then Call form1.sqlqry("insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & form1.newid("sbz_loc", "id", 40) & "','" & r!id & "','" & knr$ & "'," & trm(str$(cnt)) & ");")
  r.MoveNext
Wend

End Sub

Private Sub List1_DblClick()
Call Command3_Click
End Sub

Private Sub List2_DblClick()
Call Command3_Click
End Sub

Private Sub List3_DblClick()
Call Command4_Click
End Sub

Private Sub List4_DblClick()
Call Command4_Click
End Sub

Private Sub List5_dblClick()
i% = List5.ListIndex
List5.RemoveItem i%
List6.RemoveItem i%

End Sub


Private Sub List6_DblClick()
i% = List6.ListIndex
wert$ = List6.List(i%): If Left$(wert$, 1) = "=" Then wert$ = Mid$(wert$, 2)
wert$ = InputBox(transe("Wert von") + " " + List6.List(i%), "Festwertzuweisung", wert$)
List5.RemoveItem i%
List5.AddItem "=" + wert$
wert$ = List6.List(i%)
List6.RemoveItem i%
List6.AddItem wert$

End Sub

Private Sub List7_Click()
Text2.Text = List7.List(List7.ListIndex)
End Sub

Private Sub List7_DblClick()

o% = FreeFile
List5.Clear
List6.Clear
If List7.ListIndex < 0 Then Exit Sub
Open List7.List(List7.ListIndex) For Input As #o%
Text2.Text = List7.List(List7.ListIndex)
Line Input #o%, l$
Label3.Caption = l$
For i% = 0 To List2.ListCount - 1
  If Left$(List2.List(i%), Len(l$)) = l$ Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%
Line Input #o%, l$
Label4.Caption = l$
For i% = 0 To List1.ListCount - 1
  If Left$(List1.List(i%), Len(l$)) = l$ Then
    List1.ListIndex = i%
    i% = List1.ListCount
  End If
Next i%
Call Command3_Click
Line Input #o%, cnt
While cnt > 0
  Line Input #o%, l$
  List5.AddItem l$
  Line Input #o%, l$
  List6.AddItem l$
  cnt = cnt - 1
Wend
Close #o%

End Sub

Private Sub Text1_Change()
If exist(Text1.Text) > 0 Then
  Command3.Enabled = True
  Command1.Enabled = True
Else
  Command3.Enabled = False
  Command1.Enabled = False
End If

End Sub
Function penkstlid(w$) As String
Dim r$, rr$
r$ = Left$(w$, Len(w$) - 10)
rr$ = Right$(w$, 10)
While isdigit(Left$(rr$, 1))
  r$ = r$ + Left$(rr$, 1)
  rr$ = Mid$(rr$, 2)
Wend
penkstlid = r$
End Function
Function pendatum(w$) As String
Dim r$, rr$
r$ = Left$(w$, Len(w$) - 10)
rr$ = Right$(w$, 10)
While isdigit(Left$(rr$, 1)) = 0
  r$ = r$ + Left$(rr$, 1)
  rr$ = Mid$(rr$, 2)
Wend
pendatum = datum2sql(rr$)
End Function

