VERSION 5.00
Begin VB.Form tsts 
   Caption         =   "Tests ... "
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command22 
      Caption         =   "Lösche Benutzer"
      Height          =   255
      Left            =   9000
      TabIndex        =   62
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      Caption         =   "create"
      Height          =   255
      Left            =   3120
      TabIndex        =   61
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Alle"
      Height          =   255
      Left            =   3120
      TabIndex        =   60
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox enced 
      Height          =   1215
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   59
      Text            =   "tsts.frx":0000
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox unenced 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   58
      Text            =   "tsts.frx":0006
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   56
      Text            =   "Text5"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command19 
      Caption         =   "-->"
      Height          =   255
      Left            =   3120
      TabIndex        =   55
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<--"
      Height          =   255
      Left            =   3120
      TabIndex        =   54
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "löschen für referenz.mdb"
      Height          =   255
      Left            =   9000
      TabIndex        =   53
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   27000
      Left            =   1440
      Top             =   2520
   End
   Begin VB.CommandButton Command17 
      Caption         =   "stop"
      Height          =   255
      Left            =   720
      TabIndex        =   52
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "go"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "kill ""My Docs"""
      Height          =   255
      Left            =   1440
      TabIndex        =   50
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command12e"
      Height          =   255
      Left            =   1440
      TabIndex        =   49
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1920
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12w"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Werke löschen"
      Height          =   255
      Left            =   9000
      TabIndex        =   46
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<-- command"
      Height          =   255
      Left            =   9000
      TabIndex        =   45
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   675
      Left            =   4320
      TabIndex        =   44
      Text            =   "Text4"
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "medien2adr"
      Height          =   315
      Left            =   1920
      TabIndex        =   43
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3120
      TabIndex        =   42
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "chk_tp"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "chk_init"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   15
      Left            =   10680
      TabIndex        =   37
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   14
      Left            =   9600
      TabIndex        =   36
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   13
      Left            =   8520
      TabIndex        =   33
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   12
      Left            =   7440
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   11
      Left            =   6360
      TabIndex        =   29
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   10
      Left            =   5280
      TabIndex        =   28
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   9
      Left            =   4200
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   8
      Left            =   3120
      TabIndex        =   24
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   7
      Left            =   10680
      TabIndex        =   21
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   6
      Left            =   9600
      TabIndex        =   20
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   5
      Left            =   8520
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   4
      Left            =   7440
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   3
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   2
      Left            =   5280
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "cr_tplan"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "cr_prog"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "1"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Alles Löschen"
      Height          =   255
      Left            =   10320
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "don&e"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   4560
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cr_adresse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   15
      Left            =   10680
      TabIndex        =   39
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   9600
      TabIndex        =   38
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   13
      Left            =   8520
      TabIndex        =   35
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   12
      Left            =   7440
      TabIndex        =   34
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   6360
      TabIndex        =   31
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   30
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   27
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   10680
      TabIndex        =   23
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   9600
      TabIndex        =   22
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   8520
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   18
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "tsts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t1yp$(99)
Dim wrkJet As Workspace
Dim sqla As Database
Dim werk$(999), o_r$(499), s_r$(499), t_r$(499), d_r$(499), p_r$(499), v_r$(499), tor%, tsr%, ttr%, tpr%, tvr%
Dim h_r$(499), thr%, h1_r$(499), th1r%
Dim tlist(1 To 9), nlist(1 To 9), t1_cnt%, stressinit As Integer, tpp$(1 To 4)

Private Sub Command1_Click()
Dim n%, rtmp As Recordset, ky$

Set rtmp = sqla.OpenRecordset("SELECT id from adresstypen", dbOpenDynaset, dbReadOnly)
p% = 0
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  t1yp$(p%) = rtmp!id
  p% = p% + 1
  rtmp.MoveNext
Wend
rtmp.Close
mt% = p% - 1

merk$ = Command1.Caption
For p% = 1 To Val(Text1.text)
Command1.Caption = p%
neuid$ = mknam(6 + Int(Rnd * 4))
N1ame$ = mknam(4 + Int(Rnd * 4))
ky$ = neuid$ & ", " + N1ame$
N1ame$ = N1ame$ & " " + neuid$
neuid$ = ky$
strasse$ = mknam(10) + str$(Int(Rnd * 90))
plzort$ = Trim(str$(Int(Rnd * 8999 + 1000))) & " " + mknam(10)
tel$ = Trim(str$(Int(Rnd * 9999))) & " / " + Trim(str$(Int(Rnd * 9999)))
fax$ = Trim(str$(Int(Rnd * 9999))) & " / " + Trim(str$(Int(Rnd * 9999)))
handy$ = Trim(str$(Int(Rnd * 9999))) & " / " + Trim(str$(Int(Rnd * 9999)))
telfaxhandy = onlynums(tel$) & " " + onlynums(fax$) + " " + onlynums(handy$)
s$ = "insert into adresse (id,name,strasse,ort,tel,fax,handy,telfaxhandy) values('" + ky$ + "','" + N1ame$ + "','" + strasse$ + "','" + plzort$ + "','" + tel$ + "','" + fax$ + "','" + handy$ + "','" + telfaxhandy + "')"
form1.sqlqry (s$)
s$ = "insert into adresstyp (id,vid,typ) values('" + form1.newid("adresstyp", "id", 20) + "','" + neuid$ + "','Random')"
form1.sqlqry (s$)
DoEvents
For i% = 0 To mt%
  If t1yp$(i%) <> "Random" Then
    If Rnd > 0.9 Then
      s$ = "insert into adresstyp (id,vid,typ,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + neuid$ + "','" + t1yp$(i%) + "','-1')"
      form1.sqlqry (s$)
      DoEvents
    End If
  End If
Next i%
n% = Rnd * 5
While n% > 0
  N1ame$ = mknam(6) + " " + mknam(6)
  tel$ = Trim(str$(Int(Rnd * 9999))) + " / " + Trim(str$(Int(Rnd * 9999)))
  fax$ = Trim(str$(Int(Rnd * 9999))) + " / " + Trim(str$(Int(Rnd * 9999)))
  handy$ = Trim(str$(Int(Rnd * 9999))) + " / " + Trim(str$(Int(Rnd * 9999)))
  telfaxhandy = onlynums(tel$) + " " + onlynums(fax$) + " " + onlynums(handy$)
  kid$ = mknam(10)
  s$ = "insert into kontakt (id,vid,name,tel,fax,handy,telfaxhandy) values('" + kid$ + "','" + neuid$ + "','" + N1ame$ + "','" + tel$ + "','" + fax$ + "','" + handy$ + "','" + telfaxhandy + "')"
  form1.sqlqry (s$)
  
  n% = n% - 1
Wend
Next p%
Command1.Caption = merk$

End Sub

Private Sub Command10_Click()
MousePointer = 11: DoEvents
form1.sqlqry ("delete from b_loc")
form1.sqlqry ("delete from sbz_loc")
form1.sqlqry ("delete from k_loc")
form1.sqlqry ("delete from w_loc")
MousePointer = 0

End Sub



Private Sub Command11_Click()
Dim aKey() As Byte

aKey = Text5.text
Call blf_KeyInit(aKey())

l$ = enced.text
unenced.text = blf_StringDec(l$)
enced.text = ""
End Sub

Private Sub Command12_Click()
If form1.List1.ListCount > 0 Then
  form1.List1.ListIndex = 0
  DoEvents
  DoEvents
  Call form1.List1_DblClick
  DoEvents
  Text1.text = "1"
  Call Command1_Click
  DoEvents
  X = Shell("D:\Office2000\PFiles\MSOffice\Office\winword.exe", 1)
  For n = 1 To 5
  For i = 1 To 7
    SendKeys shwAdrDetail.datf(i).text, 1
    SendKeys "{Enter}", 1
    DoEvents
  Next i
  Next n
End If

t1_cnt% = 40

Timer1.Interval = 200
Timer1.Enabled = True
End Sub

Private Sub Command13_Click()

stressinit = 1
shwAdrDetail.BackColor = form1.cleancolor()

form1.Combo1.text = mknam(4)
DoEvents
shwAdrDetail.BackColor = form1.cleancolor()
Call form1.combo1_Change
DoEvents
DoEvents

If Rnd > 0.5 Then
  Call Command12_Click
Else
  Call Command14_Click
End If
End Sub

Private Sub Command14_Click()
If form1.List1.ListCount > 0 Then
  form1.List1.ListIndex = 0
  DoEvents
  DoEvents
  Call form1.List1_DblClick
  DoEvents
  Text1.text = "1"
  Call Command1_Click
  DoEvents
  X = Shell("D:\Office2000\PFiles\MSOffice\Office\excel.exe", 1)
  For n = 1 To 10
  For i = 1 To 7
    SendKeys shwAdrDetail.datf(i).text, 1
    SendKeys "{Enter}", 1
  Next i
  Next n
End If

t1_cnt% = 40

Timer1.Interval = 300
Timer1.Enabled = True

End Sub

Private Sub Command15_Click()
  p$ = "c:"
  i% = 1
  hp$ = ""
  While p$ = "c:" And Environ$(i%) <> ""
    t$ = LCase(Environ$(i%))
    Debug.Print Environ$(i%)
    If InStr(LCase(t$), "homedrive=") = 1 Then hd$ = Mid$(Environ$(i%), 11)
    If InStr(LCase(t$), "homepath=") = 1 Then hp$ = Mid$(Environ$(i%), 10)
    If InStr(LCase(t$), "username=") = 1 Then u$ = Mid$(Environ$(i%), 10)
    i% = i% + 1
  Wend
  If hp$ <> "" Then
    If InStr(hp$, ":") = 0 Then hp$ = hd$ & hp$
    p$ = hp$
  End If
If hp$ = "" Then End
hp$ = hp$ & "\My Documents"
tr = Dir(hp$ & "\*.doc")
While tr <> ""
  Kill hp$ & "\" & tr
  tr = Dir
  DoEvents
Wend
tr = Dir(hp$ & "\*.xls")
While tr <> ""
  Kill hp$ & "\" & tr
  tr = Dir
  DoEvents
Wend
tr = Dir(hp$ & "\*.mdb")
While tr <> ""
  Kill hp$ & "\" & tr
  tr = Dir
  DoEvents
Wend

End Sub

Public Sub Command16_Click()
Timer3.Interval = Int(13000 + Rnd * 17000)
Timer3.Enabled = True

End Sub

Private Sub Command17_Click()
Timer3.Enabled = False

End Sub

Private Sub Command18_Click()
For i% = 0 To sqla.TableDefs.Count - 1
  c$ = "delete from " & sqla.TableDefs(i%).name
  form1.sqlqry (c$)
Next i%
End
End Sub

Private Sub Command19_Click()
Dim aKey() As Byte

aKey = Text5.text

Call blf_KeyInit(aKey())
enced.text = blf_StringEnc(unenced.text)
unenced.text = ""
End Sub

Private Sub Command2_Click()
Hide
Unload tsts

End Sub

Private Sub Command22_Click()
MousePointer = 11: DoEvents
form1.sqlqry ("delete from benutzergruppen")
form1.sqlqry ("delete from benutzerdaten")
form1.sqlqry ("insert into benutzerdaten (id,name) values('system','Systemweite Einstellungen');")
MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim rtmp As Recordset, s As Recordset

For i% = 0 To 15
  List1(i%).Clear
  Label1(i%).Caption = ""
Next i%
Load adrselect

Set rtmp = sqla.OpenRecordset( _
    "SELECT id FROM adresstypen", dbOpenDynaset, dbReadOnly)

i% = 0
While Not rtmp.EOF And i% < 16
  Label1(i%).Caption = rtmp!id
  Call adrselect.sel_init("", Label1(i%).Caption)
  Call adrselect.rlist1("")
  For j% = 0 To adrselect.List1.ListCount - 1
    List1(i%).AddItem adrselect.List1.List(j%)
  Next j%
  DoEvents
  i% = i% + 1
  rtmp.MoveNext
Wend

Unload adrselect

List2.Clear

Set rtmp = sqla.OpenRecordset("SELECT * FROM programm", dbOpenDynaset, dbReadOnly)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!programmid) Then
    List2.AddItem rtmp!programmid
  End If
  rtmp.MoveNext
Wend
rtmp.Close

Command8.Enabled = True

End Sub

Private Sub Command4_Click()

MousePointer = 11: DoEvents
form1.sqlqry ("delete from kurse")
form1.sqlqry ("delete from adresse")
form1.sqlqry ("delete from webhosts")
form1.sqlqry ("delete from access_log")
form1.sqlqry ("delete from adresstyp")
form1.sqlqry ("delete from adressgruppen")
form1.sqlqry ("delete from adressgruppenindex")
form1.sqlqry ("delete from anreden")
form1.sqlqry ("delete from auftritt")
form1.sqlqry ("delete from auftritthigru")
form1.sqlqry ("delete from dochist")
form1.sqlqry ("delete from bkfirmen")
form1.sqlqry ("delete from kontakt")
form1.sqlqry ("delete from finanzen")
form1.sqlqry ("delete from programm")
form1.sqlqry ("delete from benutzerdaten")
form1.sqlqry ("delete from benutzergruppen")
form1.sqlqry ("delete from programmliste")
form1.sqlqry ("delete from taliste")
form1.sqlqry ("delete from mailsafe")
form1.sqlqry ("delete from dictionary_taboo")
form1.sqlqry ("delete from mailip")
form1.sqlqry ("delete from talisted")
form1.sqlqry ("delete from tplan")
form1.sqlqry ("delete from tpwernoch")
form1.sqlqry ("delete from tpprogli")
form1.sqlqry ("delete from todolist")
form1.sqlqry ("delete from alarmliste")
form1.sqlqry ("delete from poplist")
form1.sqlqry ("delete from hbabos")
form1.sqlqry ("delete from hblist")
form1.sqlqry ("delete from hbplist")
form1.sqlqry ("delete from hbpstatus")
form1.sqlqry ("delete from hbabotermine")
form1.sqlqry ("delete from kassenbuch")
form1.sqlqry ("delete from aut_werke")
form1.sqlqry ("delete from b_loc")
form1.sqlqry ("delete from bkurse")
form1.sqlqry ("delete from bplan")
form1.sqlqry ("delete from sysvars where instr(owner,'sys')=1")
form1.sqlqry ("delete from sysvars where instr(owner,'blacklist')=1")
For i% = 0 To sqla.TableDefs.Count - 1
  If Left$(sqla.TableDefs(i%).name, 4) = "usr_" Then
    form1.sqlqry ("delete from " & sqla.TableDefs(i%).name)
  End If
Next i%
For i% = 0 To sqla.TableDefs.Count - 1
  If Left$(sqla.TableDefs(i%).name, 4) = "tmp_" Then
    form1.sqlqry ("delete from " & sqla.TableDefs(i%).name)
  End If
Next i%
MousePointer = 0

End Sub

Private Sub Command5_Click()
Dim r As Recordset

Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM w_loc", dbOpenDynaset, dbReadOnly)
maxw = r!cnt

Set r = sqla.OpenRecordset("SELECT id as wid FROM w_loc", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To 999
  r.MoveNext
  werk$(j%) = r!wid
Next j%

merk$ = Command5.Caption
For p% = 1 To Val(Text2.text)
  Command5.Caption = p%
  DoEvents
  pid$ = form1.newid("programm", "programmid", 4) & Int(Rnd * 12 + 1)
  cmd$ = "INSERT INTO programm (programmID) VALUES('" & pid$ & "')"
  Call form1.sqlqry(cmd$)
  bis% = Int(Rnd * 4 + 3)
  For cnt = 0 To bis%
    pos = Int(Rnd * 1000)
    neuwerkid$ = werk$(pos)
    cmd$ = "insert into programmliste (id,programmid,werkid,position) values('" + form1.newid("programmliste", "id", 20) + "','" + pid$ + "','" + neuwerkid$ + "'," & Trim(cnt) & ")"
    Call form1.sqlqry(cmd$)
  Next cnt

Next p%
Command5.Caption = merk$
End Sub

Private Sub Command6_Click()
Dim r As Recordset, mtid As String
' s_r$(399), v_r$(399), k_r$(399), t_r$(399)
Randomize
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Künstler'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: tsr% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Künstler'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  s_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Tourneeleitung'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: ttr% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Tourneeleitung'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  t_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Orchester'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: tor% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Orchester'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  o_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Dirigent'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: tdr% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Dirigent'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  d_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Veranstalter'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: tvr% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Veranstalter'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  v_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Hotel'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: th1r% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Hotel'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  h1_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Halle'));", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: thr% = maxw
Set r = sqla.OpenRecordset("SELECT adresse.id as wid FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((adresstyp.typ)='Halle'));", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  h_r$(j%) = r!wid
Next j%
Set r = sqla.OpenRecordset("SELECT count(*) as cnt FROM programm;", dbOpenDynaset, dbReadOnly)
maxw = imin(r!cnt, 499) - 2: tpr% = maxw
Set r = sqla.OpenRecordset("SELECT programmid as wid FROM programm;", dbOpenDynaset, dbReadOnly)
r.MoveFirst
For j% = 0 To maxw
  r.MoveNext
  p_r$(j%) = r!wid
Next j%
merk$ = Command6.Caption
For p% = 1 To Val(Text3.text)
  Command6.Caption = p%
  DoEvents
  o$ = o_r$(Rnd * tor%)
  k$ = s_r$(Rnd * tsr%)
  d$ = d_r$(Rnd * tdr%)
  von = CDate(CDate("1.1.2001") + (Int(Rnd * 2000))): mtid = Trim(Month(von)): ytid = Year(von)
  If Len(mtid) = 1 Then mtid = "0" & mtid
  bis = CDate(von + CDate(Int(Rnd * 30)))
  If Rnd > 0.5 Then
    pid$ = Left$(o$, 6) & " " & mtid & " " & ytid
    cmd$ = "INSERT INTO tplan (ID,von,bis,Hauptperson,solist,tourneeleitung,orchester,dirigent) VALUES('" & pid$ & "','" & datum2sql(von) & "','" & datum2sql(bis) & "','Orchester','" & k$ & "','" & t_r$(Rnd * ttr%) & "','" & o$ & "','" & d$ & "')"
    attyp$ = "Orchesterauftritt": ptyp$ = "Orchesterprobe": pfn$ = "Saal": ppt$ = "Orchester": pppt$ = o$
  Else
    pid$ = Left$(k$, 6) & " " & mtid & " " & ytid
    cmd$ = "INSERT INTO tplan (ID,von,bis,Hauptperson,solist,tourneeleitung,orchester,dirigent) VALUES('" & pid$ & "','" & datum2sql(von) & "','" & datum2sql(bis) & "','Künstler','" & k$ & "','" & t_r$(Rnd * ttr%) & "','" & o$ & "','" & d$ & "')"
    attyp$ = "Künstlerauftritt": ptyp$ = "Künstlerprobe": pfn$ = "Halle": ppt$ = "Künstler": pppt$ = k$
  End If
  Call form1.sqlqry(cmd$)
  For i5% = 1 To 4
    nid$ = form1.newid("tpprogli", "id", 8)
    op$ = p_r$(Rnd * tpr%): tpp$(i5%) = op$
    cmd$ = "INSERT INTO tpprogli (ID,tpid,prgid) VALUES('" & nid$ & "','" & pid$ & "','" & op$ & "')"
    Call form1.sqlqry(cmd$)
  Next i5%

d0 = CDate(von)
nid$ = form1.newid("auftritt", "id", 20)
form1.sqlqry ("INSERT INTO auftritt (id, ort,zeit,TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" & ohnePLZ(form1.ortausadr(hxr$)) & "','00:00','" + pid$ + _
                 "','Tournee','Tournee " + pid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & nid$ & _
                 "','Tournee','Dirigent','" & _
                 d$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & nid$ & _
                 "','Tournee','Orchester','" & _
                 o$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & nid$ & _
                 "','Tournee','Solist','" & _
                 k$ & "')")
Dt = d0
d1 = CDate(bis)
While Dt <= d1
  hxr$ = h_r$(Rnd * thr%)
  h1xr$ = h_r$(Rnd * th1r%)
  naid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id,ort, zeit,TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 naid$ & "','" & ohnePLZ(form1.ortausadr(hxr$)) & "','20:00','" + pid$ + _
                 "','" & attyp$ & "','" + pid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Programm','" & _
                 tpp$(Int(Rnd * 4 + 1)) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Dirigent','" & _
                 d$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Orchester','" & _
                 o$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Künstler','" & _
                 k$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Veranstalter','" & _
                 v_r$(Rnd * tvr%) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Halle','" & _
                 hxr$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Hotel','" & _
                 h1xr$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Ankunft_Hotel','14:00')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Abreise_Hotel','09:00')")
  hnw$ = "": For i5% = 1 To 20: hnw$ = hnw$ & " " & mkkey(Int((Rnd * 8) + 2)): Next i5%
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Hinweise','" & _
                 hnw$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & attyp$ & "','Honorar','" & _
                 fixeur(1000 + Rnd * 4000) & " EUR')")

'############################
  naid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id,ort, zeit,TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 naid$ & "','" & ohnePLZ(form1.ortausadr(hxr$)) & "','17:00','" + pid$ + _
                 "','" & ptyp$ & "','" + pid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & ptyp$ & "','Dauer','60 Min.')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & ptyp$ & "','" & pfn$ & "','" & hxr$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','" & ptyp$ & "','" & ppt$ & "','" & pppt & "')")

'############################
  naid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id,ort, zeit,TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 naid$ & "','" & ohnePLZ(form1.ortausadr(hxr$)) & "','14:00','" + pid$ + _
                 "','Hotelaufenthalt','" + pid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','Hotelaufenthalt','Hotel','" & h1xr$ & "')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','Hotelaufenthalt','Anreise','14:00')")
  form1.sqlqry ("INSERT INTO auftritthigru (id, auftrittsID,Auftrittstyp,Feldname,felddaten) VALUES ('" + _
                 form1.newid("auftritthigru", "id", 10) & "','" & naid$ & _
                 "','Hotelaufenthalt','Abreise','09:00')")

Dt = CDate(Dt) + 1
Wend

Next p%
Command6.Caption = merk$
End Sub

Sub tpchk(a$)
Dim rtmp As Recordset, s As Recordset

Load tplan
tplan.Caption = a$
Call tplan.rlists
a$ = cut_d1(a$, " ")
tlist(1) = "Tourneeleitung"
tlist(2) = "Orchester"
tlist(3) = "Veranstalter"
tlist(4) = "Dirigent"
tlist(5) = "Künstler"
nlist(1) = 1
nlist(2) = 3
nlist(3) = 9
nlist(4) = 8
nlist(5) = 13
For i% = 0 To tplan.List1.ListCount - 1
  tplan.List1.ListIndex = i%
  DoEvents
  For k% = 1 To 5
    j% = nlist(k%)
    If tplan.Text1(j%).text = "" Then
      w$ = tlist(k%)
      tplan.Text1(j%).text = rndid(w$)
      Call tplan.Text1_LostFocus(j%)
    End If
  Next k%
Next i%

For i% = 0 To tplan.List1.ListCount - 1
  tplan.List1.ListIndex = i%
  DoEvents
  tpid$ = tplan.List1.List(i%)
  j% = Int(Rnd * 5 + 1)
  While j% > 0
    nwert$ = tsts.rndid("Programm")
    cmd$ = "insert into tpprogli (id,tpid,prgid,_desc) values('" + form1.newid("tpprogli", "id", 18) + "','" + tpid$ + "','" + nwert$ + "'," & Trim(10 * (100 - j%)) & ")"
    form1.sqlqry (cmd$)
    j% = j% - 1
  Wend
  Call tplan.Command26_Click
  For j% = 1 To tplan.List6.ListCount - 1
    l$ = tplan.List6.List(j%)
    If InStr(l$, "Neuer Auftritt") > 0 Then
      aid$ = Mid$(l$, InStr(l$, "(AID:") + 5)
      form1.sqlqry ("update auftritt set auftrittstyp='" + a$ + "probe' where id='" + aid$ + "'")
      form1.sqlqry ("update auftritt set zeit='16:00:00' where id='" + aid$ + "'")
    End If
  Next j%
  Call tplan.Command26_Click
  For j% = 1 To tplan.List6.ListCount - 1
    l$ = tplan.List6.List(j%)
    If InStr(l$, "Neuer Auftritt") > 0 Then
      aid$ = Mid$(l$, InStr(l$, "(AID:") + 5)
      form1.sqlqry ("update auftritt set auftrittstyp='" + a$ + "auftritt' where id='" + aid$ + "'")
      form1.sqlqry ("update auftritt set zeit='20:00:00' where id='" + aid$ + "'")
    End If
  Next j%
  Call tplan.Command26_Click
  For j% = 1 To tplan.List6.ListCount - 1
    l$ = tplan.List6.List(j%)
    If InStr(l$, "Neuer Auftritt") > 0 Then
      aid$ = Mid$(l$, InStr(l$, "(AID:") + 5)
      form1.sqlqry ("update auftritt set auftrittstyp='Hotelaufenthalt' where id='" + aid$ + "'")
      form1.sqlqry ("update auftritt set zeit='16:00:00' where id='" + aid$ + "'")
    End If
  Next j%
  Call tplan.Command26_Click
  For j% = 1 To tplan.List6.ListCount - 1
    l$ = tplan.List6.List(j%)
    If InStr(l$, "Neuer Auftritt") > 0 Then
      aid$ = Mid$(l$, InStr(l$, "(AID:") + 5)
      form1.sqlqry ("update auftritt set auftrittstyp='Reise' where id='" + aid$ + "'")
      form1.sqlqry ("update auftritt set zeit='08:30:00' where id='" + aid$ + "'")
    End If
  Next j%
Next i%
Unload tplan

End Sub

Private Sub Command7_Click()
Dim r As Recordset


dn$ = form1.s0dir() + "\" + form1.medien() + "\"
tr = Dir(dn$ + "\*.*", vbDirectory)
While tr <> ""
  If tr <> "." And tr <> ".." And (GetAttr(dn$ + "\" + tr) And vbDirectory) = vbDirectory Then
    Set r = sqla.OpenRecordset("SELECT id from adresse where id='" + tr + "'", dbOpenDynaset, dbReadOnly)
    If r.EOF Then
      s$ = "insert into adresse (id,name) values('" + tr + "','" + tr + "')"
      form1.sqlqry (s$)
    End If
  End If
  tr = Dir
Wend
End Sub

Private Sub Command8_Click()

Call tpchk("Orchester - Projekt")
Call tpchk("Künstler - Projekt")

End Sub



Private Sub Command9_Click()
'form1.sqlqry (Text4.Text)

End Sub

Private Sub Form_Load()

Randomize
stressinit = 0
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)

Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
dbpara$ = form1.getconnstr()
If dbpara$ <> "msaccessmdb" Then
  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
Else
  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
End If
Text5.text = "jhfzer6498"


End Sub


Public Function rndid(typ$) As String

rndid = ""

If LCase$(typ$) = "programm" Then
  t$ = List2.List(Int(Rnd * List2.ListCount))
  rndid = Trim(t$)
  Exit Function
Else
  For i% = 0 To 15
    If Label1(i%).Caption = typ$ Then
      t$ = List1(i%).List(Int(Rnd * List1(i%).ListCount))
      rndid = Trim(Left$(t$, InStr(t$, "(") - 1))
      Exit Function
    End If
  Next i%
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)


End Sub

Private Sub Timer1_Timer()
If t1_cnt% > 0 Then
  t1_cnt% = t1_cnt% - 1
  z$ = Chr$(Rnd * 25 + Asc("a"))
  SendKeys z$, 1
  If Rnd > 0.9 Then
    If Rnd > 0.1 Then
      SendKeys " ", 1
    Else
      SendKeys "{Enter}", 1
    End If
  End If
Else
  Timer1.Enabled = False
  shwAdrDetail.BackColor = form1.cleancolor()
  SendKeys "%{F4}", 1: DoEvents: DoEvents: DoEvents
  SendKeys "{Enter}", 1: DoEvents: DoEvents: DoEvents
  shwAdrDetail.BackColor = form1.cleancolor()
  SendKeys mknam(8), 1: DoEvents: DoEvents: DoEvents
  SendKeys "{Enter}", 1: DoEvents: DoEvents: DoEvents
  shwAdrDetail.BackColor = form1.cleancolor()
  stressinit = 0
End If
End Sub

Private Sub Timer3_Timer()
Dim rtmp As Recordset

If stressinit = 1 Then Exit Sub
stressinit = 1
cmd$ = "select id from todolist where an='" + form1.getuserid() + "' order by datum,zeit"
On Error Resume Next
Set rtmp = sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  If Not rtmp.EOF Then
    Load todolist
    Call todolist.SetFocus
    Call todolist.Command4_Click
    DoEvents
    todolist.List1.ListIndex = 0
    DoEvents
    If InStr(todolist.List1.List(i), "runcmd") > 0 Then
      cmd$ = todolist.Text1.text
      Call todolist.delme_Click
      Select Case cmd$
        Case "c13": Call Command13_Click: Exit Sub
        Case "cradr": Call Command1_Click: Exit Sub
        Case "kmd": Command15_Click: Exit Sub
        Case "schüss": cmd$ = "delete from todolist where an='" + form1.getuserid() + "'"
                     Call form1.sqlqry(cmd$)
                     Unload form1
                     DoEvents
                     End
        Case Else: stressinit = 0
      End Select
      DoEvents
      Exit Sub
    Else
      Call todolist.delme_Click
    End If
    Unload todolist
  End If
End If

DoEvents
stressinit = 0

End Sub
