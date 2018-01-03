VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form waehrung 
   Caption         =   "Währungen"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2040
      Picture         =   "waehrung.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Neue Kurse eingeben"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Grafik"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "waehrung.frx":1152
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   960
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xport"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Alle importieren"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Import"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Import"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "waehrung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim wrkJet As Workspace
'Dim sqla As Database, dbpara$
Dim nflds As Integer

Sub rlist1()
Dim w As Recordset

d2infile = "waehrung": d2insub = "rlist1"
cmd$ = "select * from waehrung order by id"
Set w = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not w.EOF
  List1.AddItem w!id & ", " & w!name & ", " & w!webname & ", " & w!srchname
  w.MoveNext
Wend

cmd$ = "select top 1 id from kurse order by id desc"
Set w = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
If Not w.EOF Then
  waehrung.Caption = "Währungskurse bis " & Left(w!id, 10)
End If


End Sub

Sub rlist3()

d2infile = "waehrung": d2insub = "rlist3"
List3.Clear
tr = Dir(form1.getusersetting("Waehrungskurse") + "\*.sqk")
While tr <> ""
  List3.AddItem tr
  tr = Dir
Wend
tr = Dir(form1.s0dir() + "\kurse.txt")
If tr <> "" Then
  List3.AddItem tr
End If

End Sub
Private Sub Command1_Click()
Dim w As Recordset, i%, o%, fn$, l$, dtg$, kw As Double, k$, wid$

d2infile = "waehrung": d2insub = "Command1_Click"
Command1.Enabled = False
i% = List3.ListIndex
If i% < 0 Then Exit Sub

If LCase(List3.List(i%)) = "kurse.txt" Then

  List3.RemoveItem i%
  o% = FreeFile
  fn$ = form1.s0dir() + "\kurse.txt"
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    If Left$(l$, 1) <> "#" Then
      List2.AddItem l$
      dtg$ = datum2sql(cut_d1(l$, ":")): l$ = cut_d2bis(l$, ":")
      wid$ = cut_d1(l$, ":"): l$ = cut_d2bis(l$, ":")
      kw = var2dbl(l$)
      cmd$ = "delete from kurse where id='" + dtg$ + wid$ + "'": form1.sqlqry (cmd$)
      If kw <> 0 Then
        cmd$ = "insert into kurse (id,wid,kurs,einheit) values('" + dtg$ + wid$ + "','" + wid$ + "'," + trm(strrepl(trm(kw), ",", ".")) + ",1)": form1.sqlqry (cmd$)
      End If
    End If
  Wend
  Close #o%
  On Error Resume Next
  Kill fn$
  On Error GoTo 0

Else

dtg$ = Left(List3.List(List3.ListIndex), InStr(List3.List(List3.ListIndex), ".") - 1)
dtg$ = Left$(dtg$, 4) + "-" + Mid$(dtg$, 5, 2) + "-" + Mid$(dtg$, 7, 2)
While List2.ListCount > 0
  If trm(List2.List(i%)) <> "" Then
    l$ = Left$(List2.List(i%), InStr(List2.List(i%), ":") - 1)
    cmd$ = "select id from waehrung where srchname='" + l$ + "'"
    Set w = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
    If Not w.EOF Then
      wid$ = w!id
      id$ = dtg$ + wid$
      k$ = strrepl(Mid$(List2.List(i%), InStr(List2.List(i%), ":") + 1), ",", ".")
      cmd$ = "select kurs from kurse where id='" + id$ + "'"
      Set w = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
      If Not w.EOF Then
        If w!kurs <> k$ Then form1.sqlqry ("update kurse set kurs='" + k$ + "' where id='" + id$ + "'")
      Else
        form1.sqlqry ("insert into kurse (id,wid,kurs,einheit) values('" + _
          id$ + "','" + wid$ + "','" + k$ + "','1')")
      End If
    End If
  End If
  List2.RemoveItem 0
  DoEvents
Wend
Kill form1.getusersetting("Waehrungskurse") + "\" + List3.List(List3.ListIndex)
List3.RemoveItem List3.ListIndex

End If

End Sub

Private Sub Command2_Click()
d2infile = "waehrung": d2insub = "Command2_Click"
While List3.ListCount > 0
  List3.ListIndex = 0
  Call Command1_Click
Wend
End Sub

Private Sub Command27_Click()
Dim fn$, o%, X

fn$ = form1.s0dir() + "\kurse.txt"
If nexist(fn$) Then
  o% = FreeFile
  Open fn$ For Output As #o%
  Print #o%, "######################################################################"
  Print #o%, "# To enter exchange rates use this format:"
  Print #o%, "# Date:Currency:Rate"
  Print #o%, "# Date: day.month.4-digit-year (1.10.2009)"
  Print #o%, "# Currency: Use abbreviation from the list of currencies."
  Print #o%, "# Rate: like 1000,472"
  Print #o%, "# Lines starting with # are ignored."
  Print #o%, "# Example: 1.10.2009:USD:1,4729"
  Print #o%, "######################################################################"
  Close #o%
End If
X = Shell("notepad.exe " + fn$, 1)
List3.Clear
List3.AddItem "kurse.txt"

End Sub

Private Sub Command3_Click()

d2infile = "waehrung": d2insub = "Command3_Click"
MousePointer = 11
DoEvents
fn$ = form1.getusersetting("Kursexport") + "\kurse.dump"
On Error Resume Next
Kill fn$
On Error GoTo 0
Call form1.pg_xp("kurse", fn$)
MousePointer = 0

End Sub

Private Sub Command4_Click()
d2infile = "waehrung": d2insub = "Command4_Click"
Unload waehrung

End Sub

Private Sub Command6_Click()
d2infile = "waehrung": d2insub = "Command6_Click"
Load kurse
On Error Resume Next
Call kurse.SetFocus
On Error GoTo 0

End Sub

Private Sub Form_Load()
d2infile = "waehrung": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

Command1.Enabled = False
'dbpara$ = form1.getconnstr()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
'End If
Show
Call rlist1
Call rlist3
End Sub

Private Sub Form_Resize()
d2infile = "waehrung": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "waehrung": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Private Sub List1_Click()
Dim r As Recordset, id$

d2infile = "waehrung": d2insub = "List1_Click"
Command1.Enabled = False
List3.ListIndex = -1
w0% = List1.ListIndex
If w0% < 0 Then Exit Sub

w1$ = List1.List(w0%)
id$ = Left$(w1$, InStr(w1$, ",") - 1)
List2.Clear
cmd$ = "select * from kurse where wid='" + id$ + "' order by id desc"
Set w = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not w.EOF
  wk$ = trm(w!kurs)
  If wk$ <> "" And Left(wk$, 4) <> "HTTP" Then
    List2.AddItem w!id + ":" & wk$ & " je " + w!einheit
  Else
    Call form1.sqlqry("delete from kurse where id='" & w!id & "';")
  End If
  w.MoveNext
Wend

End Sub

Private Sub List3_Click()
Dim fn$

d2infile = "waehrung": d2insub = "List3_Click"
If List3.ListIndex >= 0 Then
List2.Clear
If LCase(List3.List(List3.ListIndex)) <> "kurse.txt" Then
fn$ = form1.getusersetting("Waehrungskurse") + "\" + List3.List(List3.ListIndex)
o% = FreeFile
Open fn$ For Input As #o%
While Not EOF(o%)
  Line Input #o%, lx$
  While lx$ <> ""
    p% = InStr(lx$, Chr$(10))
    If p% > 0 Then
      l$ = trm(Left$(lx$, p% - 1))
      lx$ = trm(Mid$(lx$, p% + 1))
    Else
      l$ = lx$
      lx$ = ""
    End If
    l$ = strrepl(l$, Chr$(13), "")
    If InStr(l$, "n/a") = 0 And InStr(l$, "ERROR") = 0 Then
      For i% = 0 To List2.ListCount - 1
        If Left$(l$, InStr(l$, ":")) = Left$(List2.List(i%), InStr(List2.List(i%), ":")) Then
          List2.RemoveItem i%
          DoEvents
          i% = List2.ListCount
        End If
      Next i%
      List2.AddItem l$
    End If
  Wend
Wend
Close #o%
End If
If List2.ListCount > 0 Then Command1.Enabled = True
End If
End Sub
