VERSION 5.00
Begin VB.Form q4 
   Caption         =   "Form2"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   ScaleHeight     =   6300
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check1 
      Caption         =   "Hilfe"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Spielvorschlag des Rechners"
      Top             =   5520
      Width           =   615
   End
   Begin VB.Timer w4ply 
      Enabled         =   0   'False
      Left            =   0
      Top             =   1440
   End
   Begin VB.ListBox gespielt 
      Height          =   1425
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   4920
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+ Compi"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Computerspieler hinzufügen"
      Top             =   5040
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cheftaste"
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
      Left            =   8520
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+ ich"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Computerspieler hinzufügen"
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command21 
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
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "neues Siel erstellen"
      Top             =   4440
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "q4.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Spieler"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image q4p 
      Height          =   5415
      Left            =   2400
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label ncrds 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "q4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currkid$, pcode, kq_ncrds As Integer, totalcards As Integer
Dim ltab$(19), mw(19) As Double, w4init As Boolean

Sub rlist1ingame()
Dim c$, r As Recordset, iam$

List1.Clear
iam$ = form1.getuserid()

c$ = "select * from q4dek where gname='" & Combo1.Text & "' and player='" & iam$ & "' order by pos"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  List1.AddItem r!card
  r.MoveNext
Wend
kq_ncrds = List1.ListCount
ncrds.Caption = Trim(kq_ncrds) & " Karten"
If List1.ListCount = 0 Then
  Label2.Caption = "Sie sind draussen"
  w4ply.Enabled = False
  Combo1.Text = ""
  Command5.Caption = "beenden"
End If
If kq_ncrds = totalcards Then
  MsgBox "Sie haben gewonnen."
  Command5.Caption = "beenden"
  Call Command5_Click
End If
c$ = "select * from q4gms where gname='" & Combo1.Text & "' and player='" & iam$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  On Error Resume Next
  Label2.Caption = "am Spiel: " & r!caller
  On Error GoTo 0
  If r!caller = r!player Then
    If trm(r!spiel) <> "" Then
      On Error Resume Next
      Label2.Caption = "Ihr Spiel: " & r!spiel
      On Error GoTo 0
      If w4init Then Call w4plyinit
    End If
  Else
   If w4init Then Call w4plyinit
  End If
End If
If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub

Sub w4plyinit()
'gespielt.Top = List2.Top
'gespielt.Left = List2.Left
'gespielt.Width = List2.Width
'gespielt.Height = List2.Height
gespielt.Visible = True
'List2.Visible = False
Call w4ply_Timer
w4ply.Interval = 1000
w4ply.Enabled = True
End Sub

Private Sub Check1_Click()
Call form1.setmylastFormVar(Me.name, "hilfe", Trim(Check1.value))
End Sub

Private Sub combo1_Change()
Call rlist3
End Sub
Sub rlist3()
Dim r As Recordset, c$
Dim bf As Boolean, bs As String

bf = False
List3.Clear
Command5.Caption = ""
c$ = "SELECT * FROM q4gms where gname='" & Combo1.Text & "';"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If bf = False Then
    bf = True
    bs = r!Status
  End If
  List3.AddItem r!player
  r.MoveNext
Wend
If bs = "neu" Then
  Command5.Caption = "austeilen"
  Timer1.Interval = 2000
  Timer1.Enabled = True
  Command2.Enabled = True
  Command4.Enabled = True
End If
If bs = "closed" Then
  Timer1.Enabled = False
  Command5.Caption = "beenden"
  Command2.Enabled = False
  Command4.Enabled = False
  Call rlist1ingame
End If

End Sub

Private Sub Combo1_Click()
gespielt.Visible = False
Call combo1_Change
End Sub

Private Sub Combo1_DropDown()
Dim r As Recordset, c$

c$ = "SELECT gname FROM q4gms order by gname ;"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
Combo1.Clear
c$ = ""
While Not r.EOF
  If c$ <> r!gname Then
    Combo1.AddItem r!gname
    c$ = r!gname
  End If
  r.MoveNext
Wend

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Dim i%, iam$, c$, r As Recordset, gid$

iam$ = form1.getuserid()
gid$ = Combo1.Text

c$ = "select * from q4gms where gname='" & gid$ & "' and player='" & iam$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then Exit Sub
c$ = "insert into q4gms (id,gname,player,status) values('" & _
           form1.newid("q4gms", "id", 20) & "','" & gid$ & "','" & iam$ & "','neu')"
Call form1.sqlqry(c$)
Command2.Enabled = False
Call rlist3
End Sub

Private Sub Command21_Click()
Dim i%, iam$, c$, r As Recordset, gid$

iam$ = form1.getuserid()
gid$ = iam$ & " " & Date & " " & Time
gespielt.Visible = False
c$ = "select * from q4gms where gname='" & gid$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  Command2.Enabled = False
  Exit Sub
End If
Command2.Enabled = False
c$ = "insert into q4gms (id,gname,player,status) values('" & _
           form1.newid("q4gms", "id", 20) & "','" & gid$ & "','" & iam$ & "','neu')"
Call form1.sqlqry(c$)
Combo1.Text = gid$
Command2.Enabled = False
Call rlist1

End Sub

Private Sub Command3_Click()
Dim c$
c$ = "delete from q4dek": Call form1.sqlqry(c$)
c$ = "delete from q4gms": Call form1.sqlqry(c$)
Call Command1_Click
End Sub

Private Sub Command4_Click()
Dim i%, iam$, c$, r As Recordset, pln%, gid$

pln% = 0

Do

pln% = pln% + 1
iam$ = "Compi-" & Trim(pln%)
gid$ = Combo1.Text
If gid$ = "" Then Exit Sub

c$ = "select * from q4gms where gname='" & gid$ & "' and player='" & iam$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If r.EOF Then
c$ = "insert into q4gms (id,gname,player,status) values('" & _
           form1.newid("q4gms", "id", 20) & "','" & gid$ & "','" & iam$ & "','neu')"
  Call form1.sqlqry(c$)
  Exit Do
End If

Loop
Call rlist3

End Sub

Private Sub Command5_Click()
Dim c$, i%, p%, pos As Integer, crd$, cl$, iam$, s As Recordset, m0 As Double, m1 As Double
Dim mt As Double
If Command5.Caption = "austeilen" Then
  totalcards = List1.ListCount
  p% = 0: pos = 0
  Command5.Caption = ""
  Timer1.Enabled = False
  iam$ = form1.getuserid()
  For i% = 0 To List3.ListCount - 1
    If List3.List(i%) = iam$ Then
      cl$ = List3.List(i%)
      Exit For
    End If
  Next i%
  c$ = "update q4gms set status='closed',caller='" & Trim(cl$) & "' where gname='" & Combo1.Text & "'"
  Call form1.sqlqry(c$)
  While List1.ListCount > 0
    pos = pos + 1
    i% = Int(Rnd * List1.ListCount)
    c$ = "insert into q4dek (id,gname,player,card,pos) values('" & _
        form1.newid("q4dek", "id", 20) & "','" & _
        Combo1.Text & "','" & _
        List3.List(p%) & "','" & _
        List1.List(i%) & "'," & Trim(pos) & ")"
    Call form1.sqlqry(c$)
    p% = p% + 1
    If p% >= List3.ListCount Then p% = 0
    List1.RemoveItem i%
  Wend
  Call rlist1ingame
End If
If Command5.Caption = "beenden" Then
  c$ = "delete from q4dek where gname='" & Combo1.Text & "'"
  Call form1.sqlqry(c$)
  c$ = "delete from q4gms where gname='" & Combo1.Text & "'"
  Call form1.sqlqry(c$)
  Combo1.Text = ""
  Label2.Caption = ""
  gespielt.Visible = False
  Combo1.Text = ""
  Call rlist1
End If
Call rlist3

End Sub

Private Sub Form_Load()
Dim i%, c$, r As Recordset, M As Double, mt As Double, m0 As Double, m1 As Double
Dim s As Recordset, klrv%

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Caption = "Quartett"
w4init = True
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
'dbpara$ = form1.getconnstr()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
'End If
Show
Call rlist1
totalcards = List1.ListCount
klrv% = Val(form1.mylastFormVar(Me.name, "hilfe", "0"))
If klrv% <> 0 Then klrv% = 1
Check1.value = klrv%

  For i% = 0 To 19: ltab$(i%) = "": Next i%
  c$ = "SELECT feldname,Felddaten From auftritthigru where auftrittstyp='Kiosk' and feldname<>'offen-von-bis' order by feldname"
  Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  c$ = ""
  m0 = -1
  i% = 0
  While Not s.EOF
    Debug.Print s!feldname; " "; s!felddaten
    If c$ <> s!feldname Then
      If m0 >= 0 Then
        Select Case LCase(c$)
          Case "herrie05":
            ltab$(i%) = c$ & ":" & m1 & "|" & m0
          Case Else
            ltab$(i%) = c$ & ":" & m0 & "|" & m1
        End Select
        i% = i% + 1
      End If
      c$ = s!feldname
      m0 = CDbl(s!felddaten)
      m1 = m0
    Else
      mt = CDbl(s!felddaten)
      If mt < m0 Then m0 = mt
      If mt > m1 Then m1 = mt
    End If
    s.MoveNext
  Wend
  If i% > 0 Then ltab$(i%) = c$ & ":" & m0 & "|" & m1
End Sub
Sub rlist1()
Dim c$, r As Recordset

List1.Clear
c$ = "select * from adresse where id like 'KQ-*';"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  List1.AddItem r!id & ":" & r!name
  r.MoveNext
Wend
kq_ncrds = List1.ListCount
ncrds.Caption = Trim(kq_ncrds) & " Karten"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Hide

End Sub


Private Sub List1_Click()
Dim i%, id$, dn$, fn$, sfddat$, sfwdat$, z As Integer, cmd$, ksel$, r As Recordset
Dim s As Recordset, iam$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
List2.Clear

id$ = List1.List(i%)
id$ = Left$(id$, InStr(id$, ":") - 1)
dn$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(id$)
fn$ = Mid$(id$, InStr(id$, "-") + 1) & ".bmp"
q4p.Picture = LoadPicture(dn$ & "\" & fn$)
cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='Kiosk' ORDER BY typ, position"
Set r = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  sfddat$ = "": sfwdat$ = ""
  z = r!zeilen
  sfwdat$ = r!feldname
  ksel$ = ""
  cmd$ = "SELECT id,Felddaten From auftritthigru where auftrittstyp='Kiosk' and auftrittsid='" + id$ + ksel$ + "' and feldname='" + r!feldname + "'"
  Set s = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
  If Not s.EOF Then
    If Not IsNull(s!felddaten) Then sfddat$ = s!felddaten
    List2.AddItem r!feldname & ": " & sfddat$
  End If
  DoEvents
  r.MoveNext
Wend
r.Close

If Command5.Caption = "beenden" Then
  iam$ = form1.getuserid()
  If List1.ListCount > 3 Then
    List1.ListIndex = 0
  End If
End If
End Sub

Private Sub List1_DblClick()
Dim i%, id$

If List1.ListCount <> 32 Then Exit Sub
i% = List1.ListIndex
If i% < 0 Then Exit Sub
List2.Clear

id$ = List1.List(i%)
id$ = Left$(id$, InStr(id$, ":") - 1)
Load shwAdrDetail
Call shwAdrDetail.savecheck
Call shwAdrDetail.refreshadrdetail(id$, "")
On Error Resume Next
shwAdrDetail.Combo3.Text = id$
Call shwAdrDetail.SetFocus
shwAdrDetail.srchit% = 1
On Error GoTo 0
End Sub

Private Sub List2_DblClick()
Dim i%, iam$, isdran$, c$, r As Recordset, kid$, atyp$, s As Recordset
Dim j%, fd$, l2b$

i% = List2.ListIndex
If i% < 0 Then Exit Sub
isdran$ = Label2.Caption
If InStr(isdran$, ": ") = 0 Then Exit Sub
isdran$ = Trim(Mid$(isdran$, InStr(isdran$, ": ") + 1))
iam$ = form1.getuserid()
If Command5.Caption = "beenden" Then
  If iam$ = isdran$ Then
'    MsgBox List2.List(i%)
    l2b$ = List2.List(i%): j% = InStr(l2b$, " (")
    If j% > 0 Then l2b$ = Trim(Left(l2b$, j% - 1))
    c$ = "update q4gms set spiel='" & l2b$ & "' where gname='" & Combo1.Text & "' " & _
       "and player='" & iam$ & "'"
    Call form1.sqlqry(c$)
    Label2.Caption = "Ihr Spiel: " & l2b$
    Call w4plyinit
    DoEvents
    For j% = 0 To List3.ListCount - 1
      If Left$(List3.List(j%), 6) = "Compi-" Then
        c$ = "select top 1 * from q4dek where gname='" & Combo1.Text & "' and player='" & List3.List(j%) & "' order by pos"
        Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
        If Not r.EOF Then
          atyp$ = List2.List(i%): atyp$ = Left(atyp$, InStr(atyp$, ":") - 1)
          kid$ = r!card: kid$ = Left$(kid$, InStr(kid$, ":") - 1)
          c$ = "SELECT Felddaten From auftritthigru where feldname='" & atyp$ & "' and auftrittsid='" & kid$ & "' and auftrittstyp='Kiosk'"
          Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
          fd$ = "-"
          If Not s.EOF Then fd$ = trm(s!felddaten)
          c$ = "update q4gms set spiel='" & atyp$ & ": " & fd$ & "' where gname='" & Combo1.Text & "' " & _
               "and player='" & List3.List(j%) & "'"
          Call form1.sqlqry(c$)
          DoEvents
        Else
          'MsgBox List3.List(j%) & " ist draussen."
          c$ = "delete from q4gms where player='" & List3.List(j%) & "' and gname='" & Combo1.Text & "'"
          Call form1.sqlqry(c$)
          List3.RemoveItem j%
          j% = j% - 1
        End If
      End If
    Next j%
  Else
    MsgBox "Sie sind nicht dran!" & vbCrLf & Label2.Caption
  End If
End If

End Sub

Private Sub List3_Click()
Dim i%, iam$

i% = List3.ListIndex
If i% < 0 Then Exit Sub
iam$ = form1.getuserid()
If Command5.Caption = "beenden" Then
  If iam$ <> List3.List(i%) Then
    For i% = 0 To List3.ListCount - 1
      If iam$ = List3.List(i%) Then
        List3.ListIndex = i%
        Exit Sub
      End If
    Next i%
  End If
End If
End Sub

Private Sub Timer1_Timer()
Call rlist3
End Sub

Private Sub w4ply_Timer()
Dim c$, r As Recordset, iam$, sp$, spw$, spw0$, sp0$, p%, bpl As Boolean
Dim i%, habschon As Boolean, alleda As Boolean, alleok As Boolean, winner$
Dim j%, cpl$, s As Recordset, rcd$, m0 As Double, m1 As Double, mi%, M As Double, mt As Double
Dim atyp$, kid$, fd$, newmov As Boolean, l2b$


iam$ = form1.getuserid()

c$ = "select * from q4gms where gname='" & Combo1.Text & "' and player='" & iam$ & "' and status='siehmich'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  'MsgBox "Rundenende"
  c$ = "update q4gms set status='gesehen' where ((gname='" & Combo1.Text & "') and ((player='" & iam$ & "') or (player like 'Compi-*')))"
  Call form1.sqlqry(c$)
  Exit Sub
End If
c$ = "select * from q4gms where gname='" & Combo1.Text & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
gespielt.Clear
alleda = True
alleok = True
bpl = False
habschon = False
While Not r.EOF
  sp$ = trm(r!spiel)
  p% = InStr(sp$, ":")
  If p% > 0 Then
    spw$ = Trim(Mid$(sp$, p% + 1))
    spw0$ = spw0$
    sp$ = Left(sp$, p% - 1)
    sp0$ = sp$
    bpl = True
    If iam$ = r!player Then habschon = True
  Else
    spw$ = "": sp$ = ""
    alleda = False
  End If
  gespielt.AddItem r!player & " " & spw$ & " " & sp$
  If r!Status <> "gesehen" Then alleok = False
  r.MoveNext
Wend
If Not bpl Then
  c$ = "select * from q4gms where gname='" & Combo1.Text & "'"
  Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If Not r.EOF Then
    If Label2.Caption <> "am Spiel: " & r!caller Then
      Label2.Caption = "am Spiel: " & r!caller
      w4init = False
      Call rlist1ingame
      w4init = True
    End If
  End If
End If
If Not habschon And List2.ListIndex < 0 Then
  If Check1.value = 1 Then Call helper
End If
If bpl And Not habschon Then
  For i% = 0 To List2.ListCount - 1
    If InStr(List2.List(i%), sp0$) = 1 Then
      List2.ListIndex = i%
    l2b$ = List2.List(i%): j% = InStr(l2b$, " (")
    If j% > 0 Then l2b$ = Trim(Left(l2b$, j% - 1))
      c$ = "update q4gms set spiel='" & l2b$ & "' where gname='" & Combo1.Text & "' " & _
         "and player='" & iam$ & "'"
      Call form1.sqlqry(c$)
      Exit For
    End If
  Next i%
End If
If alleda Then
  Call enderunde
End If
If alleok And Label2.Caption <> "" Then
  gespielt.Visible = False
  For i% = 0 To List3.ListCount - 1
    If Left$(List3.List(i%), 6) <> "Compi-" Then
      If Left$(List3.List(i%), 6) <> iam$ Then
        Exit Sub
      Else
        winner$ = word1(Label2.Caption)
        c$ = "update q4gms set caller='" & winner$ & "',spiel='',status='closed' where gname='" & Combo1.Text & "'"
        Call form1.sqlqry(c$)
        Label2.Caption = ""
        Exit For
      End If
    End If
  Next i%
  Call rlist1ingame
End If
mi% = -1
If InStr(Label2.Caption, "am Spiel: Compi-") = 1 Then
  For i% = 0 To List3.ListCount - 1
    If Left$(List3.List(i%), 6) <> "Compi-" Then
      If Left$(List3.List(i%), 6) <> iam$ Then
        Exit Sub
      Else
        cpl$ = Label2.Caption
        cpl$ = Trim(Mid$(cpl$, InStr(cpl$, ":") + 1))
        c$ = "select top 1 * from q4dek where gname='" & Combo1.Text & "' and player='" & cpl$ & "' order by pos"
        Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
        If r.EOF Then
          'MsgBox cpl$ & " ist draussen."
          c$ = "delete from q4gms where player='" & cpl$ & "' and gname='" & Combo1.Text & "'"
          Call form1.sqlqry(c$)
          For j% = 0 To List3.ListCount - 1
            If List3.List(j%) = cpl$ Then
              List3.RemoveItem j%
              Exit For
            End If
          Next j%
        Else
          For j% = 0 To 19: mw(j%) = 0: Next j%
          rcd$ = r!card: rcd$ = Left(rcd$, InStr(rcd$, ":") - 1)
          c$ = "SELECT * From auftritthigru where auftrittstyp='Kiosk' and auftrittsid='" & rcd$ & "'"
          Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
          While Not s.EOF
            For j% = 0 To 19
              If ltab$(j%) = "" Then Exit For
              If InStr(LCase(ltab$(j%)), LCase(s!feldname)) = 1 Then
                mw(j%) = CDbl(s!felddaten)
                Exit For
              End If
            Next j%
            s.MoveNext
          Wend
          M = 0: mi% = 0
          For j% = 0 To 19
            If ltab$(j%) = "" Then Exit For
            c$ = ltab$(j%)
            c$ = Mid$(c$, InStr(c$, ":") + 1): m0 = CDbl(Left(c$, InStr(c$, "|") - 1))
            c$ = Mid$(c$, InStr(c$, "|") + 1): m1 = CDbl(c$)
            mt = (mw(j%) - m0) / (m1 - m0)
            If mt > M Then
              M = mt
              mi% = j%
            End If
          Next j%
          atyp$ = ltab$(mi%): atyp$ = Left(atyp$, InStr(atyp$, ":") - 1)
          c$ = "update q4gms set spiel='" & atyp$ & ": " & Trim(mw(mi%)) & "' where gname='" & Combo1.Text & "' " & _
               "and player='" & cpl$ & "'"
          Call form1.sqlqry(c$)
          Label2.Caption = cpl$ & ": " & atyp$ & ": " & Trim(mw(mi%))
          Exit For
        End If
      End If
    End If
  Next i%
  For j% = 0 To List3.ListCount - 1
    If j% <> i% And Left$(List3.List(j%), 6) = "Compi-" Then
      c$ = "select top 1 * from q4dek where gname='" & Combo1.Text & "' and player='" & List3.List(j%) & "' order by pos"
      Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
      If Not r.EOF Then
        kid$ = r!card: kid$ = Left$(kid$, InStr(kid$, ":") - 1)
        c$ = "SELECT Felddaten From auftritthigru where feldname='" & atyp$ & "' and auftrittsid='" & kid$ & "' and auftrittstyp='Kiosk'"
        Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
        fd$ = "-"
        If Not s.EOF Then fd$ = trm(s!felddaten)
        c$ = "update q4gms set spiel='" & atyp$ & ": " & fd$ & "' where gname='" & Combo1.Text & "' " & _
             "and player='" & List3.List(j%) & "'"
        Call form1.sqlqry(c$)
        DoEvents
      Else
        'MsgBox List3.List(j%) & " ist draussen."
        c$ = "delete from q4gms where player='" & List3.List(j%) & "' and gname='" & Combo1.Text & "'"
        Call form1.sqlqry(c$)
        List3.RemoveItem j%
        j% = j% - 1
      End If
    End If
  Next j%
End If
End Sub
Sub enderunde()
Dim c$, r As Recordset, tobeat As Double, ibeat As Double, winner$, fld$, ansp$
Dim iam$, ccount%, i%, rrr

iam$ = form1.getuserid()
c$ = "SELECT * FROM q4gms where gname='" & Combo1.Text & "' and player='" & iam$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  If r!Status = "gesehen" Then Exit Sub
End If
For i% = 0 To List3.ListCount - 1
  If Left$(List3.List(i%), 6) <> "Compi-" Then
    If Left$(List3.List(i%), 6) <> iam$ Then
      Exit Sub
    Else
      Exit For
    End If
  End If
Next i%
c$ = "SELECT * FROM q4gms where gname='" & Combo1.Text & "' and player=caller;"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  winner$ = r!spiel
  fld$ = word1(winner$): fld$ = Left(fld$, Len(fld$) - 1)
  winner$ = Trim(Mid$(winner$, InStr(winner$, ":") + 1))
  tobeat = CDbl(word1(winner$))
  winner$ = r!caller
End If
c$ = "SELECT * FROM q4gms where gname='" & Combo1.Text & "';"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If r!caller <> r!player Then
    ansp$ = r!spiel
    ansp$ = Trim(Mid$(ansp$, InStr(ansp$, ":") + 1))
    On Error Resume Next
    ibeat = CDbl(ansp$)
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then ibeat = 0
    If LCase(fld$) = "herrie05" Then
      If ibeat < tobeat Then
        winner$ = r!player
        tobeat = ibeat
      End If
    Else
      If ibeat > tobeat Then
        winner$ = r!player
        tobeat = ibeat
      End If
    End If
  End If
  r.MoveNext
Wend
'MsgBox winner$ & " gewinnt mit " & tobeat
Label2.Caption = winner$ & " gewinnt diese Runde"
c$ = "update q4gms set status='siehmich' where gname='" & Combo1.Text & "';"
Call form1.sqlqry(c$)
ccount% = 10000
For i% = 0 To List3.ListCount - 1
  c$ = "select top 1 id from q4dek where gname='" & Combo1.Text & "' and player='" & List3.List(i%) & "' order by pos"
  Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If Not r.EOF Then
    c$ = "update q4dek set player='" & winner$ & "',pos=" & ccount% & " where id='" & r!id & "'"
    Call form1.sqlqry(c$)
  End If
Next i%
ccount% = 1
c$ = "select id from q4dek where gname='" & Combo1.Text & "' and player='" & winner$ & "' order by pos"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  c$ = "update q4dek set pos=" & ccount% & " where id='" & r!id & "'"
  ccount% = ccount% + 1
  Call form1.sqlqry(c$)
  r.MoveNext
Wend
End Sub
Function itsmyturn() As Boolean
Dim iam$, c$, r As Recordset

itsmyturn = False
iam$ = form1.getuserid()
c$ = "SELECT * FROM q4gms where gname='" & Combo1.Text & "' and player='" & iam$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  If r!player = r!caller Then itsmyturn = True
End If

End Function
Sub helper()
Dim j%, c$, r As Recordset, s As Recordset, mt As Double, mi%, M As Double
Dim atyp$, rcd$, m0 As Double, m1 As Double, iam$, lprb As Double
Dim xm1 As Long, xm As Long, xR As Recordset, sorder$, k%
Dim myturn As Boolean

iam$ = form1.getuserid()

          For j% = 0 To 19: mw(j%) = 0: Next j%
          rcd$ = List1.List(0): rcd$ = Left(rcd$, InStr(rcd$, ":") - 1)
          c$ = "SELECT * From auftritthigru where auftrittstyp='Kiosk' and auftrittsid='" & rcd$ & "'"
          Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
          mt = 1.1
          myturn = itsmyturn()
          While Not s.EOF
            For j% = 0 To 19
              If ltab$(j%) = "" Then Exit For
              If InStr(LCase(ltab$(j%)), LCase(s!feldname)) = 1 Then
                mw(j%) = CDbl(s!felddaten)
                sorder$ = ">"
                If LCase(s!feldname) = "herrie05" Then
                  sorder$ = "<"
                End If
                If Not myturn Then sorder$ = sorder$ & "="
                c$ = "SELECT count(*) as wert " & _
                     "FROM auftritthigru, q4gms INNER JOIN q4dek ON (q4gms.player = q4dek.player) AND (q4gms.gname = q4dek.gname) " & _
                     "WHERE (((q4gms.player)<>'" & iam$ & "') AND ((auftritthigru.auftrittsid)=Left(card,InStr(card,':')-1)) " & _
                           "AND ((auftritthigru.FeldName)='" & s!feldname & "') AND (q4gms.gname ='" & Combo1.Text & "') and (CDbl(auftritthigru.Felddaten)" & sorder$ & d2db(s!felddaten) & ") ) "
                Set xR = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
                xm1 = xR!wert
                c$ = "SELECT count(*) as wert " & _
                     "FROM auftritthigru, q4gms INNER JOIN q4dek ON (q4gms.player = q4dek.player) AND (q4gms.gname = q4dek.gname) " & _
                     "WHERE (((q4gms.player)<>'" & iam$ & "') AND ((auftritthigru.auftrittsid)=Left(card,InStr(card,':')-1)) " & _
                           "AND ((auftritthigru.FeldName)='" & s!feldname & "') AND (q4gms.gname ='" & Combo1.Text & "')  ) "
                Set xR = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
                xm = xR!wert
                If xm = 0 Then Exit Sub
                M = xm1 / xm
                Debug.Print s!feldname; "="; s!felddaten; "==>"; M
          For k% = 0 To List2.ListCount - 1
            If InStr(LCase(List2.List(k%)), LCase(s!feldname)) = 1 Then
              lprb = M * (List3.ListCount - 1): If lprb > 1 Then lprb = 1
              List2.List(k%) = List2.List(k%) & " (" & Int((1# - lprb) * 100) & "%)"
              Exit For
            End If
          Next k%
                If M < mt Then
                  mt = M
                  mi% = j%
                End If
'                Exit For
              End If
            Next j%
            s.MoveNext
          Wend

          lprb = mt * (List3.ListCount - 1): If lprb > 1 Then lprb = 1
          Me.Caption = "Gewinnwahrscheinlichkeit=" & Int((1# - lprb) * 100) & "%"
          atyp$ = ltab$(mi%): atyp$ = Left(atyp$, InStr(atyp$, ":") - 1)
          For j% = 0 To List2.ListCount - 1
            If InStr(LCase(List2.List(j%)), LCase(atyp$)) = 1 Then
              List2.ListIndex = j%
              Exit For
            End If
          Next j%
End Sub
