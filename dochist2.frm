VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form dochist2 
   Caption         =   "Kontakthisthorie"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton csvxfelder 
      Caption         =   "Fields"
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton csvx 
      Caption         =   "CSV"
      Height          =   495
      Left            =   1200
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btnPrj 
      Caption         =   "--> Projekt"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10680
      TabIndex        =   34
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox msonly 
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   9960
      MaskColor       =   &H00000000&
      Picture         =   "dochist2.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   31
      ToolTipText     =   "save topic"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   12240
      Picture         =   "dochist2.frx":0672
      Style           =   1  'Grafisch
      TabIndex        =   30
      ToolTipText     =   "move to archive"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   11880
      Picture         =   "dochist2.frx":07FC
      Style           =   1  'Grafisch
      TabIndex        =   29
      ToolTipText     =   "open archive"
      Top             =   3240
      Width           =   375
   End
   Begin VB.ListBox remlist 
      Height          =   855
      IntegralHeight  =   0   'False
      Left            =   9840
      TabIndex        =   28
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ListBox adrlinks 
      Height          =   855
      IntegralHeight  =   0   'False
      Left            =   9840
      TabIndex        =   26
      ToolTipText     =   "select and press delete to remove an entry"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
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
      Height          =   270
      Left            =   10440
      TabIndex        =   27
      ToolTipText     =   "Link to another address"
      Top             =   3360
      Width           =   255
   End
   Begin VB.ListBox topics 
      Height          =   1935
      IntegralHeight  =   0   'False
      Left            =   9840
      TabIndex        =   16
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox tfilter 
      Height          =   285
      Left            =   10680
      TabIndex        =   24
      ToolTipText     =   "Suchworter ( keine Umlaute )"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Datum / User"
      Height          =   255
      Left            =   10560
      TabIndex        =   23
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command42 
      Height          =   375
      Left            =   12000
      Picture         =   "dochist2.frx":0E26
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Sort"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   12120
      Picture         =   "dochist2.frx":1198
      Style           =   1  'Grafisch
      TabIndex        =   20
      ToolTipText     =   "Wiedervorlage"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Abwahl"
      Height          =   255
      Left            =   10560
      TabIndex        =   19
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox topic 
      Height          =   4335
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.CommandButton Command11 
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
      Left            =   12240
      Picture         =   "dochist2.frx":1804
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "New Topic"
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox sortlist 
      Height          =   2400
      IntegralHeight  =   0   'False
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox suchw 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      ToolTipText     =   "Suchworter ( keine Umlaute )"
      Top             =   4800
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   4440
   End
   Begin VB.TextBox shmax 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "20"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command18 
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
      TabIndex        =   8
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "no access"
      Height          =   735
      Left            =   1680
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "dochist2.frx":1B96
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton wvl 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   960
      Picture         =   "dochist2.frx":1DE6
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Wiedervorlage"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Antworten"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Weiterleiten"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
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
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "mailsafe only"
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
      Left            =   360
      TabIndex        =   32
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Linked:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      ToolTipText     =   "In Kontakten suchen"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label zleft 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10320
      TabIndex        =   22
      ToolTipText     =   "In Kontakten suchen"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Suche:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "In Kontakten suchen"
      Top             =   4815
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Treffer:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      ToolTipText     =   "In Kontakten suchen"
      Top             =   4455
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Topics"
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
      Left            =   9840
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "dochist2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currkid$, dhrepl$
Dim currvid$, currtopic As String
Dim tm_brk%, shm As Integer, runmode As String
Dim msnoupd As Boolean, shm_igno As Boolean, aktwerk$
Dim critter As String

Private Sub adrlinks_DblClick()
Dim i%, sid$

i% = adrlinks.ListIndex
If i% < 0 Then Exit Sub

sid$ = adrlinks.List(i%)
Load shwAdrDetail
Call shwAdrDetail.savecheck
Call shwAdrDetail.refreshadrdetail(sid$, "")

End Sub

Private Sub adrlinks_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, j%, c$

If KeyCode = 8 Or KeyCode = 46 Then
  i% = adrlinks.ListIndex
  If i% < 0 Then Exit Sub
  j% = topics.ListIndex
  If j% < 0 Then Exit Sub

  c$ = "delete from sysvars where owner='sysvar_system_tlnk_" + topics.List(j%) + "_" + adrlinks.List(i%) + "'"
  Call form1.sqlqry(c$)
  If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "dochist2", "adrlinks_KeyDown")
  adrlinks.RemoveItem i%
End If
End Sub

Private Sub btnPrj_Click()
Dim i%, c$
Dim r As ADODB.Recordset, l$, neuid As String, altid As String, n$

  i% = topics.ListIndex
  If i% < 0 Then Exit Sub
  c$ = topics.List(i%)

  If btnPrj.Caption = "rename topic" Then
    altid = c$
    neuid = InputBox(transe("neuer Name"), altid, altid)
    If trm(neuid) = "" Or neuid = altid Then Exit Sub
    MousePointer = 11: DoEvents
    c$ = "select id,Betreff,Nachricht from todolist where Betreff like '" + altid + " [%'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    While Not r.EOF
      l$ = strrepl(trm(r!betreff), altid, neuid)
      n$ = strrepl(trm(r!nachricht), altid, neuid)
'Debug.Print r!Owner; " ("; r!wert; ")"; vbCrLf; l$
      c$ = "update todolist set Betreff='" + l$ + "',Nachricht='" + n$ + "' where id='" + trm(r!id) + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
    c$ = "update opt_topics set topicid='" & neuid & "' where topicid='" & altid & "'"
    Call form1.sqlqry(c$)
    c$ = "select id,owner,wert from sysvars where owner like 'sysvar_system_tlnk_" + altid + "_%'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    While Not r.EOF
      l$ = strrepl(trm(r!Owner), altid, neuid)
'Debug.Print r!Owner; " ("; r!wert; ")"; vbCrLf; l$
      c$ = "update sysvars set owner='" + l$ + "' where id='" + trm(r!id) + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
    Call rtopics
    For i% = 0 To topics.ListCount - 1
      If topics.List(i%) = neuid Then
        topics.ListIndex = i%
        Exit For
      End If
    Next i%
    MousePointer = 0: DoEvents
  Else
    If btnPrj.Caption <> "dead link" Then
      If Len(c$) <> 0 Then
        Load tplan
        Call tplan.rlists
        Call tplan.nulldsp
        Call tplan.showrec(c$)
        On Error Resume Next
        Call tplan.SetFocus
        On Error GoTo 0
      End If
    Else
      l$ = "delete from opt_topics where topicid='" & c$ & "'"
      Call form1.sqlqry(l$)
      l$ = "delete from sysvars where owner like 'sysvar_system_tlnk_" + c$ + "_%'"
      Call form1.sqlqry(l$)
      If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(l$, "dochist2", "btnPrj_Click")
      Call rtopics
      If i% < topics.ListCount And i% >= 0 Then topics.ListIndex = i%
    End If
  End If

End Sub

Private Sub Command1_Click()
'd2infile = "dochist2": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command11_Click()
Dim neuid$, c$, i%
Dim id$, aid$, kid$, r As ADODB.Recordset

If currvid$ = "" Then
  Call notopics
  Exit Sub
End If
neuid$ = trmx1(InputBox(transe("Neues Thema:"), transe("Neues Notiz-Thema"), idn$))
If neuid$ = "" Then Exit Sub
For i% = 0 To topics.ListCount - 1
  If neuid$ = topics.List(i%) Then
    topics.ListIndex = i%
    Exit Sub
  End If
Next i%

c$ = "insert into opt_topics (id,vid,kid,topicid) values("
c$ = c$ + "'" + form1.newid("opt_topics", "id", 10) + "',"
c$ = c$ + "'" + currvid$ + "','-1',"
c$ = c$ + "'" + neuid$ + "')"
Call form1.sqlqry(c$)
Call rtopics
End Sub

Private Sub Command12_Click()
topics.ListIndex = -1
topic.Visible = False
Command25.Enabled = False
Command5.Enabled = False

End Sub

Private Sub Command18_Click()
'd2infile = "dochist2": d2insub = "Command18_Click"
Call form1.handbuchcall("06-Adressen.htm")
End Sub

Private Sub Command2_Click()
Dim id$, aid$, kid$, p%, r As ADODB.Recordset, em$, tgi$, rrr

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "Command2_Click"
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
MousePointer = 11: DoEvents
id$ = Mid$(id$, p% + 4)
cmd$ = "select docname,kontakt,adresse from dochist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 And Not r.EOF Then
  Load smtp
  smtp.xattach.Clear
  tgi$ = form1.dupcheck(r!docname)
  If Not nexist(tgi$) Then
    smtp.xattach.AddItem tgi$
    Call smtp.cmdAdd_Click
    smtp.xattach.Clear
  End If
End If
MousePointer = 0
End Sub

Private Sub Command25_Click()
Dim pos As Integer

Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), topics.List(topics.ListIndex) & " [Wiedervorlage] Adresse:" + _
               currvid$, "", Date, Left(Time, 5))
Call create2do.SetFocus
create2do.Text1(1).Enabled = False
create2do.Text1(3).Enabled = False
create2do.Text1(4).text = "TOPIC: " + topics.List(topics.ListIndex) + vbCrLf: pos = Len(create2do.Text1(4).text)
c$ = transe("Ändern Sie nicht die erste Zeile.")
create2do.Text1(4).text = create2do.Text1(4).text + c$
create2do.Text1(4).SelStart = pos
create2do.Text1(4).SelLength = Len(c$)
On Error Resume Next
Call create2do.Text1(4).SetFocus
On Error GoTo 0

End Sub

Private Sub Command29_Click()
Dim dts$, X

On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  dts$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(currvid$): MkDir dts$
  dts$ = dts$ + "\topics": MkDir dts$
  X = Shell("explorer.exe " + dts$, vbNormalFocus)
On Error GoTo 0
End Sub

Private Sub Command3_Click()
Dim id$, aid$, kid$, p%, r As ADODB.Recordset, em$, tgi$, o%

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "Command3_Click"
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
MousePointer = 11: DoEvents
id$ = Mid$(id$, p% + 4)
cmd$ = "select docname,kontakt,adresse,betreff from dochist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  aid$ = r!adresse
  kid$ = r!kontakt
  Load smtp
  If kid$ <> "" And kid$ <> "-1" Then
    em$ = form1.getkontaktemailbyid(kid$)
  Else
    em$ = form1.getemailbyid(aid$)
  End If
  Call smtp.callback(aid$, kid$, em$)
  tgi$ = form1.dupcheck(r!docname)
  o% = FreeFile
  Open tgi$ For Input As #o%
  hd% = 1
  brk% = 0
  smtp.txtMessageSubject = "AW: " & r!betreff
  While Not EOF(o%) And brk% = 0
    Line Input #o%, l$
    'Debug.Print l$
    If trm(l$) = "" And hd% = 1 Then
      l$ = form1.get_kontaktname_by_id(kid$)
      If l$ = "" Then
        l$ = form1.getnamebyid(aid$)
      End If
      l$ = vbCrLf & l$ & " schrieb:"
      hd% = 0
      dop% = 1
    End If
    If hd% = 0 And InStr(LCase(l$), "content-type:") = 1 Then
      dop% = 0
      While l$ <> ""
        If InStr(LCase(l$), "text/plain") Then dop% = 1
        'SMTP.txtMessageText = SMTP.txtMessageText & vbCrLf & "|    " & l$
        Line Input #o%, l$
      Wend
      If dop% = 0 Then
        l$ = "..."
        'brk% = 1
      End If
    End If
    If hd% = 0 And dop% = 1 Then
      smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "> " & l$
    End If
  Wend
  Close #o%
End If
MousePointer = 0

End Sub

Private Sub Command4_Click()
Dim i As Integer, lvitem

'd2infile = "dochist2": d2insub = "Command4_Click"
For i = 1 To gd1.ListItems.Count
  gd1.ListItems(i).Selected = False
Next i
On Error Resume Next
Call gd1.SetFocus
On Error GoTo 0
DoEvents
For i = 1 To gd1.ListItems.Count
  Set lvitem = gd1.ListItems(i)
  If lvitem.SubItems(5) = "no access" Then
    gd1.ListItems(i).Selected = True
  End If
Next i

End Sub

Private Sub Command5_Click()
If Right(topic.text, 2) <> vbCrLf Then topic.text = topic.text + vbCrLf
topic.text = topic.text + "____________" + vbCrLf + datum2sql(Date) + vbCrLf + form1.getuserid() + " - "
topic.SelLength = 0
topic.SelStart = Len(topic.text)
On Error Resume Next
Call topic.SetFocus
On Error GoTo 0
End Sub

Private Sub Command6_Click()
Dim c$, j%
Dim dts$, o%, i%

j% = topics.ListIndex
If j% < 0 Then Exit Sub
If adrlinks.ListCount < 0 Then Exit Sub
dts$ = ""
For i% = 0 To adrlinks.ListCount - 1
  c$ = adrlinks.List(i%)
  On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  dts$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(c$): MkDir dts$
  dts$ = dts$ + "\topics": MkDir dts$
  On Error GoTo 0
  o% = FreeFile
  Open dts$ + "\" + validatefn(trm(currtopic)) + ".txt" For Output As #o%
  Print #o%, topic.text
  Close #o%
  c$ = "delete from sysvars where owner='sysvar_system_tlnk_" + currtopic + "_" + adrlinks.List(i%) + "'"
  Call form1.sqlqry(c$)
  If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "dochist2", "Command6_Click")
Next i%
c$ = "delete from opt_topics where topicid='" + currtopic + "'"
Call form1.sqlqry(c$)
topics.RemoveItem j%
topic.text = "": currtopic = ""
While remlist.ListCount > 0
  remlist.ListIndex = 0
  id$ = remlist.List(i%)
  pos% = InStr(id$, "(ID:")
  If pos% > 0 Then
    id$ = Mid$(id$, pos% + 4)
    c$ = "delete from todolist where id='" + id$ + "'"
    Call form1.sqlqry(c$)
    DoEvents
  End If
  remlist.RemoveItem 0
Wend
If dts$ <> "" Then
End If
End Sub

Private Sub Command7_Click()
Call svctt
End Sub

Private Sub Command8_Click()
Dim s0$, c$, j%

j% = topics.ListIndex
If j% < 0 Then Exit Sub

  Load adrselect
  Call adrselect.sel_init("", transe("Person"))
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then
    s0$ = adrselect.sel_getselected()
    c$ = "insert into sysvars (id,owner,wert) values ('" + form1.newid("sysvars", "id", 35) + "',"
    c$ = c$ + "'sysvar_system_tlnk_" + topics.List(j%) + "_" + s0$ + "','" + s0$ + "')"
    Call form1.sqlqry(c$)
    adrlinks.AddItem s0$
  End If
  Unload adrselect
End Sub

Private Sub csvx_Click()
Dim j As Integer, i As Integer, aid As String, c$, diesefelder As String, w$, l$, d$
Dim r As ADODB.Recordset, s As ADODB.Recordset
Dim tabs$(9), tptr%, xport$(9), xld$

w$ = form1.getusersetting("workplayedwherefields", "Halle,Dirigent,Orchester,Solist|Solist1|Künstler")
xld$ = form1.getusersetting("exceldelimiter", ",")

MousePointer = 11: DoEvents
diesefelder = "": tptr% = 0
While w$ <> ""
  c$ = cut_d1(w$, ","): w$ = cut_d2bis(w$, ",")
  If InStr(c$, "|") > 0 Then
    l$ = ""
    While c$ <> ""
      d$ = cut_d1(c$, "|"): c$ = cut_d2bis(c$, "|")
      If diesefelder <> "" Then diesefelder = diesefelder + " or "
      diesefelder = diesefelder + "Feldname='" + d$ + "'"
      tabs$(tptr%) = tabs$(tptr%) + "|" + d$ + "|"
    Wend
    tabs$(tptr%) = strrepl(tabs$(tptr%), "||", "|")
    tptr% = tptr% + 1
  End If
  If c$ <> "" Then
    If diesefelder <> "" Then diesefelder = diesefelder + " or "
    diesefelder = diesefelder + "Feldname='" + c$ + "'"
    tabs$(tptr%) = c$
    tptr% = tptr% + 1
  End If
Wend
o% = FreeFile
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) = "" Then Exit Sub
Open fn$ For Output As #o%
Print #o%, """" + form1.getkompnamebywerkid(aktwerk$) & ": " & form1.getwerknamebyid(aktwerk$) + """"
Print #o%, """"; transe("Datum"); """" + xld$;
Print #o%, """"; transe("Ort"); """" + xld$;
For j% = 0 To tptr% - 1
  l$ = tabs$(j%)
  If InStr(tabs$(j%), "|") > 0 Then l$ = cut_d1(cut_d2bis(l$, "|"), "|")
  Print #o%, """"; transe(l$); """" + xld$;
Next j%
Print #o%,

For i = 1 To gd1.ListItems.Count
  aid = gd1.ListItems(i).SubItems(5)
  For j% = 0 To tptr% - 1: xport(j%) = "": Next j%
  c$ = "select datum,ort from auftritt where id='" + aid + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    While Not r.EOF
      c$ = "select FeldName,FeldDaten from auftritthigru where auftrittsid='" + aid + "' and (" + diesefelder + ")"
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
      rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
      If rrr = 0 Then
        While Not s.EOF
          For j% = 0 To tptr% - 1
            If tabs$(j%) = trm(s!feldname) Or InStr(tabs$(j%), "|" + trm(s!feldname) + "|") > 0 Then Exit For
          Next j%
          j% = imin(j%, tptr% - 1)
          If xport$(j) <> "" Then xport$(j) = xport$(j) + vbCrLf
          xport$(j%) = xport$(j%) + trm(s!felddaten)
'          Debug.Print trm(r!Datum) + ", " + trm(s!feldname) + ": " + trm(s!felddaten) + " (" + trm(j) + ")"
          s.MoveNext
        Wend
      End If
      Print #o%, """"; trm(r!datum); """" + xld$;
      Print #o%, """"; trm(r!ort); """" + xld$;
      For j% = 0 To tptr% - 1
'        Debug.Print """" + xport$(j%) + """" + ",";
        Print #o%, """"; xport$(j%); """" + xld$;
      Next j%
'      Debug.Print
      Print #o%,
      r.MoveNext
    Wend
  End If
Next i
Close #o%
X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)
MousePointer = 0
End Sub

Private Sub csvxfelder_Click()
Dim n$, wert$, warn$

  wert$ = form1.getusersetting("workplayedwherefields", "Halle,Dirigent,Orchester,Solist|Solist1|Künstler")
  warn$ = "A word of warning: 'Wrong' input will possibly crash" + vbCrLf + "the export - and the program"
  n$ = InputBox(transe("Neue Benutzereinstellung:") + vbCrLf + wert$ + vbCrLf + vbCrLf + warn$, "Fields to export", wert$)
  If n$ = "" Then Exit Sub
  Call form1.setusersetting("workplayedwherefields", n$)
End Sub

Private Sub Form_Load()

'd2infile = "dochist2": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Timer1.Enabled = False
tm_brk% = 0
shm_igno = True
runmode = "dochist"
btnPrj.Enabled = False
form1.dochistisopen = True
    gd1.View = lvwReport
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum"), 1600)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Benutzer"), 800)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Kontakt"), 1500)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Typ"), 1500)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Betreff"), 2900)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Dokument"), 1100)

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
shmax.text = form1.getusersetting("zzzdochistmax", "20")
shm_igno = False
Call form1.formpos(Me)
Command2.Visible = False
Command3.Visible = False
Command1.ToolTipText = transe("Formular schliessen")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command3.Caption = transe("&Antworten")
Command5.Caption = transe("Datum") + " / " + transe("Benutzer")
Command2.Caption = transe("Weiterleiten")
Command4.Visible = True
csvx.Visible = False
csvxfelder.Visible = False
btnPrj.Caption = "--> " + transe("Projekt")
Label4.Caption = transe("Suche")
Label3.Caption = transe("Max. Treffer:")
wvl.ToolTipText = transe("Wiedervorlage")
dhrepl$ = form1.getusersetting("dhreplace", "")
Me.BackColor = form1.cleancolor()
msonly.value = 0
If form1.getusersetting("ignoremailsfromdochist", "nein") = "ja" Then
  msnoupd = True
  msonly.value = 1
  msnoupd = False
End If
Show

If form1.isfieldmissing("opt_topics", "id") Then
  Call notopics
  Exit Sub
End If

End Sub

Private Sub notopics()
  
  topics.Clear
  topics.AddItem "Function not available."
  Command11.Enabled = False
  Command12.Enabled = False
  Command6.Enabled = False
  Command29.Enabled = False
  topics.Enabled = False
End Sub

Private Sub Form_Resize()
'd2infile = "dochist2": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "dochist2": d2insub = "Form_Unload"

Call savecheck
form1.dochistisopen = False
Hide
On Error GoTo exuld

Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub
Sub r1q(c$, d$, mskrit$)
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, tsts$, prvs$, tfn$, i%, dx$
Dim upd1$, upd2$, dct As String, uId$, fn$, f2$, rv As Boolean, rtv As Boolean, tke As String, d1$

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "r1q"
'On Error GoTo exr1q
prvs$ = ""
rv = True: rtv = True
n% = 0
'If InStr(d$, "kstrjg55aserg@srjbthsjh.sdjbv") > 0 Then Exit Sub
dx$ = d$
form1.dbg (dx$)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, dx$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  rv = False
End If
form1.dbg (c$)
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then rtv = False
pb1.Max = shm%
pb1.Visible = True
uId$ = form1.getuserid()
While (rtv Or rv) And tm_brk% = 0 And n% < shm%
  If rtv Then If rtmp.EOF Then rtv = False
  If rv Then If r.EOF Then rv = False
  tke = ""
  If rtv And rv Then
    If rtmp!erstellt > r!erstellt Then
      tke = "rtmp"
    Else
      tke = "r"
    End If
  Else
    If rtv Then tke = "rtmp"
    If rv Then tke = "r"
  End If

  If tke = "rtmp" Then
    If rtmp!kontakt <> "-1" Then
      kkid$ = form1.get_kontaktname_by_id(rtmp!kontakt) & " - " & trm(rtmp!adresse)
    Else
      kkid$ = trm(rtmp!adresse)
    End If
    tsts$ = rtmp!erstellt & rtmp!Owner & kkid$ & rtmp!doctyp & rtmp!betreff & form1.dupcheck(trm(rtmp!docname))
    If tsts$ = prvs$ Then
      form1.sqlqry ("delete from dochist where id='" & rtmp!id & "'")
    Else
      dct = transe(rtmp!doctyp)
      If dct = rtmp!doctyp Then dct = transex1(rtmp!doctyp, "_")
      Set lvitem = gd1.ListItems.add(, , rtmp!erstellt & Space$(20) & "(ID:" & rtmp!id)
      lvitem.SubItems(1) = rtmp!Owner
      lvitem.SubItems(2) = kkid$
      If Not IsNull(rtmp!doctyp) Then lvitem.SubItems(3) = dct
      If Not IsNull(rtmp!betreff) Then lvitem.SubItems(4) = rtmp!betreff
      tfn$ = trm(form1.dupcheck(trm(rtmp!docname)))
      fn$ = tfn$
      f2$ = form1.composeemlname(trm(rtmp!docname)): If f2$ <> "" Then fn$ = f2$
      If exist(fn$) Then
        lvitem.SubItems(5) = FileName(tfn$)
      Else
        lvitem.SubItems(5) = "no access"
      End If

      If Mid(rtmp!erstellt, 3, 1) = "." Then
        upd1$ = datum2sql(word1(rtmp!erstellt))
        upd2$ = word2bis(rtmp!erstellt)
        Call form1.sqlqry("update dochist set erstellt='" & upd1$ & " " & upd2$ & "' where id='" & rtmp!id & "'")
      End If
      n% = n% + 1
      If (n% Mod 10) = 0 Then
        dochist2.Caption = trm(n%)
        pb1.value = imin(n%, pb1.Max)
      End If
      DoEvents
    End If
    prvs$ = tsts$
    rtmp.MoveNext
  End If
  
  If tke = "r" Then
    Set lvitem = gd1.ListItems.add(, , r!erstellt & Space$(20) & "(ID:" & r!id)
    lvitem.SubItems(1) = r!Owner
    lvitem.SubItems(2) = r!frm
    lvitem.SubItems(3) = "Mailsafe"
    lvitem.SubItems(5) = FileName(form1.dupcheck(trm(r!message)))
    If Not IsNull(r!Subject) Then lvitem.SubItems(4) = r!Subject
    n% = n% + 1
    r.MoveNext
  End If

Wend
pb1.Visible = False
dochist2.Caption = transe("Kontakthistorie") + ": " & n% & " " + transe("Einträge")
Exit Sub

exr1q:
On Error GoTo 0

End Sub
Sub rtopics()
Dim rtmp As ADODB.Recordset, c$, rrr
Dim d2infile As String, d2insub As String
Dim sw$, sw0$

d2infile = "dochist2": d2insub = "rtopics"

If form1.isfieldmissing("opt_topics", "id") Then Exit Sub
topics.Clear
adrlinks.Clear
sw$ = trm(tfilter.text)
If sw$ <> "" Then
  sw$ = LCase(sw$)
  sw0$ = sw$
  sw$ = " and (lcase(topicid) like '%" + sw$ + "%' or lcase(topicid) like '" + sw$ + "%' or lcase(topicid) like '%" + sw$ + "' or lcase(topicid)='" + sw$ + "')"
End If
c$ = "select * from opt_topics where vid='" + currvid$ + "'" + sw$ + " order by topicid desc"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  While Not rtmp.EOF
    topics.AddItem trm(rtmp!topicid)
    rtmp.MoveNext
  Wend
End If

c$ = "select owner from sysvars where owner like 'sysvar_system_tlnk_%' and wert='" + currvid$ + "' order by owner desc"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  sw$ = Mid$(trm(rtmp!Owner), 20)
  sw$ = Left(sw$, InStr(sw$, currvid$) - 2)
  If sw0$ = "" Or InStr(sw$, sw0$) > 0 Then
    topics.AddItem sw$
  End If
  rtmp.MoveNext
Wend

End Sub

Sub rlist1()
Dim ks$, lic$, mkrit$, madr$, c$, shm2 As Integer, i%, ksm$, adrxact As Boolean, xactc As String

'd2infile = "dochist2": d2insub = "rlist1"
Me.MousePointer = 11
gd1.ListItems.Clear
lic$ = trm(Label1.Caption)
ksm$ = ""
'If lic$ = "" Then Exit Sub
shm = Val(shmax.text)
If shm <= 0 Then
  shm = 20
  shmax.text = trm(shm)
End If

If critter <> "" Then
  Me.MousePointer = 0
  Call getwerkebyort(critter, kid$)
  Exit Sub
End If
s$ = trm(suchw.text)
ks$ = ""
If s$ <> "" Then
  ks$ = "((instr(lcase(betreff),'" + LCase(s$) + "')>0) or (instr(lcase(betreff),'" + LCase(s$) + "')>0) "
  ks$ = ks$ & " or (instr(lcase(adresse),'" + LCase(s$) + "')>0) or (instr(lcase(doctyp),'" + LCase(s$) + "')>0) or (instr(lcase(adresse),'" + LCase(s$) + "')>0)) "
  ksm$ = " (mailsafe.Subject like '%" + s$ + "%' or mailsafe.volltext like '%" + s$ + "%') "
End If
adrxact = False
If form1.getusersetting("dochistadrsearch", "") = "exact" Then adrxact = True
krit$ = "SELECT id,docname,erstellt,kontakt,owner,betreff,doctyp,adresse from dochist "
mkrit$ = "SELECT id,frm,owner,Message,"
If Not form1.isfieldmissing("mailsafe", "optan") Then mkrit$ = mkrit$ + "optan,"
If Not form1.isfieldmissing("mailsafe", "optcc") Then mkrit$ = mkrit$ + "optcc,"
mkrit$ = mkrit$ + "subject,erstellt,frm from mailsafe where "
If msonly.value = 1 Then
  If lic$ <> "" Then krit$ = krit$ + "where (not doctyp like 'Email%') and adresse='" + lic$ + "' "
Else
  If lic$ <> "" Then krit$ = krit$ + "where adresse='" + lic$ + "' "
End If
If currkid$ <> "" And currkid$ <> "-1" Then
  If ks$ <> "" Then
    adw$ = "where ": If lic$ <> "" Then adw$ = "AND "
    krit$ = krit$ & adw$ & ks$
  End If
  krit$ = krit$ + "and ((kontakt='" & currkid$ & "') "
  krit$ = krit$ + ") "
  krit$ = krit$ + "order by erstellt desc"
  form1.dbg (krit$)
  madr$ = form1.allmailadresses(lic$, currkid$): If trm(madr$) = "" Then madr$ = "kstrjg55aserg@srjbthsjh.sdjbv"
  adkrit$ = "("
  While madr$ <> ""
    c$ = cut_d1(madr$, ","): madr$ = cut_d2bis(madr$, ",")
    If adrxact Then
      adkrit$ = adkrit$ + "frm='" + c$ + "' "
    Else
      adkrit$ = adkrit$ + "frm LIKE '%" + c$ + "%' "
    End If
    If Not form1.isfieldmissing("mailsafe", "optan") Then
      If adrxact Then
        adkrit$ = adkrit$ + "or optan='" + c$ + "' "
      Else
        adkrit$ = adkrit$ + "or optan LIKE '%" + c$ + "%' "
      End If
    End If
    If Not form1.isfieldmissing("mailsafe", "optcc") Then
      If adrxact Then
        adkrit$ = adkrit$ + "OR optcc='" + c$ + "' "
      Else
        adkrit$ = adkrit$ + "OR optcc LIKE '%" + c$ + "%' "
      End If
    End If
    adkrit$ = adkrit$ + " or "
  Wend
  If Right$(adkrit$, 3) = "or " Then
    adkrit$ = Left(adkrit$, Len(adkrit$) - 3)
  End If
  adkrit$ = adkrit$ + ") "
  If adkrit$ <> "() " Then mkrit$ = mkrit$ + adkrit$
  If ksm$ <> "" Then mkrit$ = mkrit$ + " and " + ksm$ + " "
  If Right$(mkrit$, 6) = "where " Then
    mkrit$ = Left(mkrit$, Len(mkrit$) - 6)
  End If
shm2 = 2 * shm: If shm2 < 1 Then shm2 = 5
mkrit$ = mkrit$ + " order by erstellt desc limit " + trm(shm2)
'  mkrit$ = mkrit$ + " order by erstellt desc"
Call r1q(krit$, mkrit$, ksm$)
End If
krit$ = "SELECT id,docname,erstellt,kontakt,owner,betreff,doctyp,adresse from dochist "
mkrit$ = "SELECT id,frm,owner,Message,"
If Not form1.isfieldmissing("mailsafe", "optan") Then mkrit$ = mkrit$ + "optan,"
If Not form1.isfieldmissing("mailsafe", "optcc") Then mkrit$ = mkrit$ + "optcc,"
mkrit$ = mkrit$ + "subject,erstellt,frm from mailsafe where "
If msonly.value = 1 Then
  If lic$ <> "" Then krit$ = krit$ + "where (not doctyp like 'Email%') and adresse='" + lic$ + "' "
Else
  If lic$ <> "" Then krit$ = krit$ + "where adresse='" + lic$ + "' "
End If
If ks$ <> "" Then
  adw$ = "where ": If lic$ <> "" Then adw$ = "AND "
  krit$ = krit$ & adw$ & ks$
End If

If currkid$ <> "" And currkid$ <> "-1" Then
  krit$ = krit$ + "and (kontakt='-1') "
End If
If InStr(krit$, "where") = 0 Then krit$ = krit$ + " where (not doctyp like 'Email%')"
krit$ = krit$ + " order by erstellt desc"
limit$ = ""
If form1.uselimitinsql Then limit$ = " limit 0," + trm(shm)
If trm(shm) <> "" Then krit$ = krit$ + limit$ + ";"
madr$ = form1.allmailadresses(lic$, currkid$): If trm(madr$) = "" Then madr$ = "kstrjg55aserg@srjbthsjh.sdjbv"
adkrit$ = "("
While madr$ <> ""
  c$ = cut_d1(madr$, ","): madr$ = cut_d2bis(madr$, ",")
  If trm(c$) <> "" Then
    If adrxact Then
      adkrit$ = adkrit$ + "frm='" + c$ + "' "
    Else
      adkrit$ = adkrit$ + "frm LIKE '%" + c$ + "%' "
    End If
    If Not form1.isfieldmissing("mailsafe", "optan") Then
      If adrxact Then
        adkrit$ = adkrit$ + "or optan='" + c$ + "' "
      Else
        adkrit$ = adkrit$ + "or optan LIKE '%" + c$ + "%' "
      End If
    End If
    If Not form1.isfieldmissing("mailsafe", "optcc") Then
      If adrxact Then
        adkrit$ = adkrit$ + "OR optcc='" + c$ + "' "
      Else
        adkrit$ = adkrit$ + "OR optcc LIKE '%" + c$ + "%' "
      End If
    End If
    adkrit$ = adkrit$ + " or "
  End If
Wend
If Right$(adkrit$, 3) = "or " Then adkrit$ = Left(adkrit$, Len(adkrit$) - 3)
adkrit$ = adkrit$ + ") "
If adkrit$ <> "() " Then mkrit$ = mkrit$ + adkrit$
If ksm$ <> "" Then mkrit$ = mkrit$ + " and " + ksm$ + " "
If Right$(mkrit$, 6) = "where " Then
  mkrit$ = Left(mkrit$, Len(mkrit$) - 6)
End If
shm2 = 2 * shm: If shm2 < 1 Then shm2 = 5
mkrit$ = mkrit$ + " order by erstellt desc limit " + trm(shm2)
Call r1q(krit$, mkrit$, ksm$)

Me.MousePointer = 0
End Sub

Public Sub setkrit(adr$, kid$)
Dim r As ADODB.Recordset
Dim s As ADODB.Recordset, V%, smiet As String, typlist(3) As String, tli As Integer
Dim a$, i As Integer, ol As Integer, j As Integer, k As Integer
Dim d2infile As String, d2insub As String

d2infile = "dochist2": d2insub = "setkrit"
typlist(0) = "konzerte"
typlist(1) = "künstlerauftritt"
typlist(2) = "orchesterauftritt"
aktwerk$ = ""
critter = ""
Command4.Visible = True
csvx.Visible = False
csvxfelder.Visible = False
Call savecheck
If InStr(adr$, "((Repert: ") = 1 Then
  Me.MousePointer = 11
  Call notopics
  runmode = "repert"
  gd1.View = lvwReport
  Label1.Caption = ""
  Label2.Caption = ""
  gd1.ListItems.Clear
  sortlist.Clear
  gd1.ColumnHeaders.Clear
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Interpret"), 4000)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Bezeichnung"), 2500)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("von"), 1000)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("bis"), 1000)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Rolle"), 1500)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("ID"), 0)

  kid$ = Mid$(adr$, InStr(adr$, "((Repert: ") + 10)
  cmd$ = "select * from opt_repertoire where wid='" + kid$ + "' and neverever=0"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    Set lvitem = gd1.ListItems.add(, , trm(r!vid))
    lvitem.SubItems(1) = trm(r!bezeichnung)
    lvitem.SubItems(2) = trm(r!von)
    lvitem.SubItems(3) = trm(r!bis)
    lvitem.SubItems(4) = trm(r!Rolle)
    lvitem.SubItems(5) = trm(r!wid)
    r.MoveNext
  Wend
  Me.MousePointer = 0
  Exit Sub
End If

If InStr(adr$, "((Werke: ") = 1 Then
  critter = adr$
  Call getwerkebyort(adr$, kid$)
  Exit Sub
End If

Label1.Caption = adr$
Label2.Caption = ""
currkid$ = kid$
currvid$ = adr$
If kid$ <> "-1" Then Label2.Caption = form1.get_kontaktname_by_id(kid$)
Call rtopics
Call rlist1
End Sub

Private Sub gd1_Click()
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "gd1_Click"
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
id$ = Mid$(id$, p% + 4)
cmd$ = "select adresse,kontakt,docname,doctyp from dochist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fn$ = form1.dupcheck(trm(r!docname))
  f2$ = form1.composeemlname(fn$)
  If f2$ <> "" Then
    fn$ = f2$
  End If
  Me.Caption = fn$
  If exist(fn$) = 1 Then
    If r!doctyp = "Emaileingang" Then
      Command2.Visible = True
      Command3.Visible = True
    Else
      Command2.Visible = False
      Command3.Visible = False
    End If
  End If
  If r!kontakt <> "-1" Then
    form1.Combo1.text = form1.get_kontaktname_by_id(r!kontakt)
  Else
    form1.Combo1.text = trm(r!adresse)
  End If
End If
End Sub

Private Sub gd1_DblClick()
Dim r As ADODB.Recordset, cl$, mlc$, xx$, id$, mlvr$, kid$, wid$, i%, ext$, f2$

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "gd1_DblClick"

If runmode = "werkhist" Then
  id$ = gd1.SelectedItem.SubItems(5)
  Unload auftritt
  DoEvents
  Load auftritt
  Call auftritt.SetFocus
  Call auftritt.showrec(id$, 0)
  Exit Sub
End If
If runmode = "repert" Then
  wid$ = gd1.SelectedItem
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(wid$, "")
  Call shwAdrDetail.SetFocus
  Exit Sub
End If
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
MousePointer = 11: DoEvents
id$ = Mid$(id$, p% + 4)
cmd$ = "select docname,doctyp from dochist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fn$ = form1.dupcheck(trm(r!docname))
  ext$ = FileExtension(fn$)
  If ext$ = "msg" Then
    f2$ = form1.composeemlname(fn$)
    If f2$ <> "" Then
      fn$ = f2$
    End If
  End If
  If exist(fn$) = 1 Then
    If r!doctyp = "Emaileingang" Or r!doctyp = "Emailausgang" Then
      mlc$ = form1.getusersetting("mailserver")
      cl$ = form1.getusersetting("mailclient")
      If InStr(LCase(cl$), "netscape") > 0 Or LCase(form1.getusersetting("Mozillaclient")) = "ja" Then cl$ = "NETSCAPE47"
      fno$ = FileName(fn$)
      mlvr$ = form1.getusersetting("mailviewer")
      If cl$ = "NETSCAPE47" Then
        mlcl$ = strrepl(form1.getusersetting("netscape47inbox"), """", "")
        If exist(mlcl$) = 0 Then
          mlvr$ = "ja"
        Else
          o% = FreeFile
          Open fn$ For Input As #o%
          p% = FreeFile
          Open mlcl$ For Append As #p%
          Print #p%, "From " & from$
          While Not EOF(o%)
            Line Input #o%, l$
            If Left(LCase(l$), 9) <> "x-mozilla" Then Print #p%, l$
          Wend
          Close #o%
          Close #p%
          If mlvr$ <> "ja" Then
            mlcl$ = form1.getusersetting("mailclient")
            MousePointer = 0: DoEvents
            If exist(word1(mlcl$)) > 0 Then
              X = Shell(mlcl$, 1)
              Exit Sub
            End If
          End If
        End If
      End If
      If InStr(LCase(cl$), "outlook") > 0 Or mlvr$ = "ja" Then
        If exist(fn$) = 1 Then
          xx$ = trm(form1.getmyeditor(FileExtension(fn$)))
          If xx$ <> "" Then
            Call form1.openthisdoc(fn$, "")
          Else
            Load mexplore
            On Error Resume Next
            'do not Call mexplore.SetFocus
            mexplore.fnam = fn$
            On Error GoTo 0
          End If
        End If
      End If
      MousePointer = 0: DoEvents
      Exit Sub
    End If
    MousePointer = 0: DoEvents
    Call form1.openthisdoc(fn$, "")
  Else
     MousePointer = 0: DoEvents
     MsgBox (fn$ + " " + transe("kann nicht gefunden werden."))
  End If
Else
  cmd$ = "select Message from mailsafe where id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    fn$ = form1.dupcheck(trm(r!message))
    ext$ = FileExtension(fn$)
    If ext$ = "msg" Then
      f2$ = form1.composeemlname(fn$)
      If f2$ <> "" Then
        fn$ = f2$
      End If
    End If
    Call form1.openthisdoc(fn$, "")
  End If
End If
MousePointer = 0: DoEvents
End Sub

Private Sub gd1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, sq$
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "gd1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  ask% = MsgBox(transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Historyeintrag löschen?"))
  If ask% = vbYes Then
    ask% = MsgBox(transe("Auch die Dokumente löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Dokumente löschen?"))
    dochist2.MousePointer = 11: DoEvents
    For i = gd1.ListItems.Count To 1 Step -1
      If (gd1.ListItems(i).Selected = True) Then
        id$ = gd1.ListItems(i)
        If gd1.ListItems(i).SubItems(3) <> transe("Rechnungsnummer") And gd1.ListItems(i).SubItems(3) <> transe("Datenänderung") And InStr(gd1.ListItems(i).SubItems(3), transe("Vertragsnummer")) <> 1 Then
        p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
        id$ = Mid$(id$, p% + 4)
        If id$ <> "" Then
          If ask% = vbYes Then
            sq$ = "select docname,doctyp from dochist where id='" + id$ + "'"
            Set r = New ADODB.Recordset
            r.CursorLocation = adUseServer
rrr = form1.adoopen(r, sq$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not r.EOF Then
              fn$ = form1.dupcheck(trm(r!docname))
              If exist(fn$) = 1 Then
                On Error Resume Next
                Kill fn$
                On Error GoTo 0
              End If
            End If
            sq$ = "select message from mailsafe where id='" + id$ + "'"
            Set r = New ADODB.Recordset
            r.CursorLocation = adUseServer
rrr = form1.adoopen(r, sq$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not r.EOF Then
              fn$ = form1.dupcheck(trm(r!message))
              If exist(fn$) = 1 Then
                On Error Resume Next
                Kill fn$
                On Error GoTo 0
              End If
            End If
          End If
          sq$ = "delete from dochist where id='" + id$ + "'"
          Call form1.sqlqry(sq$)
          sq$ = "delete from mailsafe where id='" + id$ + "'"
          Call form1.sqlqry(sq$)
        End If
        End If
      End If
    Next i
    dochist2.MousePointer = 0: DoEvents
    Call rlist1
  End If
End If
If KeyCode = 67 Then
  For i = gd1.ListItems.Count To 1 Step -1
    If gd1.ListItems(i).SubItems(5) = "no access" Then
      gd1.ListItems(i).Selected = True
    Else
      gd1.ListItems(i).Selected = False
    End If
    DoEvents
  Next i
End If
End Sub

Private Sub Label4_DblClick()
Dim now As String, txt As String

now = form1.getusersetting("dochistadrsearch", "")
If now = "" Then
  now = "exact"
  txt = "When searching the contact history an exact match for" + vbCrLf + "mail addresses is now required."
  txt = txt + vbCrLf + "This experimental setting is not recommended (yet)" + vbCrLf + "but speeds up the search dramatically in slow databases." + vbCrLf + "Check the results."
Else
  now = ""
  txt = "When searching the contact history an exact match for" + vbCrLf + "mail addresses is NOT required."
  txt = txt + vbCrLf + "This is the recommended setting.."
End If
Call form1.setusersetting("dochistadrsearch", now)
MsgBox (txt)
End Sub

Private Sub Label7_Click()
If msonly.value = 1 Then
  msonly.value = 0
Else
  msonly.value = 1
End If

End Sub

Private Sub msonly_Click()
If msonly.value = 1 Then
  Call form1.setusersetting("ignoremailsfromdochist", "ja")
Else
  Call form1.setusersetting("ignoremailsfromdochist", "nein")
End If
If msnoupd Then Exit Sub
Call rlist1
End Sub

Private Sub remlist_DblClick()
Dim i%, id$, rrr, cmd$
Dim r As ADODB.Recordset

i% = remlist.ListIndex
If i% < 0 Then Exit Sub
id$ = remlist.List(i%)
pos% = InStr(id$, "(ID:")
If pos% > 0 Then
  id$ = Mid$(id$, pos% + 4)
  On Error Resume Next
  Load create2do
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  Call create2do.initmsg(form1.getuserid(), form1.getuserid(), "", "", Date, Left(Time, 5))
  create2do.Text1(1).Enabled = False
  Call create2do.SetFocus

  cmd$ = "select * from todolist where id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
  If Not r.EOF Then
    If Not IsNull(r!betreff) Then create2do.Text1(3).text = trm(r!betreff)
    If Not IsNull(r!nachricht) Then create2do.Text1(4).text = trm(r!nachricht)
    If Not IsNull(r!datum) Then create2do.Text1(5).text = datfromsql(trm(r!datum))
    If Not IsNull(r!zeit) Then create2do.Combo1.text = trm(r!zeit)
    If Not IsNull(r!an) Then
      For i% = 0 To create2do.List1.ListCount - 1
        If create2do.List1.List(i%) = r!an Then
          create2do.List1.Selected(i%) = True
        Else
          create2do.List1.Selected(i%) = False
        End If
      Next i%
      create2do.Combo1.text = trm(r!zeit)
    End If
    Call form1.sqlqry("delete from todolist where id='" + id$ + "'")
  End If
End If
End Sub

Private Sub shmax_Change()
'd2infile = "dochist2": d2insub = "shmax_Change"

If Not shm_igno Then Call form1.setusersetting("zzzdochistmax", shmax.text)
Call timerreset

End Sub
Sub timerreset()
'd2infile = "dochist2": d2insub = "timerreset"
Timer1.Enabled = False
tm_brk% = 1
DoEvents
Timer1.Interval = form1.getsuchvz()
Timer1.Enabled = True

End Sub

Private Sub suchw_Change()
'd2infile = "dochist2": d2insub = "suchw_Change"
Call timerreset
End Sub

Private Sub tfilter_Change()
Call rtopics
End Sub

Private Sub Timer1_Timer()
'd2infile = "dochist2": d2insub = "Timer1_Timer"
Call form1.dbg2f("dochist Timer1 start")
Timer1.Enabled = False
tm_brk% = 0
DoEvents
Call rlist1
tm_brk% = 0
Call form1.dbg2f("dochist Timer1 exit")

End Sub

Private Sub topic_Change()
Dim t$, le As Double
t$ = strrepl(Trim("" & topic.text), "'", "´")
le = 32000 - Len(t$)
zleft.Caption = cut_d1(fixeur(le), ",") + " " + transe("Zeichen frei")
Me.BackColor = form1.dirtycolor()

End Sub

Private Sub svctt()
Dim c$, t$, i%
t$ = strrepl(Trim("" & topic.text), "'", "´")
c$ = "update opt_topics set toptext='" + t$ + "' where topicid='" + currtopic + "'"
Call form1.sqlqry(c$)
Me.BackColor = form1.cleancolor()
End Sub

Public Sub topics_Click()
Dim i%, uId$, pos As Integer
Dim r As ADODB.Recordset, c$, l$
Dim rtmp As ADODB.Recordset, truelink As Boolean

Call savecheck
i% = topics.ListIndex
If i% < 0 Then Exit Sub

btnPrj.Enabled = False
topic.Top = 0
topic.Left = 0
topic.text = ""
topic.Visible = True
truelink = False
Command25.Enabled = True
Command5.Enabled = True
adrlinks.Clear
remlist.Clear
c$ = "select toptext,vid from opt_topics where topicid='" + topics.List(i%) + "'"
currtopic = topics.List(i%)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
If Not r.EOF Then
  topic.text = Trim("" & r!toptext)
  On Error Resume Next
  Call topic.SetFocus
  On Error GoTo 0
  adrlinks.AddItem trm(r!vid)
  truelink = True
End If
r.Close
If Not form1.isfieldmissing("opt_topics", "id") Then
  c$ = "select * from tplan where ID='" + currtopic + "'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    btnPrj.Enabled = True
    If Not rtmp.EOF Then
      btnPrj.Caption = "--> " + transe("Projekt")
    Else
      If truelink Then
        btnPrj.Caption = "rename topic"
      Else
        btnPrj.Caption = "dead link"
      End If
    End If
  End If
End If
If Right(topic.text, 2) <> vbCrLf Then topic.text = topic.text + vbCrLf
topic.SelLength = 0
topic.SelStart = Len(topic.text)

c$ = "select wert from sysvars where owner like 'sysvar_system_tlnk_" + topics.List(i%) + "_%'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  l$ = trm(rtmp!wert)
  adrlinks.AddItem l$

If Not form1.isfieldmissing("opt_topics", "id") Then
  c$ = "select * from tplan where ID='" + l$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    btnPrj.Enabled = True
    If Not r.EOF Then
      btnPrj.Caption = "--> " + transe("Projekt")
    Else
      btnPrj.Caption = "rename topic"
    End If
  End If
End If
  
  rtmp.MoveNext
Wend

'next action(s)
uId$ = form1.getuserid()
cmd$ = "select * from todolist where Betreff like '" + topics.List(i%) + " [Wiedervorlage] Adresse:%' order by Datum,Zeit"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
While Not r.EOF
  remlist.AddItem r!datum + " " + r!an + ": " + lineof(2, r!nachricht) + Space$(160) + "(ID:" + r!id
  r.MoveNext
Wend
On Error Resume Next
Call topic.SetFocus
On Error GoTo 0
Me.BackColor = form1.cleancolor()
End Sub

Private Sub wvl_Click()
Dim r As ADODB.Recordset, cl$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "dochist2": d2insub = "wvl_Click"
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
id$ = Mid$(id$, p% + 4)
cmd$ = "select docname,doctyp,betreff from dochist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fn$ = form1.dupcheck(r!docname)
  wvlt$ = "Dokument Name": If r!doctyp = "Emaileingang" Then wvlt$ = "Nachricht Name"
  Load create2do
  Call create2do.initmsg(form1.getuserid(), form1.getuserid(), wvlt$ & ":" + _
               form1.dupcheck(trm(r!docname)), trm("Von " & Label1.Caption & " " & Label2.Caption & vbCrLf & r!betreff), Date, Left(Time, 5))
  Call create2do.SetFocus
  create2do.Text1(1).Enabled = False
  create2do.Text1(3).Enabled = False
End If

End Sub

Private Sub savecheck()
Dim antw As Integer
'd2infile = "shwAdrDetail": d2insub = "savecheck"
If Me.BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton1, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command7_Click
  End If
End If
BackColor = form1.cleancolor()
End Sub

Public Sub getwerkebyort(adr$, kid$)
Dim r As ADODB.Recordset, rr1
Dim s As ADODB.Recordset, V%, smiet As String, typlist(3) As String, tli As Integer
Dim a$, i As Integer, ol As Integer, j As Integer, k As Integer
Dim d2infile As String, d2insub As String

d2infile = "dochist2": d2insub = "getwerkebyort"
typlist(0) = "konzerte"
typlist(1) = "künstlerauftritt"
typlist(2) = "orchesterauftritt"
aktwerk$ = ""
Command4.Visible = True
csvx.Visible = False
csvxfelder.Visible = False
Call savecheck
  
  Command4.Visible = False
  csvx.Left = Command4.Left
  csvx.Top = Command4.Top
  csvxfelder.Left = csvx.Left
  csvxfelder.Top = csvx.Top + csvx.Height
  csvx.Visible = True
  csvxfelder.Visible = True
  Me.MousePointer = 11
  Call notopics
  runmode = "werkhist"
  gd1.View = lvwReport
  Label1.Caption = ""
  Label2.Caption = ""
  gd1.ListItems.Clear
  sortlist.Clear
  gd1.ColumnHeaders.Clear
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum"), 1200)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Ort/Abo"), 1600)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Orchester"), 3000)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Dirigent"), 1900)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("Solisten"), 2900)
  Set colHeader = gd1.ColumnHeaders.add(, , transe("ID"), 0)

  kid$ = Mid$(adr$, InStr(adr$, "((Werke: ") + 9)
  aktwerk$ = kid$
  cmd$ = "select programmid from programmliste where werkid='" + kid$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

  If r.EOF Then
    MsgBox transe("Keine Programme gefunden.")
    Me.MousePointer = 0
    Exit Sub
  End If
  ol = 0
  While Not r.EOF
    For tli = 0 To 2
    cmd$ = "SELECT auftritt.Datum as dtg,auftritt.id as aid, auftritt.Bezeichnung, auftritt.Ort, usr_" + typlist(tli) + ".* "
    cmd$ = cmd$ + "FROM auftritt INNER JOIN usr_" + typlist(tli) + " ON auftritt.id = usr_" + typlist(tli) + ".id "
    cmd$ = cmd$ + "WHERE usr_" + typlist(tli) + ".programm='" + trm(r!programmid) + "' order by auftritt.datum desc"
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
    rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr = 0 Then
    While Not s.EOF
      smiet = ""
      Set lvitem = gd1.ListItems.add(, , datfromsql(s!dtg))
      On Error Resume Next: smiet = trm(s!miete): On Error GoTo 0
      If smiet = "" Then
        On Error Resume Next
        smiet = typlist(tli)
        On Error GoTo 0
      End If
      lvitem.SubItems(1) = trm(s!ort) + " " + transe(smiet)
      On Error Resume Next
      lvitem.SubItems(2) = trm(s!ensemble1) + " " + trm(s!ensemble2)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        On Error Resume Next
        lvitem.SubItems(2) = trm(s!orchester)
        On Error GoTo 0
      End If
      On Error Resume Next: lvitem.SubItems(3) = trm(s!dirigent): On Error GoTo 0
      On Error Resume Next
      lvitem.SubItems(4) = trm(s!solist1)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        On Error Resume Next
        lvitem.SubItems(4) = trm(s!künstler)
        On Error GoTo 0
      End If
      On Error Resume Next
      a$ = trm(s!solist2): If a$ <> "" Then lvitem.SubItems(4) = lvitem.SubItems(4) + ", " + a$
      a$ = trm(s!solist3): If a$ <> "" Then lvitem.SubItems(4) = lvitem.SubItems(4) + ", " + a$
      a$ = trm(s!solist4): If a$ <> "" Then lvitem.SubItems(4) = lvitem.SubItems(4) + ", " + a$
      a$ = trm(s!solist5): If a$ <> "" Then lvitem.SubItems(4) = lvitem.SubItems(4) + ", " + a$
      On Error GoTo 0
      lvitem.SubItems(5) = trm(s!aid)
      sortlist.AddItem trm(s!dtg) + " (ID:" + s!aid
      ol = ol + 1
      s.MoveNext
    Wend
    End If
    Next tli
    DoEvents
    r.MoveNext
  Wend
  V% = sortlist.ListCount - 1
  For i = V% To 0 Step -1
    If ol > 0 Then
    For j = 1 To ol
      If gd1.ListItems(j).SubItems(5) = cut_d2bis(sortlist.List(i), ":") Then
        Set lvitem = gd1.ListItems.add(, , datfromsql(word1(sortlist.List(i))))
        For k = 1 To 5
          lvitem.SubItems(k) = gd1.ListItems(j).SubItems(k)
        Next k
        gd1.ListItems.Remove (j)
        j = ol: ol = ol - 1
        DoEvents
      End If
    Next j
    End If
  Next i
  Me.MousePointer = 0
End Sub
