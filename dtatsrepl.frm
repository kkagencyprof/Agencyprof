VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form Datenreplikator 
   Caption         =   "Datenreplikator"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4320
      Top             =   2520
   End
   Begin VB.CheckBox autoconnect 
      Caption         =   "autoconnect"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   2580
      Width           =   1335
   End
   Begin VB.PictureBox cb1 
      Height          =   255
      Index           =   1
      Left            =   6120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox cb1 
      Height          =   255
      Index           =   0
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "dbtest"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   615
   End
   Begin VB.ListBox dellist 
      Height          =   2790
      IntegralHeight  =   0   'False
      Left            =   7680
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox id2get 
      Height          =   2790
      IntegralHeight  =   0   'False
      Left            =   6480
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Turbo"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "verbinden"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Text            =   "59000"
      Top             =   2340
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5760
      Top             =   2280
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      Picture         =   "dtatsrepl.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Schliessen, Übertragungen fortsetzen"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "dtatsrepl.frx":062A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Schliessen, Übertragungen beenden"
      Top             =   2400
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1965
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   1965
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   1680
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   11
      Left            =   6000
      Picture         =   "dtatsrepl.frx":087A
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   17
      Left            =   4680
      Picture         =   "dtatsrepl.frx":0A04
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   26
      Left            =   5040
      Picture         =   "dtatsrepl.frx":0B8E
      Top             =   0
      Width           =   360
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fallbackverzeichnis"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   7
      Top             =   2460
      Width           =   255
   End
   Begin VB.Label fallbackq 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aktualisierung alle"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fallbackserver"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Datenreplikator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim backslashhandler As String, dbrepl_syncdrive As String
Dim sqlf As ADODB.Connection
Dim wrkJet As Workspace
Dim tm2lock As Boolean, chkingdb As Boolean, get_didsomething As Boolean
'Dim sqla As Database
Dim connok As Boolean

Sub chknewfiles()
Dim rtmp As ADODB.Recordset, rrr, c As String, buffs%, i%, terg As Long
Dim r As ADODB.Recordset, fn$, o%, Src$, dst$, someerr As Boolean

If form1.getusersetting("ismailhandler", "") <> "ja" Then Exit Sub
If form1.isfieldmissing("opt_topics", "id") Then Exit Sub
c = "select id,toptext,topicid from opt_topics where toptext='EOF' and vid='Agencyprof' and kid='-1'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c, form1.adoc, adOpenDynamic, adLockReadOnly, "Datenreplikator", "chknewfiles")
While Not rtmp.EOF
  c = cut_d1(rtmp!topicid, "_")
  buffs% = Val(c)
  fn$ = cut_d2bis(rtmp!topicid, "_")
  dst$ = form1.mylocaldatadir() + "\mail\inbox\" + fn$
  Src$ = form1.s0dir() + "\tmp\" + fn$ + ".b64"
  o% = FreeFile
  someerr = False
  Open Src$ For Output As #o%
  dellist.Clear
  For i% = 0 To buffs% - 1
    c = "select id,toptext from opt_topics where topicid='" + trm(i%) + "_" + fn$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c, form1.adoc, adOpenDynamic, adLockReadOnly, "Datenreplikator", "chknewfiles")
    DoEvents
    If Not r.EOF Then
      c = strrepl(r!toptext, vbCrLf, "")
      Print #o%, c;
      dellist.AddItem r!id
      DoEvents
      If Not form1.isstarting Then Call delay(200)
    Else
      someerr = True
    End If
  Next i%
  Close #o%
  If Not someerr Then
    DoEvents
    Call tm_start(0)
    Call DecodeFileB64(Src$, dst$)
    terg = tm_stop(0)
    If terg > 200 Then
      List1(1).AddItem trm(Date) + " " + trm(Time) + " decode " + trm(terg) + "ms"
      List1(1).ListIndex = List1(1).ListCount - 1
    End If
    On Error Resume Next
    Kill Src$
    On Error GoTo 0
    DoEvents
  End If
  c = "delete from opt_topics where id='" + rtmp!id + "'":
  Call form1.sqlqry(c)
  While dellist.ListCount > 0
    c = "delete from opt_topics where id='" + dellist.List(0) + "'"
    Call form1.sqlqry(c)
    dellist.RemoveItem 0
    DoEvents
  Wend
  rtmp.MoveNext
Wend
'form1.BackColor = form1.cleancolor()
End Sub

Private Sub autoconnect_Click()
If autoconnect.value = 1 Then
  Call form1.setusersetting("replicationautoconnect", "1")
Else
  Call form1.setusersetting("replicationautoconnect", "0")
End If
End Sub

Private Sub Command1_Click()

form1.Command16.Height = 615
Unload Me
End Sub

Public Sub Command2_Click()
Dim rrr, myid$, c$, dbfp$, pos%, i%
Dim r As ADODB.Recordset

  On Error Resume Next
  Command2.Enabled = False
  MousePointer = 11
  For i% = 0 To 1: Call cbset(i%, "yellow"): Next i%
  DoEvents
  c$ = form1.getusersetting("fallbackserverdatenbank", "")
  dbrepl_syncdrive = form1.getusersetting("syncdrive", "")
  'Set sqlf = wrkJet.OpenDatabase(c$, dbDriverCompleteRequired, False, form1.dbfpara$)
  Set sqlf = New ADODB.Connection
  sqlf.ConnectionString = form1.dbfpara$
  On Error Resume Next
  sqlf.Open
  rrr = Err
  On Error GoTo 0
  MousePointer = 0
  If rrr <> 0 Then
    dbfp$ = form1.dbfpara$: pos% = InStr(LCase(dbfp$), "passw")
    If pos% > 0 Then dbfp$ = Left$(dbfp$, pos%) + "..."
    For i% = 0 To 1: Call cbset(i%, "red"): Next i%
    MsgBox ("Verbindung mit dem Replikationsserver ist fehlgeschlagen." + vbCrLf + dbfp$)
    form1.dbfpara$ = ""
    Command2.Enabled = True
    On Error Resume Next
    sqlf.Close
    On Error GoTo 0
    Exit Sub
  End If
  c$ = "select count(lfdnr) as rc from opt_repliken"
  If form1.getusersetting("ismailhandler", "") <> "ja" Then
    c$ = c$ + " where (id not like '%-mailtraffic-%') "
  End If
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
  If rrr = 0 Then
    If r!rc > 0 Then
      r.Close
      Call Command4_Click
      Exit Sub
    End If
  End If
  connok = True
  myid$ = form1.getsystemsetting("replid_" + form1.computername, "")
  If myid$ = "" Then
    myid$ = GUID()
    c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 20) + "',"
    c$ = c$ + "'sysvar_system_replid_" & form1.computername & "','" + myid$ + "')"
    form1.sqlqry (c$)
    c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 20) + "',"
    c$ = c$ + "'sysvar_system_replnode_" & form1.computername & "','" + form1.getdbname() + "')"
    form1.sqlqry (c$)
  End If
End Sub

Public Sub Command27_Click()

Hide
End Sub

Private Sub Command3_Click()

Timer2.Interval = 50
Timer1.Enabled = False
End Sub

Public Sub Command4_Click()
Dim fn$, o%, rrr, ltme As Long, dbver$, myid$, c$, i%
Dim mynode$, ask As Integer
Dim r As ADODB.Recordset

connok = False: DoEvents
For i% = 0 To 1: Call cbset(i%, "red"): Next i%
c$ = "delete from sysvars where owner='sysvar_system_replicationmode'": form1.sqlqry (c$)
Me.BackColor = form1.cleancolor()
myid$ = form1.computername
mynode$ = form1.computername
Call form1.sqlqry("delete from sysvars where owner='sysvar_system_replid_" + form1.computername + "'")
Call form1.sqlqry("delete from sysvars where owner='sysvar_system_replikant_" + form1.computername + "'")
Call form1.sqlqry("delete from sysvars where owner='sysvar_system_replnode_" + form1.computername + "'")
Call form1.setusersetting4user("system", "replid_" + form1.computername, myid$)
Call form1.setusersetting4user("system", "replikant_" + form1.computername, myid$)
Call form1.setusersetting4user("system", "replnode_" + form1.computername, myid$)
dbver$ = form1.getusersetting("dbserno", "1")
If dbver$ = "1" Then Call form1.setusersetting4user("system", "dbserno", "1")

fn$ = form1.replicationfilename(form1.computername)
ltme = 0
c$ = "select max(lfdnr) as rc from opt_repliken"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
  If Not r.EOF Then
    ltme = CLng(trm0(r!rc))
  End If
Else
  Call MsgBox("Error accessing the replication table." + vbCrLf + "No reset done.")
  Exit Sub
End If
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, ltme
Print #o%, dbver$
Close #o%
Command2.Enabled = True
On Error Resume Next
form1.adoc.Execute "delete from opt_repliken"
On Error GoTo 0

'MsgBox "Die Replikationsverbindung wurde zurückgesetzt." + vbCrLf + "Die Verbindung wird aufgebaut ..."
List1(1).Clear: DoEvents
Call Command2_Click
End Sub

Private Sub chktable(tbl$, idf$)
Dim r0 As ADODB.Recordset
Dim r1 As ADODB.Recordset, vid$
Dim f$, c$, rrr, cnt, cntf, i%, rcnt, upd$

c = "select count(*) as cnt from " + tbl$
Set r0 = New ADODB.Recordset
r0.CursorLocation = adUseServer
rrr = form1.adoopen(r0, c, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
cnt = 0
If rrr = 0 Then cnt = r0!cnt
Set r0 = New ADODB.Recordset
r0.CursorLocation = adUseServer
rrr = form1.adoopen(r0, c, sqlf, adOpenDynamic, adLockReadOnly, "", "")
cntf = 0
If rrr = 0 Then cntf = r0!cnt
If cntf = 0 Or cnt = 0 Then Exit Sub
List1(1).AddItem "testing " + trm(cnt) + " records from " + tbl$: rcnt = 0
List1(1).ListIndex = List1(1).ListCount - 1
c = "select * from " + tbl$
Set r0 = New ADODB.Recordset: r0.CursorLocation = adUseServer
rrr = form1.adoopen(r0, c, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r0.EOF
    DoEvents
    rcnt = rcnt + 1
    If (rcnt Mod 500) = 0 Then
      List1(1).AddItem trm(rcnt) + " records tested"
      List1(1).ListIndex = List1(1).ListCount - 1
    End If
    f$ = "select * from " + tbl$ + " where " + idf$ + "='" + r0!id + "'"
    Set r1 = New ADODB.Recordset: r1.CursorLocation = adUseServer
    rrr = form1.adoopen(r1, f$, sqlf, adOpenDynamic, adLockReadOnly, "", "")
    If rrr = 0 Then
      If Not r1.EOF Then
        upd$ = ""
        For i% = 0 To r0.Fields.Count - 1
          DoEvents
          If r1.Fields(i%).value <> r0.Fields(i%).value And r1.Fields(i%).name <> "tstamp" And r1.Fields(i%).name <> "Stand" Then
            Debug.Print idf$ + "='" + r0!id + "' " + r1.Fields(i%).name + ": " + trm(r0.Fields(i%).value) + " vs. " + trm(r1.Fields(i%).value)
            upd$ = trm(r0!id)
          End If
        Next i%
      Else
        Debug.Print "Datensatz fehlt in " + tbl$ + ": " + idf$ + "='" + r0!id + "'"
        upd$ = trm(r0!id)
      End If
      If upd$ <> "" Then
        List1(1).AddItem tbl$ + " " + idf$ + "=" + trm(r0!id) + " nicht ok"
        List1(1).ListIndex = List1(1).ListCount - 1
        Select Case (tbl$)
          Case "kontakt":
            Load shwAdrDetail
            Call shwAdrDetail.savecheck
            vid$ = form1.getadridbykontaktid(upd$)
            If vid$ <> "" Then
              Call shwAdrDetail.refreshadrdetail(vid$, upd$)
              DoEvents
              Call shwAdrDetail.Command5_Click
            Else
            End If
            DoEvents
          Case "auftritt":
            Unload auftritt
            DoEvents
            Load auftritt
            Call auftritt.SetFocus
            Call auftritt.showrec(upd$, 0)
            DoEvents
            Call auftritt.Command10_Click
            DoEvents
          Case "tplan":
            Load tplan
            Call tplan.rlists
            Call tplan.nulldsp
            Call tplan.showrec(upd$)
            DoEvents
            Call tplan.Command17_Click
            DoEvents
          Case "adresse":
            Load shwAdrDetail
            Call shwAdrDetail.savecheck
            Call shwAdrDetail.refreshadrdetail(upd$, "")
            DoEvents
            Call shwAdrDetail.Command4_Click
            DoEvents
          Case Else:
            List1(1).AddItem tbl$ + " " + idf$ + "=" + trm(r0!id)
            List1(1).ListIndex = List1(1).ListCount - 1
        End Select
      End If
    Else
      Debug.Print "Fehler #" + trm(rrr) + " " + Error$(rrr)
      Debug.Print "readerror from fallbackserver in chktable, query=" + f$
      Exit Sub
    End If
    r0.MoveNext
  Wend
End If

End Sub

Private Sub Command5_Click()
chkingdb = True
Call chktable("adresse", "ID")
Call chktable("kontakt", "id")
Call chktable("auftritt", "id")
Call chktable("tplan", "ID")
chkingdb = False
List1(1).AddItem "test finished"
List1(1).ListIndex = List1(1).ListCount - 1
End Sub

Private Sub Form_Load()
Dim ti1 As Long, i%, autoconn

connok = False
tm2lock = False
chkingdb = False
get_didsomething = False
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Call Timer1_Timer
Text1.text = 10000
backslashhandler = form1.getusersetting("backslashhandler", "an")
Datenreplikator.Caption = transe("Datenreplikator")
Command3.Caption = transe("Turbo")
Command2.Caption = transe("verbinden")
Command4.ToolTipText = transe("Zurücksetzen der Replikationsverbindung")
Command27.ToolTipText = transe("Schliessen, Übertragungen fortsetzen")
Command1.ToolTipText = transe("Schliessen, Übertragungen beenden")
autoconnect.value = 0
If form1.getusersetting("replicationautoconnect", "0") = "1" Then
  autoconnect.value = 1
  Timer3.Enabled = True
End If
fallbackq(3).Caption = transe("ms")
fallbackq(2).Caption = transe("Aktualisierung alle")
fallbackq(1).Caption = transe("In:")
fallbackq(0).Caption = form1.getusersetting("fallbackserver", "")
If form1.getusersetting("replicationmode", "") = "hold" Then Me.BackColor = RGB(255, 0, 0)
If form1.getusersetting("dbcheckok", "") <> form1.getuserid() Then Command5.Visible = False
For i% = 0 To 1: Call cbset(i%, "red"): Next i%

Show
If Timer3.Enabled Then Hide
End Sub

Private Sub Form_Resize()
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call cbset(0, "red")
Call cbset(1, "red")
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub Image1_Click(Index As Integer)
Dim o%, fn$, i%, X, c$

If Index = 11 Then
  c$ = "delete from sysvars where owner='sysvar_system_replicationmode'": form1.sqlqry (c$)
  c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 20) + "',"
  c$ = c$ + "'sysvar_system_replicationmode','hold')"
  form1.sqlqry (c$)
  c$ = "delete from sysvars where owner='sysvar_system_dbserno'": form1.sqlqry (c$)
  c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 20) + "',"
  c$ = c$ + "'sysvar_system_dbserno','" + trm(Int(10000 * Rnd())) + "')"
  form1.sqlqry (c$)
  Me.BackColor = RGB(255, 0, 0)
  Exit Sub
End If
If Index = 17 Then
  If List1(1).ListCount > 0 Then
    fn$ = form1.mydatadir() + "\replog-incoming.txt"
    o% = FreeFile
    Open fn$ For Output As #o%
    For i% = 0 To List1(1).ListCount - 1
      Print #o%, List1(1).List(i%)
    Next i%
    Close #o%
    X = Shell("notepad.exe " + fn$, vbNormalFocus)
  End If
  Exit Sub
End If
If Index = 26 Then
  List1(1).Clear
  Exit Sub
End If
End Sub


Private Sub List1_DblClick(Index As Integer)
Dim i%

If Index = 1 Then
  i% = List1(1).ListIndex
  If i% < 0 Then Exit Sub
  MsgBox (List1(1).List(i%))
End If
End Sub

Private Sub Text1_Change()
Dim rrr, ti1 As Long

On Error Resume Next
ti1 = Val(Text1.text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then ti1 = 59000
Text1.text = ti1
Timer1.Enabled = False
If ti1 < 1000 Then ti1 = 1000
Timer1.Interval = ti1
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Dim X


Call rlist1
form1.fallbackq(1) = trm(List1(0).ListCount)
If List1(0).ListCount > 0 Then
  Timer2.Interval = 500
  tm2lock = False
  Timer2.Enabled = True
  Timer1.Enabled = False
Else
  If get_didsomething = True Then
    Timer1.Enabled = False
    Timer1.Interval = 1000
    Timer1.Enabled = True
  Else
    Call Text1_Change
  End If
End If
'If dbrepl_syncdrive <> "" Then
'  Debug.Print dbrepl_syncdrive
'End If
End Sub

Sub rlist1()
Dim fallbackdir$, tr, i%, c1t%, o%, ltme As Long, myid$, mynode$, rdtg As String, rid As String
Dim r As ADODB.Recordset, fn$, c$, rrr, ntme$, ask%, dbver$, lfd$, terg As Long

get_didsomething = False
If chkingdb Then Exit Sub
If form1.fallbackserverpath$ <> "" Then
  Call cbset(0, "yellow")
  tr = Dir(form1.fallbackserverpath$ & "\*.sql")
  c1t% = 0
  If tr = "" Then Call cbset(0, "green")
  While tr <> ""
    If InStr(tr, "-" & form1.getuserid() & ".") > 0 Then
      c1t% = c1t% + 1
      For i% = 0 To List1(0).ListCount - 1
        If tr = List1(0).List(i%) Then i% = List1(0).ListCount + 10
      Next i%
      If i% < List1(0).ListCount + 5 Then List1(0).AddItem tr
    End If
    tr = Dir
  Wend
  If c1t% <> List1(0).ListCount Then List1(0).Clear
End If
If connok Then
  If form1.isfieldmissing("opt_repliken", "id") Then
    For i% = 0 To 1: Call cbset(i%, "red"): Next i%
    Exit Sub
  End If
  dbver$ = ""
  myid$ = form1.getsystemsetting("replid_" + form1.computername, "")
  If myid$ <> "" And myid$ = form1.getsystemsetting("replikant_" + form1.computername, "") Then
    mynode$ = form1.getsystemsetting("replnode_" + form1.computername, form1.getdbname())
    ltme = 0
    fn$ = form1.replicationfilename(myid$)
    If Not nexist(fn$) Then
      o% = FreeFile
      Open fn$ For Input As #o%
      Line Input #o%, c$
      On Error Resume Next
      Line Input #o%, dbver$
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then dbver$ = ""
      Close #o%
      ltme = CLng(trm(c$))
    End If
    If form1.getusersetting("dbserno", "") <> dbver Then
      If form1.getusersetting("replicationmode", "") <> "hold" Then
        connok = False
        For i% = 0 To 1: Call cbset(i%, "red"): Next i%
        List1(1).AddItem "Datenbankversionskonflikt"
        List1(1).ListIndex = List1(1).ListCount - 1
        MsgBox "Datenbankversionskonflikt." + vbCrLf + "Bitte zuerst die Datenbank ersetzen," + vbCrLf + "DANACH auf [Reset] im Replikationsfenster klicken." + vbCrLf + form1.getusersetting("dbserno", "") + " vs. " + trm(dbver) + "(" + form1.replicationfilename(myid$) + ")", vbCritical, "Datenreplikation beendet."
        Exit Sub
      End If
    End If
    If List1(1).ListCount = 0 Then List1(1).AddItem trm(Date) + " " + trm(Time) + " " + transe("Replikant läuft")
    If id2get.ListCount = 0 Then
      c$ = "select lfdnr from opt_repliken where lfdnr>" + trm(ltme) + " and sourcenode<>'" + mynode$ + "' "
      If form1.getusersetting("ismailhandler", "") <> "ja" Then
        c$ = c$ + " and (id not like '%-mailtraffic-%') "
      End If
      c$ = c$ + "order by lfdnr limit 0,150"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
      Call cbset(1, "yellow")
      Call tm_start(0)
      rrr = form1.adoopen(r, c$, sqlf, adOpenDynamic, adLockReadOnly)
      terg = tm_stop(0)
      If terg > 500 Then
        List1(1).AddItem trm(Date) + " " + trm(Time) + " request " + trm(terg) + "ms needed"
        List1(1).ListIndex = List1(1).ListCount - 1
      End If
      If Not r.EOF Then
        While Not r.EOF And connok
          id2get.AddItem trm(r!lfdnr)
          r.MoveNext
          DoEvents
        Wend
      End If
    End If
    If id2get.ListCount = 0 Then
      Call cbset(1, "green")
      form1.fallbackq(0).Caption = form1.getusersetting("fallbackserver", "")
    End If
    While id2get.ListCount > 0
      get_didsomething = True
      lfd$ = id2get.List(0)
      id2get.RemoveItem 0
      DoEvents
      c$ = "select lfdnr,id,daten from opt_repliken where lfdnr=" + lfd$
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
      rrr = form1.adoopen(r, c$, sqlf, adOpenDynamic, adLockReadOnly)
'      Call delay(200)          ' min delay is 100
      If Not r.EOF Then
        While Not r.EOF And connok
          c$ = strrepl(trm(r!daten), "-|-ap-|-", "'")
          c$ = strrepl(c$, "|bckslsh|", "\")
          If backslashhandler = "an" Then
            c$ = strrepl(c$, "\", "|backslashbackslash|")
            c$ = strrepl(c$, "|backslashbackslash|", "\\")
          End If
          On Error Resume Next
          form1.adoc.Execute c$
          DoEvents
          rrr = Err
          On Error GoTo 0
          If rrr <> 0 Then
'          MsgBox "Schreibfehler bei der Übertragung aus der Replikationsdatenbank:" + vbCrLf + c$
            List1(1).AddItem trm(Date) + " " + trm(Time) + " Schreibfehler (" + trm(rrr) + ") bei der Übertragung aus der Replikationsdatenbank"
            List1(1).AddItem c$
            List1(1).ListIndex = List1(1).ListCount - 1
'        connok = False
          End If
          rid = trm(r!id)
          ltme = r!lfdnr
          DoEvents
          r.MoveNext
        Wend
        rdtg = cut_d1(rid, ":") + ":" + cut_d1(cut_d2bis(rid, ":"), ":")
        form1.fallbackq(0).Caption = rdtg: DoEvents
      End If
      DoEvents
    Wend
'!!!!!erst nach der transaktion!
    If get_didsomething Then
      fn$ = form1.replicationfilename(myid$)
      o% = FreeFile
      Open fn$ For Output As #o%
      Print #o%, ltme
      Print #o%, dbver$
      Close #o%
      List1(1).AddItem trm(Date) + " " + trm(Time) + " #" + trm(ltme) + " ok"
      List1(1).ListIndex = List1(1).ListCount - 1
    End If
  End If
End If
Call chknewfiles
End Sub

Private Sub Timer2_Timer()
Dim dbp$, rrr, fn$, o%, l$, sq$, replikant$, c$, Node$
Dim rtmp As QueryDef, tid$, sqd$, tr

If tm2lock Then Exit Sub
tm2lock = True
If List1(0).ListCount = 0 Then
  Timer2.Enabled = False
  Timer1.Enabled = True
  If form1.fallbackserverpath$ <> "" Then
    tr = Dir(form1.fallbackserverpath$ & "\*.sql")
    If tr = "" Then
      Call cbset(0, "green")
    Else
      Call cbset(0, "yellow")
    End If
  End If
Else
  If form1.dbfpara$ <> "" And connok Then
    Call cbset(0, "yellow")
    fn$ = form1.fallbackserverpath$ & "\" & List1(0).List(0)
    If Not nexist(fn$) Then
      o% = FreeFile
      Open fn$ For Input As #o%
      While Not EOF(o%)
        sq$ = ""
        Do
          On Error Resume Next
          Line Input #o%, l$
          rrr = Err
          On Error GoTo 0
          If rrr <> 0 Then
            If rrr = 62 Then GoTo errrx
            MsgBox "Fehler Nr." & rrr & " beim Import eines SQL-Kommandos." & vbCrLf & Error$(rrr)
            End
          End If
            If Len(sq$) > 0 Then sq$ = sq$ & vbCrLf
            sq$ = sq$ + l$
        Loop Until Right$(trm(sq$), 1) = ";"
        If trm(sq$) <> ";" And InStr(LCase(sq$), "set tstamp") = 0 Then
'Debug.Print sq$
          If InStr(LCase(sq$), "set tstamp") = 0 Then
            Call xQD(sq$)
            If Not form1.isfieldmissing("opt_repliken", "id") Then
              replikant$ = form1.getsystemsetting("replid_" + form1.computername, "")
              If replikant$ <> "" Then
                Node$ = form1.getsystemsetting("replnode_" + form1.computername, "")
                tid$ = datum2sql(trm(Date)) + trm(Time()) + GUID()
                sqd$ = strrepl(sq$, "'", "-|-ap-|-")
                sqd$ = strrepl(sqd$, "\", "|bckslsh|")
                c$ = "insert into opt_repliken (id,replikator,sourcenode,daten) values('" + tid$ + "',"
                c$ = c$ + "'" + replikant$ + "','" + Node$ + "','" + sqd$ + "')"
                Call xQD(c$)
              End If
            End If
          End If
        End If
      Wend
errrx:
      Close #o%
      If connok Then Kill fn$
    End If
    If connok Then
      List1(0).RemoveItem 0
    Else
      Call cbset(0, "red")
    End If
  End If
End If
form1.fallbackq(1).Caption = List1(0).ListCount

tm2lock = False
End Sub

Public Sub xQD(qdfTemp As String)
  Dim rrr

  On Error Resume Next
  sqlf.Execute qdfTemp
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
'    MsgBox "Fehler beim Schreiben in die Replikationsdatenbank:" + vbCrLf + qdfTemp
'    connok = False
     Call form1.errhdl("Fehlernummer: " & rrr & vbCr & Error$(rrr) & "statement=" & qdfTemp)
  End If
End Sub

Private Sub cbset(n%, col$)
Select Case (col$)
  Case "yellow": cb1(n%).BackColor = RGB(255, 128, 0)
  Case "green": cb1(n%).BackColor = RGB(0, 255, 0)
  Case Else:   cb1(n%).BackColor = RGB(255, 0, 0)
End Select
cb1(n%).Cls
If n% = 0 Then
  form1.cb3.BackColor = cb1(n%).BackColor
  form1.cb3.Cls
Else
  form1.cb4.BackColor = cb1(n%).BackColor
  form1.cb4.Cls
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Call Command2_Click
End Sub
