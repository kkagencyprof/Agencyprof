VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form todolist 
   BackColor       =   &H00E0E0E0&
   Caption         =   "To Do Liste - AgencyProf"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6600
      Picture         =   "todolist.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "copy to clipboard"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox fltr 
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   160
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IntegralHeight  =   0   'False
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   160
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   480
      TabIndex        =   11
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":0532
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Formular schliessen"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":0782
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":12E8
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "löschen"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":25BE
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "gewählte Nachricht bearbeiten"
      Top             =   720
      Width           =   615
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":3710
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Lege neue Nachricht an"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "todolist.frx":3D7C
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "mehr ..."
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Text            =   "todolist.frx":4EF6
      ToolTipText     =   "Inhalt der ausgewählten Nachricht"
      Top             =   2640
      Width           =   7935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      ItemData        =   "todolist.frx":4EFC
      Left            =   960
      List            =   "todolist.frx":4EFE
      MultiSelect     =   2  'Erweitert
      TabIndex        =   0
      ToolTipText     =   "Nachrichten für gewählten Benutzer"
      Top             =   480
      Width           =   7935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
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
      Left            =   3960
      TabIndex        =   13
      ToolTipText     =   "Liste aller Nachrichten"
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Inhalt der markierten Nachricht"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "Inhalt der ausgewählten Nachricht"
      Top             =   2280
      Width           =   7935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Liste der Nachrichten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Liste aller Nachrichten"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   840
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2160
      Width           =   8175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   840
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "todolist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim uId$, clipcopy As Boolean, clipt$

Sub rlist1()
Dim cmd$, r As ADODB.Recordset, r2 As ADODB.Recordset, didit%, rrr, shw As Boolean
Dim rbetreff As String, id$, currtyp As String, offset, l$
Dim d2infile As String, d2insub As String, i%, cktxt$, srch$
d2infile = "todolist": d2insub = "rlist1"
List1.Clear
didit% = 0

On Error Resume Next
offset = Val(form1.getusersetting("showremindersoffset", "0"))
rrr = Err
On Error GoTo 0
If rrr <> 0 Then offset = 0
srch$ = LCase(trm(fltr.text))
If Not form1.isfieldmissing("opt_checks", "id") Then
  cmd$ = "select * from opt_checks where (isnull(ownr) or ownr='||' or ownr='' or ownr like '%|" + uId$ + "|%'  or ownr like '%" + uId$ + "%') and dtg<='" + datum2sql(trm(CDate(Date) + offset)) + "' and (isnull(confirmed) or confirmed not like 'ok%') order by dtg limit 0,100"
Debug.Print cmd$
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  i% = 0
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    While Not r.EOF
Debug.Print fixl(trm(r!dtg), 11) + " " + r!ownr + " (2CHKID:" + trm(r!id)
      If r!checkid <> "" Then
        cktxt$ = trm(form1.check_pointbyid(r!checkid))
      Else
        cktxt$ = trm(form1.check_optpointbyid(r!id))
      End If
      If srch$ = "" Or InStr(LCase(cktxt$), srch$) > 0 Then
        shw = True
      Else
        If srch$ <> "" Then
          id$ = trm(r!id)


' when searching only: deep packet inspection
  shw = False
  cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
  Set r2 = New ADODB.Recordset
  r2.CursorLocation = adUseServer
  rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r2.EOF Then
    l$ = transe(r2!auftrittstyp) + ": " + Trim("" & r2!checkpoint) + " "
    l$ = l$ + datfromsql(r2!datum) + " " + trm(r2!ort) + " " + trm(r2!bezeichnung)
  Else
    cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
    cmd$ = "SELECT auftritt.Datum, auftritt.Auftrittstyp, auftritt.Bezeichnung, auftritt.Ort, opt_checks.checkpoint, opt_checks.confirmed, opt_checks.id "
    cmd$ = cmd$ + "FROM opt_checks INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id "
    cmd$ = cmd$ + "WHERE opt_checks.id='" + id$ + "'"
    Set r2 = New ADODB.Recordset
    r2.CursorLocation = adUseServer
    rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then
      l$ = transe(r2!auftrittstyp) + ": " + Trim("" & r2!checkpoint) + " "
      l$ = l$ + datfromsql(r2!datum) + " " + trm(r2!ort) + " " + trm(r2!bezeichnung)
    End If
  End If
  If InStr(LCase(l$), srch$) > 0 Then shw = True
          
        End If
      End If
      If shw Then List1.AddItem fixl(trm("checklist"), 8) + " " + fixl(trm(uId$), 8) + " " + fixl(trm(r!dtg), 11) + " " + fixl(cktxt$, 60) + Space$(80) + "(2CHKID:" + trm(r!id)
      r.MoveNext
      i% = i% + 1
    Wend
    If i% >= 49 Then List1.AddItem transe("mehr ... es werden nur 50 Einträge gezeigt.")
  End If
End If


cmd$ = "select * from todolist where An='" + uId$ + "' order by Datum,Zeit limit 0,50"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not r.EOF
  rbetreff = trm(r!betreff)
  shw = True
  If Mid(rbetreff, 1, 19) = "[Wiedervorlage] AT:" Then
    id$ = Mid(rbetreff, 20)
    currtyp = form1.get1erg("select Auftrittstyp as wert from auftritt where id='" + id$ + "'")
    currtyp = form1.get_atabkz(currtyp)
    id$ = form1.get1erg("select Bezeichnung as wert from auftritt where id='" + id$ + "'")
    If id$ <> "" Then rbetreff = currtyp + ": " + id$
  End If
  If srch$ <> "" Then
    shw = False
    If InStr(LCase(rbetreff), srch$) = 0 Then
      cmd$ = "select * from todolist where id='" + trm(r!id) + "'"
      Set r2 = New ADODB.Recordset
      r2.CursorLocation = adUseServer
      rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r2.EOF Then
        If InStr(LCase(trm(r2!nachricht)), srch$) > 0 Then shw = True
      End If
    Else
      shw = True
    End If
  End If
  If r!datum > datum2sql(Date) And didit% = 0 Then
    didit% = 1
    List1.AddItem "--------------------------"
  End If
  If shw Then List1.AddItem fixl(trm(r!von), 8) + " " + fixl(trm(r!an), 8) + " " + fixl(transe(trm(r!Status)), 8) + fixl(trm(r!datum), 11) + trm(r!zeit) + " " + rbetreff + Space$(60) + "(2DOID:" + trm(r!id)
  r.MoveNext
Wend
Text1.text = ""

End Sub


Private Sub Combo1_Click()
Dim i As Integer
'd2infile = "todolist": d2insub = "Combo1_Click"
i = Combo1.ListIndex
If i < 0 Then Exit Sub
uId$ = Combo1.List(i)
Call rlist1
End Sub

Private Sub Command1_Click()
'd2infile = "todolist": d2insub = "Command1_Click"
Unload todolist
End Sub

Private Sub Command18_Click()

'd2infile = "todolist": d2insub = "Command18_Click"
Call form1.handbuchcall("05-Wiedervorlagen.htm")
End Sub

Private Sub Command2_Click()
Dim r As ADODB.Recordset, i%, id$, cmd$, rrr, pos%

Dim d2infile As String, d2insub As String
d2infile = "todolist": d2insub = "Command2_Click"

i% = List1.ListIndex
If i% < 0 Then Exit Sub

id$ = List1.List(i%)
id$ = Mid$(id$, InStr(id$, "(2DOID:") + 7)

pos% = InStr(id$, "(2CHKID:")
If pos% > 0 Then
  id$ = Mid$(id$, pos% + 8)
'  cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
  cmd$ = "SELECT opt_checks.auftrittsid FROM opt_checks WHERE opt_checks.id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  Label1.Caption = ""
  If Not r.EOF Then
'    Unload auftritt: Load auftritt
'    Call auftritt.SetFocus
'    Call auftritt.showrec(trm(r!auftrittsid), 0)
'    Call auftritt.Command21_Click(1)
    Load remedit
    remedit.remid = id$
  End If
  Exit Sub
End If

Call Command5_Click

cmd$ = "select * from todolist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  If Not IsNull(r!betreff) Then create2do.Text1(3).text = trm(r!betreff)
  If Not IsNull(r!nachricht) Then create2do.Text1(4).text = trm(r!nachricht)
  If Not IsNull(r!datum) Then create2do.Text1(5).text = datfromsql(trm(r!datum))
  If Not IsNull(r!zeit) Then create2do.Combo1.text = trm(r!zeit)
  Call form1.sqlqry("delete from todolist where id='" + id$ + "'")
  List1.RemoveItem i%
  Text1.text = ""
End If
End Sub

Private Sub Command3_Click()
Dim from$, l$, mlclf$, kid$, i%
Dim f$, p%, rtmp As ADODB.Recordset, vid$, c$, sid$, fn$, X, rrr, cl$, fno$, mlcl$, o%

Dim d2infile As String, d2insub As String
d2infile = "todolist": d2insub = "Command3_Click"
f$ = Label1.Caption
p% = InStr(f$, "|")
If p% > 0 Then
  sid$ = Mid$(f$, p% + 1)
  f$ = cut_d1(f$, "|")
  If f$ = "at" Then f$ = "auftritt"
  Select Case f$
    Case "DOKBYNAME":
        If exist(sid$) = 0 Then
          MsgBox "Diese Nachricht existiert nicht (mehr?)."
          Exit Sub
        End If
        Call form1.openthisdoc(sid$, "")
    Case "EMAILBYNAME":
        If exist(sid$) = 0 Then
          MsgBox "Diese Nachricht existiert nicht (mehr?)."
          Exit Sub
        End If
        fn$ = sid$
        GoTo handlemail
    Case "voicemail":
        If exist(sid$) = 0 Then
          MsgBox "Diese Nachricht existiert nicht (mehr?)."
          Exit Sub
        End If
        X = Shell("sndrec32.exe " + sid$, 1)
    Case "EMAIL":
        c$ = "select message from mailsafe where id='" + sid$ + "'"
        Set rtmp = New ADODB.Recordset
        rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If rtmp.EOF Then
          MsgBox "Diese Nachricht existiert nicht (mehr?)."
          Exit Sub
        End If
        fn$ = rtmp!message
handlemail:
        cl$ = form1.getusersetting("mailclient")
        cl$ = "OUTLOOK"
        If InStr(LCase(cl$), "netscape") > 0 Or LCase(form1.getusersetting("Mozillaclient")) = "ja" Then cl$ = "NETSCAPE47"
        fno$ = FileName(fn$)
        If cl$ = "NETSCAPE47" Then
          mlcl$ = strrepl(form1.getusersetting("netscape47inbox"), """", "")
          If exist(mlcl$) = 0 Then
            MousePointer = 0: DoEvents
            Call form1.openthisdoc(fn$, "")
          Else
            o% = FreeFile
            Open fn$ For Input As #o%
            p% = FreeFile
            Open mlcl$ For Append As #p%
            Print #p%, "From " + from$
            While Not EOF(o%)
              Line Input #o%, l$
              Print #p%, l$
            Wend
            Close #o%
            Close #p%
            mlcl$ = form1.getusersetting("mailclient")
            MousePointer = 0: DoEvents
            If exist(word1(mlcl$)) > 0 Then
              X = Shell(mlcl$, 1)
            Else
              Call form1.openthisdoc(fn$, "")
            End If
          End If
        End If
        If cl$ = "OUTLOOK" Then
          Load mexplore
          On Error Resume Next
          'do not Call mexplore.SetFocus
          mexplore.fnam = fn$
          On Error GoTo 0
        End If
    Case "auftritt":
        Unload auftritt:
        DoEvents
        Load auftritt
        Call auftritt.SetFocus
        Call auftritt.showrec(sid$, 0)
    Case "kontakt":
        Set rtmp = New ADODB.Recordset
        rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT vid FROM kontakt where id='" + sid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If rtmp.EOF Then Exit Sub
        vid$ = rtmp!vid
        rtmp.Close
        Load shwAdrDetail
        Call shwAdrDetail.refreshadrdetail(vid$, sid$)
        Call shwAdrDetail.SetFocus
    Case "adresse":
        Load shwAdrDetail
        Call shwAdrDetail.refreshadrdetail(sid$, "-1")
        Call shwAdrDetail.SetFocus
        DoEvents
        c$ = lineof(1, trm("" & Text1.text))
        If InStr(c$, "TOPIC: ") = 1 Then
          c$ = trm(cut_d2bis(c$, ":"))
          Call shwAdrDetail.Command10_Click
          DoEvents
          For i% = 0 To dochist2.topics.ListCount - 1
            If c$ = dochist2.topics.List(i%) Then
              dochist2.topics.ListIndex = i%
              Exit For
            End If
          Next i%
        End If
    Case "benutzerdaten":
        Load einstellungen
        Call einstellungen.showrec(sid$)
        einstellungen.SetFocus
    Case "tplan":
        Call tplan.rlists
        Call tplan.nulldsp
        Call tplan.showrec(sid$)
        Call tplan.SetFocus
    Case "w_loc":
        Load werkvz
        Call werkvz.SetFocus
        kid$ = form1.getkompnamebywerkid(sid$)
        Call werkvz.showkompdetailbyname(kid$)
        Call werkvz.showwerkdetail(sid$)
    Case "taliste":
        Load taliste
        Call taliste.SetFocus
        For i% = 0 To taliste.List1.ListCount - 1
          If taliste.List1.List(i%) = sid$ Then
            taliste.List1.ListIndex = i%
            i% = taliste.List1.ListCount
          End If
        Next i%
    Case Default
  End Select
End If

End Sub

Private Sub Command33_Click()
Dim i%

MousePointer = 11: DoEvents
clipcopy = True
Clipboard.Clear
clipt$ = ""
For i% = 0 To List1.ListCount - 1
  List1.ListIndex = i%
  DoEvents
Next i%
clipcopy = False
Call Clipboard.settext(clipt$)
'MsgBox clipt$
MousePointer = 0
End Sub

Public Sub Command4_Click()
'd2infile = "todolist": d2insub = "Command4_Click"
Call rlist1
End Sub

Private Sub Command5_Click()
Dim rrr

'd2infile = "todolist": d2insub = "Command5_Click"
On Error Resume Next
Load create2do
rrr = Err
On Error GoTo 0

If rrr <> 0 Then Exit Sub
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), "" _
             , "", Date, Left(Time, 5))
create2do.Text1(1).Enabled = False
Call create2do.SetFocus

End Sub

Public Sub delme_Click()
Dim r As ADODB.Recordset, i%, id$, rrr, nd$, ask%, f$, p%, ldl%
Dim sid$, pos%, wert$

Dim d2infile As String, d2insub As String
d2infile = "todolist": d2insub = "delme_Click"
MousePointer = 11: DoEvents
ldl% = -1
For i% = 0 To List1.ListCount - 1

If List1.Selected(i%) Then
  id$ = List1.List(i%): ldl% = i%
  pos% = InStr(id$, "(2CHKID:")
  If pos% > 0 Then
    id$ = Mid$(id$, pos% + 8)
    wert$ = "ok, deleted " + trm(Date) + " " + trm(Time) + " " + form1.getuserid()
    sid$ = "update opt_checks set confirmed='" + wert$ + "' where id='" + id$ + "'"
    Call form1.sqlqry(sid$)
  End If
  pos% = InStr(id$, "(2DOID:")
  If pos% > 0 Then

  id$ = Mid$(id$, pos% + 7)
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, "select * from todolist where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    nd$ = datum2sql(Date)
    'ask% = MsgBox("Wirklich löschen?", vbYesNo + vbCritical + vbDefaultButton2, "Nachricht löschen?")
    ask% = vbYes
    If ask% = vbYes Then
      f$ = Label1.Caption
      p% = InStr(f$, "|")
      If p% > 0 Then
        sid$ = Mid$(f$, p% + 1)
        f$ = cut_d1(f$, "|")
        If f$ = "voicemail" Then
          If exist(sid$) <> 0 Then
            On Error Resume Next
            Kill sid$
            On Error GoTo 0
          End If
        End If
      End If
      Text1.text = ""
      form1.sqlqry ("delete from todolist where id='" + id$ + "'")
    End If
  End If

  Else
    pos% = InStr(id$, "(2CHKID:")
    If pos% > 0 Then
'do nothing
    End If
  End If
End If
Next i%
Call rlist1
If ldl% > List1.ListCount - 1 Then ldl% = List1.ListCount - 1
If ldl% >= 0 Then
  List1.Selected(ldl%) = True
  List1.ListIndex = ldl%
  On Error Resume Next
  Call List1.SetFocus
  On Error GoTo 0
End If
'If List1.ListCount > 0 Then List1.ListIndex = 0
If form1.dayvopen Then Call dayvw.Command4_Click
MousePointer = 0

End Sub

Private Sub fltr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command4_Click
End Sub

Private Sub Form_Load()
'd2infile = "todolist": d2insub = "Form_Load"
axsResizer1.SaveControlPositions

todolist.Caption = transe("To Do Liste - AgencyProf")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command1.ToolTipText = transe("Formular schliessen")
Command4.ToolTipText = transe("Ansicht aktualisieren")
delme.ToolTipText = transe("löschen")
Command2.ToolTipText = transe("gewählte Nachricht bearbeiten")
Command5.ToolTipText = transe("Lege neue Nachricht an")
Command3.ToolTipText = transe("Zeige To Do-Liste")
Text1.ToolTipText = transe("Inhalt der ausgewählten Nachricht")
List1.ToolTipText = transe("Nachrichten für gewählten Benutzer")
Label4.Caption = transe("Inhalt der ausgewählten Nachricht")
Label4.ToolTipText = transe("Inhalt der ausgewählten Nachricht")
Label3.Caption = transe("Liste der Nachrichten")
Label3.ToolTipText = transe("Liste aller Nachrichten")
form1.todoisopen = True
clipcopy = False

Show
form1.tdlistisopen = True
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)


uId$ = form1.getuserid()
If form1.getusersetting("alletodolisten", "nein") = "ja" Then
  Call rcombo1
  Combo1.Visible = True
End If
Call rlist1
'If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub

Sub rcombo1()
Dim rtmp As ADODB.Recordset, rrr, i As Integer
Dim d2infile As String, d2insub As String
d2infile = "todolist": d2insub = "rcombo1"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

List1.Clear
While Not rtmp.EOF
  If form1.getuserid() = "www" Or _
     form1.getuserid() = rtmp!id Or _
     form1.getusersettingfromuser(rtmp!id, "appasswort", "") = "" Then
     Combo1.AddItem rtmp!id
  End If
  rtmp.MoveNext
Wend
For i = 0 To Combo1.ListCount - 1
  If Combo1.List(i) = uId$ Then
    Combo1.ListIndex = i
    Exit For
  End If
Next i

End Sub
Private Sub Form_Resize()
'd2infile = "todolist": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "todolist": d2insub = "Form_Unload"
Hide
form1.todoisopen = False
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
form1.tdlistisopen = False
exuld:
On Error GoTo 0
End Sub

Private Sub List1_Click()
Dim r As ADODB.Recordset, rrr, erg As Boolean, id$, cmd$, nd$, n$, rn$
Dim yyyy%, mm%, dd%, d0 As Variant, ti%, i%, prest%, p_in%, tn$
Dim tid$, p1%, sid$

Dim d2infile As String, d2insub As String, pos%
d2infile = "todolist": d2insub = "List1_Click"
i% = List1.ListIndex
If i% < 0 Then Exit Sub

delme.Enabled = True
id$ = List1.List(i%)
If Left$(id$, 5) = "-----" Then Exit Sub
pos% = InStr(id$, "(2CHKID:")
If pos% > 0 Then
  id$ = Mid$(id$, pos% + 8)
  cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  Label1.Caption = ""
  If Not r.EOF Then
    Text1.text = transe(r!auftrittstyp) + ": " + Trim("" & r!checkpoint) + vbCrLf
    Text1.text = Text1.text + datfromsql(r!datum) + " " + trm(r!ort) + vbCrLf + trm(r!bezeichnung)
    cmd$ = trm(r!confirmed): If cmd$ = "" Then cmd$ = "none"
    Text1.text = Text1.text + vbCrLf + "Confirmation: " + cmd$
  Else
    cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
    cmd$ = "SELECT auftritt.Datum, auftritt.Auftrittstyp, auftritt.Bezeichnung, auftritt.Ort, opt_checks.checkpoint, opt_checks.confirmed, opt_checks.id "
    cmd$ = cmd$ + "FROM opt_checks INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id "
    cmd$ = cmd$ + "WHERE opt_checks.id='" + id$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    Label1.Caption = ""
    If Not r.EOF Then
      Text1.text = transe(r!auftrittstyp) + ": " + Trim("" & r!checkpoint) + vbCrLf
      Text1.text = Text1.text + datfromsql(r!datum) + " " + trm(r!ort) + vbCrLf + trm(r!bezeichnung)
      cmd$ = trm(r!confirmed): If cmd$ = "" Then cmd$ = "none"
      Text1.text = Text1.text + vbCrLf + "Confirmation: " + cmd$
    End If
  End If
  If clipcopy Then
    If clipt$ <> "" Then clipt$ = clipt$ + vbCrLf
    On Error Resume Next
    n$ = trm(Mid$(List1.List(i%), 20, pos% - 31))
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      clipt$ = clipt$ + n$
      If trm(Text1.text) <> "" Then clipt$ = clipt$ + vbCrLf + Text1.text + vbCrLf
    Else
      clipt$ = clipt$ + "Error, send this to the support: rrr=" + trm(rrr) + vbCrLf + "todolist, list1.click,pos=" + trm(pos%) + vbCrLf + "l=" + trm(List1.List(i%)) + vbCrLf
    End If
  End If
End If
pos% = InStr(id$, "(2DOID:")
If pos% = 0 Then Exit Sub
id$ = Mid$(id$, pos% + 7)
cmd$ = "select * from todolist where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Label1.Caption = ""
If r.EOF Then Exit Sub
nd$ = datum2sql(Date)
On Error Resume Next
erg = trm(r!an) = form1.getuserid() And (r!datum < nd$ Or (r!datum = nd$ And CDate(r!zeit) <= CDate(Time))) And transo(r!Status) = "neu"
rrr = Err
On Error GoTo 0
If (rrr <> 0 Or erg) And Not clipcopy Then
  form1.sqlqry ("update todolist set status='" + transe("gelesen") + "' where id='" + id$ + "'")
  If r!pdelta <> 0 And trm(r!pdeltaunit$) <> "" And r!poft > 0 Then
    yyyy% = Val(Left$(r!datum, 4))
    mm% = Val(Mid$(r!datum, 6, 2))
    dd% = Val(Right$(r!datum, 2))
    d0 = CDate(datfromsql(r!datum))
    Select Case r!pdeltaunit
      Case "ta": d0 = CDate(d0 + r!pdelta)
      Case "mo": ti% = mm%
                 While ti% > 0
                   mm% = mm% + 1
                   If mm% > 12 Then
                     yyyy% = yyyy% + 1
                   End If
                 Wend
                 d0 = CDate(trm(yyyy%) + "-" + trm(mm%) + "-" + trm(dd%))
      Case "ja": ti% = mm%
                 While ti% > 0
                   yyyy% = yyyy% + 1
                 Wend
                 d0 = CDate(trm(yyyy%) + "-" + trm(mm%) + "-" + trm(dd%))
      Case "wo": d0 = CDate(d0 + r!pdelta * 7)
      Case Default:
    End Select
    prest% = r!poft - 1
    n$ = "": If Not IsNull(r!nachricht) Then n$ = r!nachricht
    Call form1.new2do(r!von, r!an, r!betreff, n$, datum2sql(d0), r!zeit, r!pdelta, r!pdeltaunit, prest%)
  End If
End If
rn$ = trm(r!nachricht)
If Right$(rn$, 1) = ";" Then rn$ = trm(Left$(rn$, Len(rn$) - 1))
If Not clipcopy Then
  Text1.text = rn$
Else
  If clipt$ <> "" Then clipt$ = clipt$ + vbCrLf
  On Error Resume Next
  n$ = trm(Mid$(List1.List(i%), 29, pos% - 31))
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    clipt$ = clipt$ + n$
    If trm(rn$) <> "" Then clipt$ = clipt$ + vbCrLf + rn$ + vbCrLf
  Else
    clipt$ = clipt$ + "Error, send this to the support: rrr=" + trm(rrr) + vbCrLf + "todolist, list1.click,pos=" + trm(pos%) + vbCrLf + "l=" + trm(List1.List(i%)) + vbCrLf
  End If
  Exit Sub
End If
Command3.Enabled = False
p_in% = InStr(trm(r!betreff), "[Wiedervorlage] ")
If p_in% > 0 Then
  tn$ = Mid$(r!betreff, p_in% + 16)
  tid$ = Mid$(tn$, InStr(tn$, ":") + 1)
  tn$ = trm(LCase$(Left$(tn$, InStr(tn$, ":") - 1)))
  If tn$ = "projekt" Then tn$ = "tplan"
  If tn$ = "angebotsliste" Then tn$ = "taliste"
  If tn$ = "adresse" _
       Or tn$ = "taliste" _
       Or tn$ = "auftritt" _
       Or tn$ = "voicemail" _
       Or tn$ = "at" _
       Or tn$ = "tplan" Then
       Label1.Caption = tn$ + "|" + tid$
       Command3.Enabled = True
  End If
End If
If InStr(trm(r!betreff), "Nachricht ID:<") = 1 Then
  tn$ = Mid$(r!betreff, 14)
  tid$ = Mid$(tn$, InStr(tn$, ":") + 1)
  tn$ = "EMAIL"
  Label1.Caption = tn$ + "|" + tid$
  Command3.Enabled = True
End If
If InStr(trm(r!betreff), "Nachricht Name:") = 1 Then
  tn$ = Mid$(r!betreff, 14)
  tid$ = Mid$(tn$, InStr(tn$, ":") + 1)
  tn$ = "EMAILBYNAME"
  Label1.Caption = tn$ + "|" + tid$
  Command3.Enabled = True
End If
If InStr(trm(r!betreff), "Dokument Name:") = 1 Then
  tn$ = Mid$(r!betreff, 14)
  tid$ = Mid$(tn$, InStr(tn$, ":") + 1)
  tn$ = "DOKBYNAME"
  Label1.Caption = tn$ + "|" + tid$
  Command3.Enabled = True
End If
If InStr(trm(r!betreff), "Änderung an") = 1 Then
  If r!von <> "anreden" Then
    Command3.Enabled = True
    If Left$(r!von, 4) = "usr_" Then
      p1% = InStr(rn$, "where id=")
      If p1% > 0 Then
        sid$ = Mid$(rn$, p1% + 12)
        sid$ = Left$(sid$, Len(sid$) - 1)
        Label1.Caption = "auftritt|" + sid$
        Exit Sub
      End If
    End If

    If r!von = "adresse" _
       Or r!von = "kontakt" _
       Or r!von = "benutzerdaten" _
       Or r!von = "w_loc" _
       Or r!von = "tplan" _
       Then
      p1% = InStr(rn$, "where id=")
      If p1% > 0 Then
        sid$ = Mid$(rn$, p1% + 10)
        sid$ = Left$(sid$, Len(sid$) - 1)
        Label1.Caption = r!von + "|" + sid$
        Exit Sub
      End If
    End If
  End If
End If

End Sub

Private Sub List1_DblClick()
Dim r As ADODB.Recordset, rrr, erg As Boolean, id$, cmd$, nd$, n$, rn$
Dim yyyy%, mm%, dd%, d0 As Variant, ti%, i%, prest%, p_in%, tn$
Dim tid$, p1%, sid$

Dim d2infile As String, d2insub As String, pos%
i% = List1.ListIndex
If i% < 0 Then Exit Sub
id$ = List1.List(i%)
If Left$(id$, 5) = "-----" Then Exit Sub
pos% = InStr(id$, "(2CHKID:")
If pos% > 0 Then
  id$ = Mid$(id$, pos% + 8)
'  cmd$ = "SELECT opt_checklists.checkpoint, opt_checks.auftrittsid, opt_checks.confirmed, opt_checklists.id, auftritt.Auftrittstyp,auftritt.Datum, auftritt.Bezeichnung, auftritt.Ort FROM (opt_checks INNER JOIN opt_checklists ON opt_checks.checkid = opt_checklists.id) INNER JOIN auftritt ON opt_checks.auftrittsid = auftritt.id WHERE opt_checks.id='" + id$ + "'"
  cmd$ = "SELECT opt_checks.auftrittsid FROM opt_checks WHERE opt_checks.id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  Label1.Caption = ""
  If Not r.EOF Then
    Unload auftritt
    DoEvents
    Load auftritt
    Call auftritt.SetFocus
    Call auftritt.showrec(trm(r!auftrittsid), 0)
    Call auftritt.Command21_Click(1)
  End If
  Exit Sub
End If

Call Command3_Click
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'd2infile = "todolist": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then Call delme_Click

End Sub

Private Sub Text1_DblClick()
Dim txt$
'd2infile = "todolist": d2insub = "Text1_DblClick"
txt$ = "" + Text1.text
Load memoview
Call memoview.SetFocus
Call memoview.settext(txt$)

End Sub


