VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form hordeinfo 
   Caption         =   "WebCal Info"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form2"
   ScaleHeight     =   3945
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox caldav 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   14
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "This session - no cloud"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ISPConfig"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Left            =   3120
      Top             =   3120
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+calendar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+adressbook"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "test all"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "empty all"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "empty selected table"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Outgoing Queue"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   2640
      IntegralHeight  =   0   'False
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.ListBox List2 
      Height          =   1320
      IntegralHeight  =   0   'False
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   1560
      IntegralHeight  =   0   'False
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox accts 
      Height          =   3015
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "hordeinfo.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Auf Wiedersehen!"
      Top             =   3240
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      Caption         =   "Abo-Links"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      ToolTipText     =   "Double click: copy to clipboard, used by Thunderbird, Android, iPhone, ..."
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "color"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "hordeinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim again As Integer, aktshare As String

Private Sub accts_Click()
Dim i As Integer, hid$, c$, r As ADODB.Recordset, cdbase As String
Dim ok As Boolean, ag0 As Integer, c1$, rdrs$, cdav$, acctsel As Integer, adav$

List1.Clear
List2.Clear
List3.Clear
aktshare = ""
cdav$ = "": caldav.text = cdav$
Command3.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
ag0 = again
i = accts.ListIndex: acctsel = i
If i < 0 Then Exit Sub
hid$ = accts.List(i)
c$ = "select share_id,share_name,attribute_name from turba_sharesng where share_owner='" + hid$ + "' and attribute_name='Agencyprof'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
ok = False
While Not r.EOF
  List1.AddItem r!attribute_name + ": ID=" + r!share_name
  'List1.AddItem r!attribute_name + ": ID=" + hid$
  ok = True
  rdrs$ = trm(form1.cloudstaff + form1.cloudmanager)
  While rdrs <> ""
    c$ = cut_d1(rdrs$, "|"): rdrs$ = cut_d2bis(rdrs$, "|")
    If c$ <> "" And c$ <> hid$ Then
      c1$ = "select share_id as wert from turba_sharesng_users where share_id=" + trm(r!share_id) + " and user_uid='" + c$ + "'"
      If form1.get1hordeerg(c1$) = "" Then
        c1$ = "insert into turba_sharesng_users (share_id,user_uid,perm_2,perm_4,perm_8,perm_16) values(" + trm(r!share_id) + ",'" + c$ + "',1,1,1,1)"
        Call form1.xhorde(c1$)
      End If
      c1$ = "update turba_sharesng set share_flags=1 where share_id=" + trm(r!share_id)
      Call form1.xhorde(c1$)
    End If
  Wend
  r.MoveNext
  DoEvents
Wend
For i = 0 To List1.ListCount - 1
  c$ = cut_d2bis(cut_d2bis(List1.List(i), ":"), "=")
  c$ = "select count(*) as wert from turba_objects where owner_id='" + c$ + "'"
  c$ = form1.get1hordeerg(c$)
  List1.List(i) = "(" + c$ + ") " + List1.List(i)
  If c$ <> "0" Then Command3.Enabled = True
  DoEvents
Next i
c$ = "select share_id,share_name,attribute_name from kronolith_sharesng where share_owner='" + hid$ + "' and attribute_name='Agencyprof'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
ok = False
While Not r.EOF
  List2.AddItem r!attribute_name + ": ID=" + r!share_name
  If trm(r!attribute_name) = "Agencyprof" Then
    ok = True
    rdrs$ = trm(form1.cloudstaff + form1.cloudmanager)
    While rdrs <> ""
      c$ = cut_d1(rdrs$, "|"): rdrs$ = cut_d2bis(rdrs$, "|")
      If c$ <> "" And c$ <> hid$ Then
        c1$ = "select share_id as wert from kronolith_sharesng_users where share_id=" + trm(r!share_id) + " and user_uid='" + c$ + "'"
        If form1.get1hordeerg(c1$) = "" Then
          c1$ = "insert into kronolith_sharesng_users (share_id,user_uid,perm_2,perm_4,perm_8,perm_16,perm_1024) values(" + trm(r!share_id) + ",'" + c$ + "',1,1,0,0,0)"
          Call form1.xhorde(c1$)
        End If
        c1$ = "update kronolith_sharesng set share_flags=1 where share_id=" + trm(r!share_id)
        Call form1.xhorde(c1$)
      End If
    Wend
  End If
  r.MoveNext
  DoEvents
Wend
For i = 0 To List2.ListCount - 1
  c$ = cut_d2bis(cut_d2bis(List1.List(i), ":"), "=")
  adav$ = c$
  c$ = cut_d2bis(cut_d2bis(List2.List(i), ":"), "=")
  cdav$ = c$
  c$ = "select count(*) as wert from kronolith_events where calendar_id='" + c$ + "'"
  c$ = form1.get1hordeerg(c$)
  List2.List(i) = "(" + c$ + ") " + List2.List(i)
  If c$ <> "0" Then Command3.Enabled = True
  DoEvents
Next i
cdbase = form1.getusersetting("caldavbase", "https://" + cut_d1(form1.cloudserver$, ":"))
caldav.text = "CardDAV: " + cdbase + "/horde/rpc.php/addressbooks/" + accts.List(acctsel) + "/contacts~" + adav$ + "/"
caldav.text = caldav.text + vbCrLf + "CalDAV : " + cdbase + "/horde/rpc.php/calendars/" + accts.List(acctsel) + "/calendar~" + cdav$ + "/"
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim ttt$, X

      ttt$ = form1.newcloudqfile()
      If ttt$ = "" Then
        MsgBox "error: no queue access"
        Exit Sub
      End If
      ttt$ = DirName(ttt$)
      X = Shell("explorer.exe " & ttt$, vbNormalFocus)
End Sub

Private Sub Command3_Click()
Dim tb$, i, l$, fld$, cmd$

tb$ = ""
i = List1.ListIndex
If i >= 0 Then
  tb$ = "turba_objects"
  l$ = List1.List(i)
  'l$ = cut_d2bis(cut_d2bis(List1.List(i), ":"), "=")
  fld$ = "owner_id"
End If
i = List2.ListIndex
If i >= 0 Then
  tb$ = "kronolith_events"
  fld$ = "calendar_id"
  l$ = List2.List(i)
End If
If tb$ <> "" Then
  l$ = cut_d2bis(l$, "=")
  cmd$ = "delete from " + tb$ + " where " + fld$ + "='" + l$ + "'"
  Call form1.xhorde(cmd$)
  Call accts_Click
End If
End Sub

Private Sub Command4_Click()
Dim i, j

For i = 0 To accts.ListCount - 1
  accts.ListIndex = i
  DoEvents
  For j = 0 To List1.ListCount - 1
    List1.ListIndex = j
    DoEvents
    If Command3.Enabled Then Call Command3_Click
  Next j
  For j = 0 To List2.ListCount - 1
    List2.ListIndex = j
    DoEvents
    If Command3.Enabled Then Call Command3_Click
  Next j
Next i
End Sub

Private Sub Command5_Click()
Dim i, j

accts.ListIndex = -1: DoEvents
For i = 0 To accts.ListCount - 1
  accts.ListIndex = i
  DoEvents
  If Command6.Enabled Or Command7.Enabled Then Exit Sub
Next i

End Sub

Private Sub Command6_Click()
Dim c$, id$
Exit Sub
id$ = mkkey(24)
c$ = "insert into turba_sharesng (share_name,share_owner,attribute_name,attribute_params) values("
c$ = c$ + "'" + id$ + "','" + accts.List(accts.ListIndex) + "'"
c$ = c$ + ",'Agencyprof'"
c$ = c$ + ",'a:2:{s:6:""source"";s:8:""localsql"";s:4:""name"";s:23:""" + id$ + """;}')"
Call form1.xhorde(c$)
Command6.Enabled = False
End Sub

Private Sub Command7_Click()
Dim c$, id$
Exit Sub
id$ = mkkey(24)
c$ = "insert into kronolith_sharesng (share_name,share_owner,attribute_name,attribute_color) values("
c$ = c$ + "'" + id$ + "','" + accts.List(accts.ListIndex) + "'"
c$ = c$ + ",'Agencyprof'"
c$ = c$ + ",'#4e95ff')"
Call form1.xhorde(c$)
Command7.Enabled = False

End Sub

Private Sub Command8_Click()
Dim brw$, X

'd2infile = "shwAdrDetail": d2insub = "Label11_Click"
Unload frmBrowser
DoEvents
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & form1.isp3home, 1)
Else
  MsgBox "You should set your browser using:" + vbCrLf + "UseBrowser=<Path to Browser>"
  frmBrowser.StartingAddress = form1.isp3home
  Load frmBrowser
End If
End Sub

Private Sub Command9_Click()
form1.cloud = False
form1.btncld.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Dim c$, r As ADODB.Recordset, p$

axsResizer1.SaveControlPositions

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

BackColor = form1.cleancolor()
If form1.isp3home <> "" Then Command8.Enabled = True
Show
DoEvents
p$ = ""
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c$ = "select FeldDaten as wert from auftritthigru where FeldName='cloud' and auftrittstyp='webcal'"
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  If p$ <> trm(r!wert) Then
    p$ = trm(r!wert)
    accts.AddItem p$
  End If
  DoEvents
  r.MoveNext
Wend
again = 0
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

Private Sub Label1_Click()
Clipboard.Clear
Clipboard.settext caldav.text

End Sub

Private Sub Label6_Click()

If List2.ListIndex < 0 Then
  MsgBox "to set a calendar colour select a calendar first ;)"
  Exit Sub
End If
Load colorsel
colorsel.SetFocus
colorsel.updc (Label6.BackColor)
Timer2.Enabled = True
Timer2.Interval = 1000

End Sub

Private Sub List1_Click()
Dim i As Integer
Dim c$, r As ADODB.Recordset, acctid As String, acl%

i = List1.ListIndex
If i < 0 Then Exit Sub
acl% = accts.ListIndex
If acl% < 0 Then Exit Sub
List3.Clear
aktshare = ""
Command3.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
List2.ListIndex = -1
c$ = cut_d2bis(cut_d2bis(List1.List(i), ":"), "=")
acctid = accts.List(acl%)
c$ = "select object_lastname,object_alias from turba_objects where owner_id ='" + c$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
While Not r.EOF
  List3.AddItem trm(r!object_lastname) + " (ID:" + trm(r!object_alias) + ")"
  DoEvents
  r.MoveNext
Wend
If List3.ListCount > 0 Then Command3.Enabled = True
End Sub

Private Sub List2_Click()
Dim i As Integer, w1 As String, w2 As String
Dim c$, r As ADODB.Recordset, kfrb(0 To 2)

aktshare = ""
i = List2.ListIndex
If i < 0 Then Exit Sub
Command3.Enabled = False
Command6.Enabled = False
Command7.Enabled = False

List1.ListIndex = -1
List3.Clear
c$ = cut_d2bis(cut_d2bis(List2.List(i), ":"), "=")
aktshare = c$
w1 = "select attribute_color as wert from kronolith_sharesng where share_name='" + c$ + "'"
Label6.Caption = "color: " + form1.get1hordeerg(w1)
w1 = Mid$(trm(cut_d2bis(Label6.Caption, ":")), 2)
For i = 0 To 2
  On Error Resume Next
  kfrb(i) = hex2dec(Mid$(w1, i * 2 + 1, 2))
  On Error GoTo 0
Next i
Label6.BackColor = RGB(kfrb(0), kfrb(1), kfrb(2))

c$ = "select event_title,event_start,event_id,event_creator_id from kronolith_events where calendar_id='" + aktshare + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
While Not r.EOF
  w2 = trm(r!event_start)
  w1 = datum2sql(word1(w2)): w2 = word2bis(w2)
  List3.AddItem w1 + " " + w2 + " " + trm(r!event_title) + " (ID:" + trm(r!event_id) + ")" + " (AP:" + cut_d1(trm(r!event_creator_id), "@")
  DoEvents
  r.MoveNext
Wend
If List3.ListCount > 0 Then Command3.Enabled = True
again = 0
End Sub

Private Sub List3_DblClick()
Dim i%, c$, evid$

i% = List3.ListIndex
If i% < 0 Then Exit Sub
c$ = List3.List(i%)
i% = InStr(c$, "(AP:") + 4
If i% >= Len(c$) Then Exit Sub
evid$ = Mid$(c$, i%)
Unload auftritt
DoEvents
Load auftritt
On Error Resume Next
Call auftritt.SetFocus
On Error GoTo 0
Call auftritt.showrec(evid$, 0)

End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, j%, c$, hordeid$

If KeyCode = 46 Or KeyCode = 8 Then
  i% = List3.ListIndex: j% = i%
  If i% < 0 Then Exit Sub
  c$ = List3.List(i%)
  If InStr(c$, "(AP:") = 0 Then
    MsgBox "you can only remove events"
    Exit Sub
  End If
  i% = InStr(c$, "(ID:") + 4
  If i% >= Len(c$) Then Exit Sub
  hordeid$ = Mid$(c$, i%)
  i% = InStr(hordeid$, ")") - 1
  hordeid$ = Left(hordeid$, i%)
  c$ = "delete from kronolith_events where event_id='" + trm(hordeid$) + "'"
  Debug.Print c$
  Call form1.qhorde(c$)
  List3.RemoveItem j%
End If
End Sub

Private Sub Timer2_Timer()
Dim c As Long, i%, cmd As String
Dim w As Long, r As Long, g As Long, b As Long

c = form1.getcolorselected()

If c < -10 Then Exit Sub
Timer2.Enabled = False
If c < 0 Then Exit Sub

b = c / 65536
w = c Mod 65536
g = w / 256
r = w Mod 256

Label6.BackColor = RGB(r, g, b)
Label6.Caption = "color: #" + LCase(dec2hex(r) + dec2hex(g) + dec2hex(b))
cmd = "update kronolith_sharesng set attribute_color='" + trm(cut_d2bis(Label6.Caption, ":")) + "' where share_name='" + aktshare + "'"
Call form1.xhorde(cmd$)

End Sub
