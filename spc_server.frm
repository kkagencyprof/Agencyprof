VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form spc_server 
   Caption         =   "apdeepspace-server"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Zum Löschen deaktivieren"
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
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
      Left            =   1920
      TabIndex        =   12
      Top             =   1680
      Width           =   255
   End
   Begin VB.ListBox List2 
      Height          =   2400
      IntegralHeight  =   0   'False
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      Picture         =   "spc_server.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "1"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "System(e) erstellen"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "spc_server.frx":0B66
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   3840
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label sysanz 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label maxy 
      BackStyle       =   0  'Transparent
      Caption         =   "maxy"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label miny 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label maxx 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label minx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3615
      Left            =   4200
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3615
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "spc_server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check2_Click()
form1.d2infile="spc_server": form1.d2insub="Check2_Click"
If Check2.value = 1 Then
  Command3.Enabled = True
Else
  Command3.Enabled = False
End If
End Sub

Private Sub Command1_Click()
form1.d2infile="spc_server": form1.d2insub="Command1_Click"
Unload Me

End Sub

Private Sub Command2_Click()
Dim sysid As String

form1.d2infile="spc_server": form1.d2insub="Command2_Click"
sysid = crsys()
Call Command4_Click

End Sub

Private Sub Command3_Click()
Dim dtg As Double
form1.d2infile="spc_server": form1.d2insub="Command3_Click"
MousePointer = 11: DoEvents
Call form1.sqlqry("delete from spc_space;")
Call form1.sqlqry("delete from spc_systems;")
Call form1.sqlqry("delete from benutzerdaten where id='spcgame';")
Call form1.sqlqry("INSERT INTO benutzerdaten (ID) VALUES('spcgame')")
Call form1.sqlqry("delete from sysvars where instr(owner,'sysvar_spcgame')=1;")
Call form1.sqlqry("delete from sysvars where instr(owner,'_spc_')>0;")
Call form1.setusersetting4user("spcgame", "timescale", "0")
dtg = Date + Time
Call form1.setusersetting4user("spcgame", "t0", d2db(dtg))
Call Command4_Click
Call rlist1
Check2.value = 0
MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim rtmp As ADODB.Recordset, c As String

form1.d2infile="spc_server": form1.d2insub="Command4_Click"
c = "select min(px) as wert from spc_space;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  minx.Caption = fixeur(trm0(rtmp!wert))
End If

c = "select max(px) as wert from spc_space;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  maxx.Caption = fixeur(trm0(rtmp!wert))
End If

c = "select min(py) as wert from spc_space;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  miny.Caption = fixeur(trm0(rtmp!wert))
End If

c = "select max(py) as wert from spc_space;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  maxy.Caption = fixeur(trm0(rtmp!wert))
End If

c = "select count(*) as wert from spc_space;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  sysanz.Caption = trm0(rtmp!wert) + " Systeme"
End If


End Sub

Private Sub Command6_Click()
form1.d2infile="spc_server": form1.d2insub="Command6_Click"
Call List1_DblClick
End Sub

Private Sub Form_Load()
form1.d2infile="spc_server": form1.d2insub="Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
BackColor = form1.cleancolor()
Call Command4_Click
Show
Call rlist1

End Sub

Private Sub Form_Resize()
form1.d2infile="spc_server": form1.d2insub="Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
form1.d2infile="spc_server": form1.d2insub="Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Sub rlist1()
Dim rtmp As ADODB.Recordset, usrid As String

form1.d2infile="spc_server": form1.d2insub="rlist1"
List1.Clear
List2.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.rlist1:" & "SELECT * FROM benutzerdaten")
rtmp.Open "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly

While Not rtmp.EOF
  List1.AddItem rtmp!id
  rtmp.MoveNext
Wend
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT * FROM sysvars where instr(owner,'_spc_gamer')>0 and wert='ja'", form1.adoc, adOpenDynamic, adLockReadOnly
While Not rtmp.EOF
  usrid = cut_d2bis(trm(rtmp!Owner), "_")
  usrid = cut_d1(usrid, "_")
  List2.AddItem usrid
  rtmp.MoveNext
Wend

End Sub

Private Sub List1_DblClick()
Dim usrid$, i%
Dim sysid As String


form1.d2infile="spc_server": form1.d2insub="List1_DblClick"
i% = List1.ListIndex
If i% < 0 Then Exit Sub
usrid$ = List1.List(i%)
For i% = 0 To List2.ListCount - 1
  If usrid$ = List2.List(i%) Then Exit Sub
Next i%
If form1.getusersettingfromuser(usrid$, "spc_gamer", "nein") <> "ja" Then
  List2.AddItem usrid$
  Text1.Text = "1"
  sysid = crsys()
  Call Command4_Click
End If
For i% = 0 To List2.ListCount - 1
  If usrid$ = List2.List(i%) Then
    List2.ListIndex = i%
    Call form1.setusersetting4user(usrid$, "spc_gamer", "ja")
    Exit Sub
  End If
Next i%
End Sub

Function crsys() As String
Dim rtmp As ADODB.Recordset, c As String, brk As Boolean, nsys As Integer
Dim px As Double, py As Double, dd As Double, i As Integer, rrr, md As Long, md2 As Long
Dim n As Integer, dst As Long, sysid As String, nam As String

form1.d2infile="spc_server": form1.d2insub="crsys"
crsys = ""
On Error Resume Next
nsys = Val(Text1.Text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then nsys = 1
If nsys < 0 Then nsys = 1
md = 500
md2 = md * md
While nsys > 0
nsys = nsys - 1

dd = 1000
brk = False
While Not brk And dd < 1000000000#
  i = 0
  brk = False
  While i < 10 And Not brk
    px = dd * Rnd - dd / 2
    py = dd * Rnd - dd / 2
    c = "select * from spc_space where (px-" + d2db(px) + ")*(px-" + d2db(px) + ")+(py-" + d2db(py) + ")*(py-" + d2db(py) + ")<" + d2db(md2)
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open c, form1.adoc, adOpenDynamic, adLockReadOnly
    If rtmp.EOF Then brk = True
    i = i + 1
  Wend
  dd = dd * 2
Wend
If Not brk Then
  MsgBox "keinen Platz gefunden."
  Call Command4_Click
  Exit Function
End If
sysid = form1.newid("spc_space", "id", 40)
crsys = sysid
nam = mknam(10)
c = "insert into spc_space (id,px,py) values('" + _
   sysid + "'," + _
   d2db(px) + "," + d2db(py) + ");"
Call form1.sqlqry(c)
n = Int(Rnd * 16): dst = 1 + Int(Rnd * 4)
For i = 1 To n
  c = "insert into spc_systems (id,dst) values('" + _
   sysid + "-" + trm(i) + "'," + _
   trm(dst) + ");"
  Call form1.sqlqry(c)
  dst = dst + 1 + Int(Rnd * 3)
Next i
Text1.Text = trm(nsys)
DoEvents
Wend   'nsys
Text1.Text = "1"

End Function
