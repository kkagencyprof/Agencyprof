VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form repertoire 
   Caption         =   "Repertoire"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form2"
   ScaleHeight     =   7515
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "repertoire.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Programm in die Zwischenablage kopieren"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      Height          =   495
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Werk hinzufügen"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      Picture         =   "repertoire.frx":0532
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   3690
      IntegralHeight  =   0   'False
      Left            =   6960
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   1815
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2430
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   7935
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "mark & press <del>/<entf> to blacklist a work"
      Top             =   480
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "repertoire.frx":08D9
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Schliessen"
      Top             =   6720
      Width           =   735
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   7200
      Top             =   240
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "schliesse aus"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Repertoire"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label artid 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "repertoire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mymode, neuwerkid$

Private Sub artid_Change()
Dim c$, r As ADODB.Recordset, id As String, rrr, wid$

id = artid.Caption
If id = "" Then Exit Sub

If id$ = "repertoire_addmode" Then
  mymode = "addmode"
  List2.Visible = True
  Label3.Visible = True
  List1(0).Clear: List1(1).Clear: List3.Clear
  artid.Visible = False
  Command10.Visible = True
  Exit Sub
End If

'Debug.Print id
artid.Visible = True
c$ = "select * from opt_repertoire where vid='" + id + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r.EOF
    wid$ = trm(r!wid)
    If r!neverever = 0 Then
      List1(0).AddItem form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) + Space$(180) + "(WID:" + wid$
    Else
      List1(1).AddItem form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) + Space$(180) + "(WID:" + wid$
    End If
    r.MoveNext
  Wend
End If
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command10_Click()
Dim i%, l$, V$, w$, M%, c$

For i% = 0 To List3.ListCount - 1
  V$ = cut_d1(List3.List(i%), "|")
  w$ = cut_d2bis(List3.List(i%), "|")
  M% = Val(cut_d2bis(w$, "|"))
  w$ = cut_d1(w$, "|")
  c$ = "delete from opt_repertoire where vid='" + V$ + "' and wid='" + w$ + "'"
  Call form1.sqlqry(c$)
  c$ = "insert into opt_repertoire (id,vid,wid,neverever) values("
  c$ = c$ + "'" + form1.newid("opt_repertoire", "id", 20) + "',"
  c$ = c$ + "'" + V$ + "',"
  c$ = c$ + "'" + w$ + "'," + trm(M%) + ")"
  Call form1.sqlqry(c$)
Next i%
Call Command1_Click
End Sub

Private Sub Command3_Click()
Dim wid$, c$, i%, add$, id As String

id = artid.Caption
If id = "" Then Exit Sub

neuwerkid$ = ""
Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus
Call werkvz.callbackinit("repertoire")
While neuwerkid$ = "": DoEvents: Wend
wid$ = neuwerkid$
add$ = form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) + Space$(180) + "(WID:" + wid$
For i% = 0 To List1(0).ListCount - 1
  If List1(0).List(i%) = add$ Then Exit Sub
Next i%
List1(0).AddItem add$

c$ = "insert into opt_repertoire (id,vid,wid,neverever) values("
c$ = c$ + "'" + form1.newid("opt_repertoire", "id", 20) + "',"
c$ = c$ + "'" + id + "',"
c$ = c$ + "'" + wid$ + "',0)"
Call form1.sqlqry(c$)

End Sub

Public Sub callback(prgid$)

neuwerkid$ = prgid$

End Sub

Private Sub Command33_Click()
Dim tx$, i%, l$, p%

Clipboard.Clear
tx$ = ""
'tx$ = form1.rdrep(tx$)
For i% = 0 To List1(0).ListCount - 1
  l$ = List1(0).List(i%)
  p% = InStr(l$, "(WID")
  If p% > 0 Then l$ = trm(Left$(l$, p% - 1))
  tx$ = tx$ + l$ + vbCrLf
Next i%
Clipboard.settext tx$

End Sub

Private Sub Form_Load()
Dim mew, meh, s%

Hide
axsResizer1.SaveControlPositions
mymode = ""
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
mew = form1.mylastwidth(Me.name, 0)
meh = form1.mylastheight(Me.name, 0)
If meh > 0 And mew > 0 Then
  Me.Width = mew
  Me.Height = meh
End If
Call form1.formpos(Me)
s% = form1.myfontsize()
Label1.Caption = transe("Repertoire")
Label2.Caption = transe("schliesse aus")
Label3.Caption = transe("Künstler")
Command3.ToolTipText = transe("Werk hinzufügen")
Show
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

Private Sub List1_DblClick(Index As Integer)
Dim i%, wid$, kid$

i% = List1(Index).ListIndex
If i% < 0 Then Exit Sub

wid$ = List1(Index).List(i%)
i% = InStr(wid$, "(WID:")
If i% = 0 Then Exit Sub
wid$ = Mid$(wid$, i% + 5)
Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus
kid$ = form1.getkompnamebywerkid(wid$)
Call werkvz.showkompdetailbyname(kid$)
Call werkvz.Timer2_Timer
'Call werkvz.showwerkdetail(wid$) war hier nich sooo gut. so bessa
For i% = 0 To werkvz.List2.ListCount - 1
  If InStr(werkvz.List2.List(i%), form1.getwerknamebyid(wid$)) = 1 Then
    werkvz.List2.ListIndex = i%
    Exit Sub
  End If
Next i%

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i%, wid$, kid$, c$, j%

i% = List1(Index).ListIndex
If i% < 0 Then Exit Sub

wid$ = List1(Index).List(i%)
j% = InStr(wid$, "(WID:")
If j% = 0 Then Exit Sub
wid$ = Mid$(wid$, j% + 5)
If KeyCode = 46 Or KeyCode = 8 Then
  c$ = ""
  If mymode = "" Then
    If Index = 1 Then
      c = "delete from opt_repertoire where vid='" + artid.Caption + "' and wid='" + wid$ + "'"
    Else
      c$ = "update opt_repertoire set neverever=1 where vid='" + artid.Caption + "' and wid='" + wid$ + "'"
    End If
  End If
  If mymode = "addmode" Then
    For j% = 0 To List3.ListCount - 1
      If InStr(List3.List(j%), List2.List(List2.ListIndex) + "|" + wid$ + "|") = 1 Then
        If Index = 0 Then List3.AddItem List2.List(List2.ListIndex) + "|" + wid$ + "|1"
        List3.RemoveItem j%
        Exit For
      End If
    Next j%
  End If
  If Index = 0 Then
    List1(1).AddItem List1(0).List(i%)
  End If
  List1(Index).RemoveItem i%
  If c$ <> "" Then Call form1.sqlqry(c$)
End If

End Sub

Private Sub List2_Click()
Dim i%, e$, wid$, w$, c$, j%

List1(0).Clear: List1(1).Clear
j% = List2.ListIndex
If j% < 0 Then Exit Sub

e$ = List2.List(j%)
For i% = 0 To List3.ListCount - 1
  If InStr(List3.List(i%), e$) = 1 Then
    w$ = cut_d2bis(List3.List(i%), "|")
    c$ = cut_d2bis(w$, "|")
    wid$ = cut_d1(w$, "|")
    List1(Val(c$)).AddItem form1.getkompnamebywerkid(wid$) & ": " & form1.getwerknamebyid(wid$) + Space$(180) + "(WID:" + wid$
  End If
Next i%
If List1(0).ListCount = 0 And List1(1).ListCount = 0 Then
  List2.RemoveItem j%
  DoEvents
  If j% >= List2.ListCount Then j% = List2.ListCount - 1
  If j% < 0 Then
    Call Command10_Click
    Exit Sub
  End If
  List2.ListIndex = j%
End If
End Sub
