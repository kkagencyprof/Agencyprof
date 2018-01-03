VERSION 5.00
Begin VB.Form alarmlist 
   Caption         =   "Benachrichtigungen"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5415
   End
End
Attribute VB_Name = "alarmlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub rlist2()
Dim r As ADODB.Recordset, usrid$, rrr

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id,name FROM benutzerdaten", form1.adoc, dbOpenDynaset, dbReadOnly)

List2.Clear
While Not r.EOF
  List2.AddItem r!id + " (" + trm(r!name) + ")"
  r.MoveNext
Wend

End Sub
Sub rlist1()
Dim i%, ad$

List1.Clear
For i% = 0 To form1.sqla.TableDefs.Count - 1

  ad$ = form1.sqla.TableDefs(i%).name
  List1.AddItem ad$
Next i%


End Sub
Private Sub Command1_Click()
Unload alarmlist
End Sub

Private Sub Command2_Click()
Dim i%, usrid$, t$, p%, ontid$

i% = List1.ListIndex: If i% < 0 Then Exit Sub
t$ = List1.List(i%)

i% = List3.ListIndex: If i% < 0 Then Exit Sub
usrid$ = List3.List(i%)
usrid$ = trm(cut_d1(usrid$, "("))
ontid$ = alarmlist.Caption
p% = InStr(ontid$, ":")
If p% = 0 Then
  form1.sqlqry ("delete from alarmliste where tabelle='" + t$ + "' and uid='" + usrid$ + "'")
Else
  ontid$ = Mid$(ontid$, p% + 1)
  form1.sqlqry ("delete from alarmliste where tabelle='" + t$ + "' and uid='" + usrid$ + "' and ontid='" + ontid$ + "'")
End If
List3.RemoveItem i%

End Sub

Private Sub Command5_Click()
Dim i%, t$, usrid$, j%, ontid$, p%, tst$, k%, tst1$

i% = List1.ListIndex: If i% < 0 Then Exit Sub
t$ = List1.List(i%)

i% = List2.ListIndex: If i% < 0 Then Exit Sub
usrid$ = List2.List(i%)
usrid$ = trm(Left$(usrid$, InStr(usrid$, "(") - 1))
For j% = 0 To List3.ListCount
  If List3.List(j%) = List2.List(i%) Then
    List3.ListIndex = j%
    Exit Sub
  End If
Next j%
ontid$ = alarmlist.Caption
p% = InStr(ontid$, ":")
If p% = 0 Then
  tst$ = usrid$ + " ( "
  For k% = 0 To List3.ListCount - 1
    If List3.List(k%) = tst$ Then
      List3.ListIndex = k%
      Exit Sub
    End If
  Next k%
  form1.sqlqry ("delete from alarmliste where tabelle='" + t$ + "' and uid='" + usrid$ + "'")
  form1.sqlqry ("insert into alarmliste (id,tabelle,uid) values('" + form1.newid("alarmliste", "id", 30) + "','" + t$ + "','" + usrid$ + "')")
  List3.AddItem usrid$ + " ( "
Else
  ontid$ = Mid$(ontid$, p% + 1)
  tst$ = usrid$ + " (" + ontid$
  tst1$ = usrid$ + " ( "
  For k% = 0 To List3.ListCount - 1
    If List3.List(k%) = tst$ Or List3.List(k%) = tst1$ Then
      List3.ListIndex = k%
      Exit Sub
    End If
  Next k%
  form1.sqlqry ("insert into alarmliste (id,tabelle,uid,ontid) values('" + form1.newid("alarmliste", "id", 30) + "','" + t$ + "','" + usrid$ + "','" + ontid$ + "')")
  List3.AddItem usrid$ + " (" + ontid$
End If

End Sub

Private Sub Form_Load()
Dim rtmp As ADODB.Recordset, s As ADODB.Recordset
Randomize
alarmlist.Caption = transe("Benachrichtigungen")
Command1.Caption = transe("&Schliessen")
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

Call rlist1
Call rlist2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub
Public Sub settab(t$)
Dim i%

For i% = 0 To List1.ListCount - 1
  If LCase(t$) = LCase(List1.List(i%)) Then
    List1.ListIndex = i%
    Exit Sub
  End If
Next i%
End Sub

Private Sub List1_Click()

alarmlist.Caption = ""
Call rlist3

End Sub
Sub rlist3()
Dim r As ADODB.Recordset, t$, rrr

If List1.ListIndex < 0 Then Exit Sub
t$ = List1.List(List1.ListIndex)

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM alarmliste where tabelle='" + t$ + "'", form1.adoc, dbOpenDynaset, dbReadOnly)

List3.Clear
While Not r.EOF
  List3.AddItem r!uId + " (" + r!ontid
  r.MoveNext
Wend

End Sub

Private Sub List2_DblClick()
Call Command5_Click

End Sub

Private Sub List3_DblClick()
Call Command2_Click
End Sub
