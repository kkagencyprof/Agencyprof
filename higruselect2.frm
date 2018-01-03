VERSION 5.00
Begin VB.Form higruselect2 
   Caption         =   "Query Background Data"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "&Start"
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
      Left            =   5760
      TabIndex        =   35
      Top             =   5640
      Width           =   3735
   End
   Begin VB.TextBox sqlq 
      Height          =   2655
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   34
      Top             =   2880
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "higruselect2.frx":0000
      Left            =   2880
      List            =   "higruselect2.frx":000A
      TabIndex        =   33
      Text            =   "und"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox qry 
      Height          =   2655
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   32
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   3720
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   3720
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   3720
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   3720
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3720
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "higruselect2.frx":0019
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Auf Wiedersehen!"
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   30
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   28
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   12
      Left            =   2280
      TabIndex        =   26
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   2280
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   20
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5055
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5775
      Left            =   2040
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "higruselect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim maxl As Integer

Private Sub Combo1_Click()

Call Text1_Change(0)

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim mx%, rrr, c$, na$
Dim r As ADODB.Recordset
Dim s As ADODB.Recordset

Load adrselect
DoEvents
adrselect.Timer1.Enabled = False
DoEvents
On Error Resume Next
Call adrselect.SetFocus
On Error GoTo 0
adrselect.List1.Clear
adrselect.List2.Clear

c$ = sqlq.text + " order by auftritthigru.auftrittsid"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r.EOF
'    Debug.Print trm(r!auftrittsid); " - "; trm(r!vid)
Debug.Print trm(r!auftrittsid)
    If trm(r!opt_kid) <> "" Then
      c$ = "select vid,name from kontakt where id='" + trm(r!opt_kid) + "'"
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
      rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
      If rrr = 0 Then
        na$ = trm(s!name)
        If Not s.EOF Then adrselect.List2.AddItem form1.crlffake(na$) + Space$(160) + " (VID:" + s!vid + ") " + "ID:" + trm(r!opt_kid)
      End If
    Else
      adrselect.List1.AddItem trm(r!auftrittsid) + Space$(80) + "(ID:" + trm(r!auftrittsid)
    End If
    r.MoveNext
  Wend
Else
  adrselect.List1.AddItem "Error in SQL-request"
  adrselect.List2.AddItem "Error in SQL-request"
  If form1.isfieldmissing("auftritthigru", "opt_kid") Then
    adrselect.List1.AddItem "missing field: opt_kid"
    adrselect.List2.AddItem "missing field: opt_kid"
  End If
End If

End Sub

Private Sub Form_Load()
Dim cmd$, s As ADODB.Recordset, rrr, i As Integer

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
maxl = 15
Call nulldsp
Combo1.Clear
Combo1.AddItem transe("oder")
Combo1.AddItem transe("und")
Combo1.text = transe("und")
Show
cmd$ = "SELECT id From adresstypen order by id"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
While Not s.EOF
  If Left(trm(s!id), 4) <> "rel:" Then
    List1.AddItem transe(s!id)
    DoEvents
  End If
  s.MoveNext
Wend
End If

End Sub

Private Sub nulldsp()
Dim i

For i = 0 To maxl - 1
  Label1(i).Caption = ""
  Label1(i).Visible = False
  Text1(i).text = ""
  Text1(i).Visible = False
  qry.text = ""
  sqlq.text = ""
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld1
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld1:
On Error GoTo 0

End Sub

Private Sub List1_Click()
Dim cmd$, s As ADODB.Recordset, i%, t$, cnt As Integer, rrr

i% = List1.ListIndex
If i% < 0 Then Exit Sub

Call nulldsp
DoEvents

t$ = transo(trm(List1.List(i%)))
cmd$ = "SELECT FeldName From auftrittsfelder where typ='" + t$ + "' order by position"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  cnt = 0
  While Not s.EOF And cnt < maxl
    Label1(cnt).Caption = transe(trm(s!feldname))
    Label1(cnt).Visible = True
    Text1(cnt).Visible = True
    cnt = cnt + 1
    s.MoveNext
  Wend
End If

End Sub

Private Sub Text1_Change(Index As Integer)
Dim i As Integer, fld$, srch$, qr$, w$, qr1$, sq$, sq1$, opk$
Dim j As Integer, ij$, l0$, l1$, l2$

qr$ = "": opk$ = "auftritthigru.opt_kid"
If form1.isfieldmissing("auftritthigru", "opt_kid") Then opk$ = "'' as opt_kid"
i = List1.ListIndex
If i < 0 Then Exit Sub
sq$ = "": j = 0: ij$ = "auftritthigru": l0$ = ""
For i = 0 To maxl - 1
  If Label1(i).Visible Then
    srch$ = trm(Text1(i).text)
    If srch$ <> "" Then
      j = j + 1
      If j > 1 Then l0$ = "_" + trm(j - 1)
      fld$ = transo(Label1(i).Caption)
      If qr$ <> "" Then qr$ = qr$ + vbCrLf + transo(Combo1.text)
      If sq$ <> "" Then
        sq1$ = transo(Combo1.text)
        If sq1$ = "und" Then
          sq1$ = "and"
        Else
          sq1$ = "or"
        End If
        sq$ = sq$ + " " + sq1$
      End If
      qr1$ = "": sq1$ = ""
      w$ = word1(srch$): srch$ = word2bis(srch$)
      While w$ <> ""
        If qr1$ <> "" Then qr1$ = qr1$ + " oder"
        If sq1$ <> "" Then sq1$ = sq1$ + " or"
        If Left(w$, 1) <> "=" Then
          qr1$ = qr1$ + " " + fld$ + " enthält " + w$
          sq1$ = sq1$ + " (auftritthigru" + l0$ + ".feldname='" + fld$ + "' and auftritthigru" + l0$ + ".felddaten like '%" + w$ + "%')"
        Else
          w$ = Mid$(w$, 2)
          qr1$ = qr1$ + " " + fld$ + " ist gleich " + w$
          sq1$ = sq1$ + " (auftritthigru" + l0$ + ".feldname='" + fld$ + "' and auftritthigru" + l0$ + ".felddaten='" + w$ + "')"
        End If
        w$ = word1(srch$): srch$ = word2bis(srch$)
      Wend
      qr$ = qr$ + qr1$
      sq$ = sq$ + sq1$
    End If
  Else
    Exit For
  End If
Next i
qr$ = trm(qr$)
For i = 1 To j - 1
  ij$ = "(" + ij$
Next i
For i = 1 To j - 1
  l1$ = "auftritthigru": If i > 1 Then l1$ = l1$ + "_" + trm(i - 1)
  l2$ = "auftritthigru_" + trm(i)
  ij$ = ij$ + " inner join auftritthigru AS " + l2$ + " ON " + l1$ + ".auftrittsid = " + l2$ + ".auftrittsid)"
Next i
i = List1.ListIndex
sq$ = "select auftritthigru.auftrittsid," + opk$ + " from " + ij$ + " where auftritthigru.auftrittstyp='" + transo(List1.List(i)) + "' and (" + trm(sq$) + ")"
If qr$ <> "" And qr$ <> qry.text Then qry.text = qr$
If sq$ <> "" And sq$ <> sqlq.text Then sqlq.text = sq$

End Sub
