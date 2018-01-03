VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form auftrittslisten 
   Caption         =   "Listenfunktionen"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   ScaleHeight     =   3810
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "auftrittslisten.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Email schreiben"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Adressengruppe"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4680
      Picture         =   "auftrittslisten.frx":00B2
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Vorlage bearbeiten"
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   1800
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Serienbrief"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<--"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "==>"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<=="
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-->"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "finde nur diese Adressgruppen (leer=suche alles)"
      Top             =   720
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "finde nur diese Adressgruppen (leer=suche alles)"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4800
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "auftrittslisten.frx":1204
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schliessen"
      Top             =   3240
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label typ 
      BackStyle       =   0  'Transparent
      Caption         =   "typ"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "use"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vorhandene Listen"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1935
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "auftrittslisten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_DropDown()
'd2infile = "auftrittslisten": d2insub = "Combo1_DropDown"
Call rcombo1

End Sub

Private Sub Command1_Click()
'd2infile = "auftrittslisten": d2insub = "Command1_Click"
Unload Me

End Sub

Private Sub Command16_Click()
Dim o%, l$, rrr, i As Integer

'd2infile = "auftrittslisten": d2insub = "Command16_Click"
Call mkgrp
Load smtp
smtp.Visible = True
Call smtp.SetFocus
smtp.txtServer.Enabled = False
smtp.txtMailFrom.Enabled = False
Call form1.signaturinclude
DoEvents
For i = 0 To smtp.List3.ListCount - 1
  If smtp.List3.List(i) = Text1.Text Then
    smtp.List3.ListIndex = i
    DoEvents
    Call smtp.List3_DblClick
    Exit For
  End If
Next i
Call Command1_Click
End Sub

Private Sub Command2_Click()
'd2infile = "auftrittslisten": d2insub = "Command2_Click"
Call moveme(0, 1)
End Sub

Private Sub Command27_Click()
Dim ffn$, tr$, neuid As String, tr1$, templ$, i%, trgfn$

'd2infile = "auftrittslisten": d2insub = "Command27_Click"
ffn$ = trm(Combo1.Text)
If ffn$ = "" Then Exit Sub
tr$ = form1.vorlagenverzeichnis() + "\"
tr1$ = "serienbrief_" & ffn$ & ".rtf"
tr$ = tr$ & tr1$
If exist(tr$) = 0 Then
  MsgBox "Die Vorlage '" & tr$ & "' existiert nicht."
  Exit Sub
End If
Call form1.openthisdoc(tr$, "noconvert")

End Sub

Private Sub Command3_Click()
'd2infile = "auftrittslisten": d2insub = "Command3_Click"
Call moveme(1, 0)
End Sub

Private Sub Command4_Click()
'd2infile = "auftrittslisten": d2insub = "Command4_Click"
While List3(0).ListCount > 0
  List3(0).ListIndex = 0
  Call moveme(0, 1)
Wend

End Sub

Private Sub Command5_Click()
'd2infile = "auftrittslisten": d2insub = "Command5_Click"
While List3(1).ListCount > 0
  List3(1).ListIndex = 0
  Call moveme(1, 0)
Wend

End Sub

Private Sub Command6_Click()
Dim ifn As String, ofn As String, c As String, i As Integer, j As Integer, c1 As String
Dim rtmp As ADODB.Recordset, c2 As String

Dim d2infile As String, d2insub As String
d2infile = "auftrittslisten": d2insub = "Command6_Click"
If List3(0).ListCount = 0 Then Exit Sub
ifn = form1.s0dir() + "\" & form1.getdbname() & ".rtf\serienbrief_" + Combo1.Text + ".rtf"
If nexist(ifn) Then
  MsgBox (ifn + " " + transe("nicht gefunden"))
  Exit Sub
End If
ofn = form1.s0dir() + "\" & form1.getdbname() & ".rtf\serienbrief_" + strrepl(strrepl(strrepl(trm(Text1.Text), " ", "_"), ":", ""), ".", "") + ".rtf"
If Not nexist(ofn) Then
  On Error Resume Next
  Kill ofn
  On Error GoTo 0
End If
Call FileCopy(ifn, ofn)
Call mkgrp

Call form1.Label6_DblClick
DoEvents
For i = 0 To adrselect.allgrps.ListCount - 1
  If adrselect.allgrps.List(i) = trm(Text1.Text) Then
    adrselect.allgrps.ListIndex = i
    DoEvents
    Exit For
  End If
Next i
ofn = strrepl(strrepl(strrepl(trm(Text1.Text), " ", "_"), ":", ""), ".", "")
adrselect.Combo1.Text = ofn
Call adrselect.Command16_Click
DoEvents
Call Command1_Click
End Sub

Private Sub Command7_Click()
Dim i As Integer

'd2infile = "auftrittslisten": d2insub = "Command7_Click"
Call mkgrp
Call form1.Label6_DblClick
DoEvents
For i = 0 To adrselect.allgrps.ListCount - 1
  If adrselect.allgrps.List(i) = trm(Text1.Text) Then
    adrselect.allgrps.ListIndex = i
    DoEvents
    Exit For
  End If
Next i
DoEvents
Call Command1_Click

End Sub

Private Sub Form_Load()

'd2infile = "auftrittslisten": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Me.Caption = transe("Listenfunktionen")
Label3.Caption = transe("Vorhandene Listen")
Label4.Caption = transe("Benutze")
Command1.ToolTipText = transe("Formular schliessen")
Timer1.Interval = 200
Timer1.Enabled = True
Me.BackColor = form1.cleancolor
Show

End Sub

Private Sub Form_Resize()
'd2infile = "auftrittslisten": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "auftrittslisten": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub Label2_Change(Index As Integer)
'd2infile = "auftrittslisten": d2insub = "Label2_Change"
If Index = 0 Then Text1.Text = strrepl(Label2(Index).Caption, " ", "_")

End Sub

Private Sub List3_DblClick(Index As Integer)
'd2infile = "auftrittslisten": d2insub = "List3_DblClick"
If Index = 0 Then
  Call moveme(0, 1)
Else
  Call moveme(1, 0)
End If

End Sub

Private Sub Timer1_Timer()
'd2infile = "auftrittslisten": d2insub = "Timer1_Timer"
Timer1.Enabled = False
auftrittslisten.SetFocus
End Sub

Private Sub typ_Change()
Dim rrr
Dim rtmp As ADODB.Recordset, c As String

Dim d2infile As String, d2insub As String
d2infile = "auftrittslisten": d2insub = "typ_Change"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftrittsfelder where typ='" + typ.Caption + "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
List3(1).Clear
While Not rtmp.EOF
  c = cut_d1(cut_d2bis(trm(rtmp!feldname), "."), ".")
  If rtmp!zeilen > 10 Then List3(1).AddItem c
  rtmp.MoveNext
Wend

End Sub

Sub moveme(f%, t%)
Dim mi%

'd2infile = "auftrittslisten": d2insub = "moveme"
If List3(f%).ListCount = 0 Then Exit Sub
mi% = List3(f%).ListIndex
If mi% < 0 Then
  List3(f%).ListIndex = 0
  mi% = 0
End If
List3(t%).AddItem List3(f%).List(mi%)
List3(f%).RemoveItem mi%

If List3(f%).ListCount > 0 Then
  If mi% >= List3(f%).ListCount Then mi% = List3(f%).ListCount - 1
  List3(f%).ListIndex = mi%
End If

End Sub

Private Sub rcombo1()
Dim tr, rrr, ffn$

'd2infile = "auftrittslisten": d2insub = "rcombo1"
Combo1.Clear
tr = form1.s0dir() + "\" & form1.getdbname() & ".rtf\serienbrief_*.rtf"
tr = Dir(tr)
rrr = Err
On Error GoTo 0
While tr <> "" And rrr = 0
  ffn$ = basename(Mid$(tr, InStr(tr, "_") + 1), ".rtf")
  Combo1.AddItem ffn$
  tr = Dir
Wend

End Sub

Sub mkgrp()
Dim rrr
Dim c As String, i As Integer, j As Integer, c1 As String
Dim rtmp As ADODB.Recordset, c2 As String, ifn As String

Dim d2infile As String, d2insub As String
d2infile = "auftrittslisten": d2insub = "mkgrp"
If List3(0).ListCount = 0 Then Exit Sub

c = "delete from adressgruppen where grpid='" + trm(Text1.Text) + "'"
Call form1.sqlqry(c)
c = "delete from adressgruppenindex where id='" + trm(Text1.Text) + "'"
Call form1.sqlqry(c)
c = "insert into adressgruppenindex (id) values('" + trm(Text1.Text) + "');"
Call form1.sqlqry(c)
For i = 0 To List3(0).ListCount - 1
  List3(0).ListIndex = i
  DoEvents
  c = "select * from auftritthigru where auftrittsid='" + Label1.Caption + "' and feldname='" + List3(0).List(i) + "';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not rtmp.EOF
    c1 = trm(rtmp!felddaten)
    For j = 1 To linesof(c1)
      c2 = lineof(j, c1)
      ifn = cut_d1(c2, "|")
      If ifn <> "" Then
        c = "insert into adressgruppen (id,adressid,grpid,kid) values('" + _
           form1.newid("adressgruppen", "id", 44) + "','" + _
           ifn + "','" + _
           trm(Text1.Text) + "','-1');"
        Call form1.sqlqry(c)
      End If
    Next j
    rtmp.MoveNext
  Wend
Next i

End Sub
