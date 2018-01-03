VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSComCtl.ocx"
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form kalku 
   Caption         =   "Kalkulation"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   2760
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3600
      Picture         =   "kalku.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Als Vorlage speichern"
      Top             =   3720
      Width           =   495
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
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Eine neue Zeile oberhalb der Markierten hinzufügen"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1080
      Picture         =   "kalku.frx":04F2
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Speichern"
      Top             =   3720
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   330
      IntegralHeight  =   0   'False
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Anzahl der Stellen hinter dem Komma des Ergebnisses"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2880
      Picture         =   "kalku.frx":0899
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Neu berechnen"
      Top             =   3720
      Width           =   615
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
      TabIndex        =   6
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "kalku.frx":13FF
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Formular schliessen"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4200
      Picture         =   "kalku.frx":164F
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "löschen"
      Top             =   3720
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ComboBox waehr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Währung oder andere Einheit"
      Top             =   240
      Width           =   855
   End
   Begin MSComctlLib.ListView gd2 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Vorlage"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label afeld 
      BackStyle       =   0  'Transparent
      Caption         =   "feld"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label atyp 
      BackStyle       =   0  'Transparent
      Caption         =   "atyp"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label aid 
      BackStyle       =   0  'Transparent
      Caption         =   "aid"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label kerg 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "kalku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kid$, ktyp$, kfeld$, kdaten$

Private Function varwertermittlung(V$) As String
Dim rrr, op1 As Double
Dim i%, tbl$, fld$, c$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "kalku": d2insub = "varwertermittlung"
If V$ = "rnd" Then
  varwertermittlung = trm(Int(Rnd * 32768))
  Exit Function
End If
For i% = 1 To gd2.ListItems.Count
  If gd2.ListItems(i%) = V$ Then
    'in dieser Kalkulation gefunden
    varwertermittlung = gd2.ListItems(i%).SubItems(1)
    Exit Function
  End If
Next i%
i% = InStr(V$, "__")
If i% > 0 Then
  tbl$ = Left$(V$, i% - 1)
  fld$ = Mid$(V$, i% + 2)
  c$ = ""
  Select Case tbl$
    Case "finanzen":
      c$ = "select " & fld$ & " from " & tbl$ & " where id='" & kid$ & "'"
    Case "auftritthigru":
      c$ = "select felddaten from " & tbl$ & " where auftrittsid='" & kid$ & "' and auftrittstyp='" & ktyp$ & "' and feldname='" & fld$ & "'"
    Case Else:
  End Select
  If c$ <> "" Then
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    If rrr = 0 Then
      If Not r.EOF Then
        varwertermittlung = word1(strrepl(form1.ohnewaehrung(trm(r.Fields(0).value)), ".", ""))
        Exit Function
      End If
    End If
  End If
End If
c$ = form1.getusersetting(V$, "")
If c$ <> "" Then
  varwertermittlung = word1(strrepl(form1.ohnewaehrung(c$), ".", ""))
  Exit Function
End If
c$ = form1.getsystemsetting(V$, "")
If c$ <> "" Then
  varwertermittlung = word1(strrepl(form1.ohnewaehrung(c$), ".", ""))
  Exit Function
End If
On Error Resume Next
op1 = var2dbl(word1(strrepl(V$, ".", "")))
rrr = Err
On Error GoTo 0
If rrr = 0 Then
'festwertzuweisung
  varwertermittlung = trm(op1)
  Exit Function
End If
varwertermittlung = "0"
End Function

Private Function dofunc(fkt$) As String
Dim rrr
Dim f$, i%, dlist$, fp%
Dim op1 As Double, op2 As Double, erg As Double, p%
Dim o1$, o2$

'd2infile = "kalku": d2insub = "dofunc"
While Left$(fkt$, 1) = "=": fkt$ = Mid$(fkt$, 2): Wend
dlist$ = "+-*/(": i% = Len(dlist$)
f$ = ""
While i% > 0
  p% = InStr(fkt$, Mid$(dlist, i%, 1))
  If p% > 0 Then
    f$ = Mid$(dlist, i%, 1)
    fp% = p%
    i% = 0
  End If
  i% = i% - 1
Wend
If f$ = "" Then
  dofunc = removegarbtrailfromnumber(varwertermittlung(fkt$))
  Exit Function
Else
  o1$ = Left$(fkt$, fp% - 1)
  o2$ = Mid$(fkt$, fp% + 1)
  If f$ = "(" Then
    o2$ = removegarbtrailfromnumber(varwertermittlung(Left(o2$, Len(o2$) - 1)))
    On Error GoTo dofuncerr
    op2 = var2dbl(o2$)
    On Error GoTo 0
    Select Case o1$
      Case "int":
            On Error GoTo dofuncerr
            erg = Int(op2)
            On Error GoTo 0
      Case "fix2":
            On Error GoTo dofuncerr
            dofunc = trm(fixeur(op2))
            On Error GoTo 0
            Exit Function
      Case Else:
            dofunc = "illfunc"
            Exit Function
    End Select
  Else
    o1$ = removegarbtrailfromnumber(varwertermittlung(o1$))
    o2$ = removegarbtrailfromnumber(varwertermittlung(o2$))
    On Error GoTo dofuncerr
    op2 = var2dbl(o2$)
    op1 = var2dbl(o1$)
    Select Case f$
      Case "+": erg = op1 + op2
      Case "*": erg = op1 * op2
      Case "-": erg = op1 - op2
      Case "/": erg = op1 / op2
      Case Else:
    End Select
    rrr = Err
    On Error GoTo 0
  End If
End If
dofunc = trm(erg)
Exit Function
dofuncerr:
On Error GoTo 0
dofunc = "illfunc"
End Function
Private Sub recalc()
Dim i%, l$, j%, c$, zn$, zw$, p%, func$

'd2infile = "kalku": d2insub = "recalc"
MousePointer = 11: DoEvents
gd2.ListItems.Clear

  l$ = kdaten$
  i% = linesof(l$): j% = 1
  While i% > 0
    i% = i% - 1
    c$ = lineof(j%, l$): j% = j% + 1
    p% = InStr(c$, "|")
    If p% > 0 Then
      zn$ = Left$(c$, p% - 1)
      zw$ = Mid$(c$, p% + 1)
      p% = InStr(zw$, "|")
      func$ = ""
      If p% > 0 Then
        func$ = Mid$(zw, p% + 1)
        If p% > 1 Then
          zw$ = Left$(zw$, p% - 1)
        Else
          zw$ = ""
        End If
      End If
      If func$ <> "" Then zw$ = dofunc(func$)
      Set lvitem = gd2.ListItems.add(, , zn$)
      lvitem.SubItems(1) = zw$
      lvitem.SubItems(2) = func$
    End If
  Wend
  On Error Resume Next
  kerg.Caption = fixeur(var2dbl(word1(strrepl(gd2.ListItems(gd2.ListItems.Count).ListSubItems(1), ".", ""))))
  On Error GoTo 0
'Call gd2.SetFocus
MousePointer = 0
End Sub

Private Sub aid_Change()
Dim r As ADODB.Recordset, c$, lvlitem As ListItem
Dim nid$, i%, j%, func$, zw$, zn$, p%, tr, tr0$, o%

Dim d2infile As String, d2insub As String
d2infile = "kalku": d2insub = "aid_Change"
kid$ = aid.Caption
ktyp$ = atyp.Caption
kfeld$ = afeld.Caption
gd2.ListItems.Clear

kdaten$ = ""
c$ = "SELECT * FROM auftritthigru where auftrittsid='" + kid$ & "' and feldname='" & kfeld$ & "' and auftrittstyp='kalku_" & ktyp$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  kdaten$ = trm(r!felddaten)
  If kdaten$ = "" Then
    kdaten$ = kfeld$ & "|" & form1.ohnewaehrung(trm(kerg.Caption))
    waehr.text = form1.nurdiewaehrung(trm(kerg.Caption))
  End If
  Call recalc
Else
  nid$ = form1.newid("auftritthigru", "id", "40")
  c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname) values('" + _
        nid$ & "','" & kid$ & "','kalku_" & ktyp$ & "','" & kfeld$ & "')"
  Call form1.sqlqry(c$)
  kdaten$ = ""
  o% = FreeFile
  On Error Resume Next
  Open form1.vorlagenverzeichnis() + "\kalku_" & ktyp$ & "__" & kfeld$ & "__Standard.txt" For Input As #o%
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    While Not EOF(o%)
      Line Input #o%, l$
      If trm(l$) <> "" Then
        If kdaten$ <> "" Then kdaten$ = kdaten$ & vbCrLf
        kdaten$ = kdaten$ & trm(l$)
      End If
    Wend
    Close #o%
  Else
    kdaten$ = kfeld$ & "|" & form1.ohnewaehrung(trm(kerg.Caption))
    c$ = "update auftritthigru set felddaten='" & kdaten$ & "' where id='" & nid$ & "'"
    Call form1.sqlqry(c$)
    waehr.text = form1.nurdiewaehrung(trm(kerg.Caption))
  End If
  Call recalc
End If
r.Close
End Sub

Private Sub combo1_Change()
Dim tr As String

'd2infile = "kalku": d2insub = "combo1_Change"
tr = form1.s0dir() + "\" + form1.getdbname() & ".rtf\kalku_" & ktyp$ & "__" & kfeld$ & "__" & trm(Combo1.text) & ".txt"
If Not nexist(tr) Then
    i% = FreeFile
    kdaten$ = ""
    i% = FreeFile
    Open tr For Input As #i%
    While Not EOF(i%)
      Line Input #i%, c$
      If trm(c$) <> "" Then
        If kdaten$ <> "" Then kdaten$ = kdaten$ & vbCrLf
        kdaten$ = kdaten$ & trm(c$)
      End If
    Wend
    Close #i%
    Me.BackColor = form1.dirtycolor()
    Call recalc
End If

End Sub

Private Sub Combo1_Click()
'd2infile = "kalku": d2insub = "Combo1_Click"
Call combo1_Change
End Sub

Private Sub Combo1_DropDown()
Dim tr As String, i%, tr0$

'd2infile = "kalku": d2insub = "Combo1_DropDown"
Combo1.Clear
tr = Dir(form1.vorlagenverzeichnis() + "\kalku_" & ktyp$ & "__" & kfeld$ & "__*.txt")
tr0$ = tr
While tr <> ""
  tr0$ = strrepl(tr, "__", " ")
  While InStr(tr0$, " ") > 0
    tr0$ = word2bis(tr0$)
  Wend
  tr0$ = strrepl(tr0$, ".", " ")
  tr0$ = word1(tr0$)
  Combo1.AddItem tr0$
  tr = Dir
Wend
End Sub

Private Sub Command1_Click()
'd2infile = "kalku": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command18_Click()
'd2infile = "kalku": d2insub = "Command18_Click"
Call form1.handbuchcall("10-Termine.htm")
End Sub

Private Sub Command2_Click()
Dim fn$, o%, V$

'd2infile = "kalku": d2insub = "Command2_Click"
V$ = "Standard"
If trm(Combo1.text) <> "" Then V$ = trm(Combo1.text)
myerg$ = trm(InputBox(transe("Geben Sie einen Namen für die Vorlage ein:") & vbCrLf & wert$, "Vorlagenname", V$))
If myerg$ = "" Then Exit Sub

fn$ = form1.vorlagenverzeichnis() + "\kalku_" & ktyp$ & "__" & kfeld$ & "__" & myerg$ & ".txt"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, kdaten$
Close #o%

End Sub

Private Sub Command3_Click()
Dim fnam$, i%, mi%

'd2infile = "kalku": d2insub = "Command3_Click"
i% = 1
Do
  fnam$ = "zw" & trm(i%) & "|"
  i% = i% + 1
Loop Until InStr(kdaten$, fnam$) = 0
l$ = ""
mi% = 0
For i% = 1 To gd2.ListItems.Count
  If l$ <> "" Then l$ = l$ & vbCrLf
  If gd2.ListItems(i%).Selected Then
    l$ = l$ & vbCrLf & fnam$ & "0|" & vbCrLf
    mi% = i%
  End If
  l$ = l$ & gd2.ListItems(i%) & "|" & gd2.ListItems(i%).SubItems(1) & "|" & gd2.ListItems(i%).SubItems(2)
Next i%
kdaten$ = l$
Me.BackColor = form1.dirtycolor()
Call recalc
If mi% > 0 Then gd2.ListItems(mi% + 1).Selected = True
Call gd2.SetFocus

End Sub

Private Sub Command4_Click()

'd2infile = "kalku": d2insub = "Command4_Click"
Call recalc

End Sub

Public Sub Command5_Click()
Dim c$, i%, l$, j%

'd2infile = "kalku": d2insub = "Command5_Click"
l$ = ""
For j% = 1 To gd2.ListItems.Count
 If l$ <> "" Then l$ = l$ & vbCrLf
 l$ = l$ & gd2.ListItems(j%) & "|" & gd2.ListItems(j%).SubItems(1) & "|" & gd2.ListItems(j%).SubItems(2)
Next j%
kdaten$ = l$
c$ = "update auftritthigru set felddaten='" & kdaten$ & "' where auftrittsid='" & aid.Caption & "' and feldname='" & kfeld$ & "' and auftrittstyp='kalku_" & ktyp$ & "'"
Call form1.sqlqry(c$)
For i% = 0 To 33
  If auftritt.Label2(i%).Caption = kfeld$ And trm(kerg.Caption) <> "" Then
    Call auftritt.Text2_GotFocus(i%)
    If InStr(kfeld$, "auslastung") > 0 Then
      auftritt.Text2(i%).text = kerg.Caption & " %"
    Else
      auftritt.Text2(i%).text = kerg.Caption & " " & waehr.text
    End If
    Call auftritt.Text2_LostFocus(i%)
    Me.BackColor = form1.cleancolor()
    DoEvents
    Call auftritt.Command10_Click
    Call Command1_Click
    Exit Sub
  End If
Next i%
Me.BackColor = form1.cleancolor()

End Sub

Private Sub delme_Click()
Dim c$

'd2infile = "kalku": d2insub = "delme_Click"
c$ = "delete from auftritthigru where auftrittsid='" & aid.Caption & "' and feldname='" & kfeld$ & "' and auftrittstyp='kalku_" & ktyp$ & "'"
Call form1.sqlqry(c$)
Call Command1_Click
End Sub

Private Sub Form_Load()
Dim i%

'd2infile = "kalku": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
gd2.View = lvwReport
Set colHeader = gd2.ColumnHeaders.add(, , transe("Name"), 1300)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Wert"), 1200)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Rechenschritt"), 1600)
List1.Clear
For i% = 0 To 8: List1.AddItem trm(i%): Next i%
List1.ListIndex = 2
waehr.Clear
For i% = 0 To form1.waehrungen.ListCount - 1
  waehr.AddItem cut_d1(form1.waehrungen.List(i%), ":")
Next i%
waehr.text = transe("€")
Me.BackColor = form1.cleancolor()
Me.Caption = transe("Kalkulation")
Command2.ToolTipText = transe("Als Vorlage speichern")
Command3.ToolTipText = transe("Eine neue Zeile oberhalb der Markierten hinzufügen")
Command5.ToolTipText = transe("Speichern")
List1.ToolTipText = transe("Anzahl der Stellen hinter dem Komma des Ergebnisses")
Command4.ToolTipText = transe("Neu berechnen")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command1.ToolTipText = transe("Formular schliessen")
delme.ToolTipText = transe("löschen")
waehr.ToolTipText = transe("Währung oder andere Einheit")
Label1.Caption = transe("Vorlage")

Show

End Sub

Private Sub Form_Resize()
'd2infile = "kalku": d2insub = "Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "kalku": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Private Sub gd2_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim i%, l$

'd2infile = "kalku": d2insub = "gd2_AfterLabelEdit"
If Cancel <> 0 Then Exit Sub
l$ = ""
For i% = 1 To gd2.ListItems.Count
  If l$ <> "" Then l$ = l$ & vbCrLf
  If Not gd2.ListItems(i%).Selected Then
    l$ = l$ & gd2.ListItems(i%) & "|" & gd2.ListItems(i%).SubItems(1) & "|" & gd2.ListItems(i%).SubItems(2)
  Else
    l$ = l$ & NewString & "|" & gd2.ListItems(i%).SubItems(1) & "|" & gd2.ListItems(i%).SubItems(2)
  End If
Next i%
kdaten$ = l$
Me.BackColor = form1.dirtycolor()
Call recalc

End Sub

Private Sub gd2_DblClick()
Dim i%, la$, l$, le$, n$

'd2infile = "kalku": d2insub = "gd2_DblClick"
For i% = 1 To gd2.ListItems.Count
  If gd2.ListItems(i%).Selected Then
    'l$ = gd2.ListItems(i%) & "|" & gd2.ListItems(i%).SubItems(1) & "|" & gd2.ListItems(i%).SubItems(2)
    n$ = trm(gd2.ListItems(i%).SubItems(2))
    If n$ = "" Then n$ = trm(gd2.ListItems(i%).SubItems(1))
    n$ = trm(InputBox(transe("Wert von") + " " & gd2.ListItems(i%) & vbCrLf & wert$, transe("Neuer Wert"), "=" & n$))
    If n$ <> "" Then
      If Left$(n$, 1) = "=" Then
        n$ = strrepl(n$, ",", ".")
        l$ = gd2.ListItems(i%) & "| |" & n$
        gd2.ListItems(i%).SubItems(2) = n$
        gd2.ListItems(i%).SubItems(1) = ""
      Else
        l$ = gd2.ListItems(i%) & "|" & n$ & "|"
        gd2.ListItems(i%).SubItems(2) = ""
        gd2.ListItems(i%).SubItems(1) = n$
      End If
      l$ = ""
      For j% = 1 To gd2.ListItems.Count
        If l$ <> "" Then l$ = l$ & vbCrLf
        l$ = l$ & gd2.ListItems(j%) & "|" & gd2.ListItems(j%).SubItems(1) & "|" & gd2.ListItems(j%).SubItems(2)
      Next j%
      kdaten$ = l$
      Me.BackColor = form1.dirtycolor()
      Call recalc
      Exit Sub
    End If
  End If
Next i%
End Sub

Private Sub gd2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim l$, mi%

'd2infile = "kalku": d2insub = "gd2_KeyDown"
mi% = 0
If KeyCode = 8 Or KeyCode = 46 Then
  l$ = ""
  For i% = 1 To gd2.ListItems.Count
    If (Not gd2.ListItems(i%).Selected) Or (LCase(gd2.ListItems(i%))) = LCase(kfeld$) Then
      If l$ <> "" Then l$ = l$ & vbCrLf
      l$ = l$ & gd2.ListItems(i%) & "|" & gd2.ListItems(i%).SubItems(1) & "|" & gd2.ListItems(i%).SubItems(2)
    Else
      mi% = i%
    End If
  Next i%
  kdaten$ = l$
  Me.BackColor = form1.dirtycolor()
  Call recalc
  If mi% > 0 Then gd2.ListItems(mi%).Selected = True
  Call gd2.SetFocus
End If
End Sub

Sub savecheck()
'd2infile = "kalku": d2insub = "savecheck"
If Me.BackColor = form1.dirtycolor() Then
  antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  If antw = vbYes Then
    Call Command5_Click
  End If
End If
End Sub
