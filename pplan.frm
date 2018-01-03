VERSION 5.00
Begin VB.Form pplan 
   Caption         =   "Projekte"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton mallminus 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   20
      ToolTipText     =   "Einen Monat zurück"
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton mallplus 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      ToolTipText     =   "Einen Monat vor"
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton allclose 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   480
      Picture         =   "pplan.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   18
      ToolTipText     =   "Schliesst alle Formulare"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<<"
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
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
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
      Left            =   1920
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton refr 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      Picture         =   "pplan.frx":062A
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton exme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "pplan.frx":1190
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Kalender schliessen"
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox p1 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Zeige nur"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "bis"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "von"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu menu_proj 
      Caption         =   "Projekte"
      Visible         =   0   'False
      Begin VB.Menu menu_proj_open 
         Caption         =   "öffne Projekt"
      End
      Begin VB.Menu menu_proj_open_f 
         Caption         =   "öffne Projekt - Finanzen"
      End
      Begin VB.Menu menu_proj_open_ue 
         Caption         =   "öffne Projekt - Übersicht"
      End
   End
End
Attribute VB_Name = "pplan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p1mm As Boolean
Dim dbpara$
Dim mnams$(1 To 12), wdays$(7), mdays(1 To 12) As Integer
Dim hotx%(99), hoty%(99), hot$(99), hotcount%

Private Sub allclose_Click()
'd2infile = "pplan": d2insub = "allclose_Click"
Call form1.Form_DblClick
End Sub

Private Sub Command1_Click()
'd2infile = "pplan": d2insub = "Command1_Click"
Call Command1_Click_mminus
Call refr_Click
End Sub
Private Sub Command1_Click_mminus()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command1_Click_mminus"
M% = Val(Text1(0).text)
Y% = Val(Mid$(Text1(0).text, InStr(Text1(0).text, ".") + 1))
M% = M% - 1
If M% < 1 Then
  M% = 12
  Y% = Y% - 1
End If
Text1(0).text = trm(M%) & "." & trm(Y%)
End Sub

Private Sub Command2_Click()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command2_Click"
M% = Val(Text1(0).text)
Y% = Val(Mid$(Text1(0).text, InStr(Text1(0).text, ".") + 1))
Y% = Y% - 1
Text1(0).text = trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command3_Click()
'd2infile = "pplan": d2insub = "Command3_Click"
Call Command3_Click_mplus
Call refr_Click

End Sub
Private Sub Command3_Click_mplus()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command3_Click_mplus"
M% = Val(Text1(1).text)
Y% = Val(Mid$(Text1(1).text, InStr(Text1(1).text, ".") + 1))
M% = M% + 1
If M% > 12 Then
  M% = 1
  Y% = Y% + 1
End If
Text1(1).text = trm(M%) & "." & trm(Y%)

End Sub

Private Sub Command4_Click()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command4_Click"
M% = Val(Text1(1).text)
Y% = Val(Mid$(Text1(1).text, InStr(Text1(1).text, ".") + 1))
Y% = Y% + 1
Text1(1).text = trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command5_Click()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command5_Click"
M% = Val(Text1(0).text)
Y% = Val(Mid$(Text1(0).text, InStr(Text1(0).text, ".") + 1))
Y% = Y% + 1
Text1(0).text = trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command6_Click_mplus()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command6_Click_mplus"
M% = Val(Text1(0).text)
Y% = Val(Mid$(Text1(0).text, InStr(Text1(0).text, ".") + 1))
M% = M% + 1
If M% > 12 Then
  M% = 1
  Y% = Y% + 1
End If
Text1(0).text = trm(M%) & "." & trm(Y%)

End Sub
Private Sub Command6_Click()
'd2infile = "pplan": d2insub = "Command6_Click"
Call Command6_Click_mplus
Call refr_Click

End Sub

Private Sub Command7_Click_mminus()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command7_Click_mminus"
M% = Val(Text1(1).text)
Y% = Val(Mid$(Text1(1).text, InStr(Text1(1).text, ".") + 1))
M% = M% - 1
If M% < 1 Then
  M% = 12
  Y% = Y% - 1
End If
Text1(1).text = trm(M%) & "." & trm(Y%)

End Sub
Private Sub Command7_Click()
'd2infile = "pplan": d2insub = "Command7_Click"
Call Command7_Click_mminus
Call refr_Click

End Sub

Private Sub Command8_Click()
Dim M%, Y%

'd2infile = "pplan": d2insub = "Command8_Click"
M% = Val(Text1(1).text)
Y% = Val(Mid$(Text1(1).text, InStr(Text1(1).text, ".") + 1))
Y% = Y% - 1
Text1(1).text = trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub


Private Sub exme_Click()
'd2infile = "pplan": d2insub = "exme_Click"
Unload Me

End Sub

Private Sub Form_Load()
Dim Y%, M%
'd2infile = "pplan": d2insub = "Form_Load"
hotcount% = 0
p1mm = False
mnams$(1) = transe("Januar")
mnams$(2) = transe("Februar")
mnams$(3) = transe("März")
mnams$(4) = transe("April")
mnams$(5) = transe("Mai")
mnams$(6) = transe("Juni")
mnams$(7) = transe("Juli")
mnams$(8) = transe("August")
mnams$(9) = transe("September")
mnams$(10) = transe("Oktober")
mnams$(11) = transe("November")
mnams$(12) = transe("Dezember")
mdays(1) = 31
mdays(2) = 28
mdays(3) = 31
mdays(4) = 30
mdays(5) = 31
mdays(6) = 30
mdays(7) = 31
mdays(8) = 31
mdays(9) = 30
mdays(10) = 31
mdays(11) = 30
mdays(12) = 31

wdays$(0) = "Mo"
wdays$(1) = "Di"
wdays$(2) = "Mi"
wdays$(3) = "Do"
wdays$(4) = "Fr"
wdays$(5) = "Sa"
wdays$(6) = "So"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Width = form1.mylastwidth(Me.name, 1)
Me.Height = form1.mylastheight(Me.name, 1)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Top = Me.Height / 4
  Me.Width = 11200
  Me.Left = Me.Width / 3
End If
Call form1.formpos(Me)
Y% = apyear(now())
M% = apmonth(now()) - 1
If M% < 1 Then
  M% = 12
  Y% = Y% - 1
End If
Text1(0).text = trm(M%) & "." & trm(Y%)
Text1(1).text = apmonth(now()) & "." & apyear(now()) + 1
p1.AutoRedraw = True
pplan.Caption = form1.inmylanguage("Projekte")
mallminus.ToolTipText = form1.inmylanguage("Einen Monat zurück")
mallplus.ToolTipText = form1.inmylanguage("Einen Monat vor")
allclose.ToolTipText = form1.inmylanguage("Schliesst alle Formulare")
refr.ToolTipText = form1.inmylanguage("Ansicht aktualisieren")
exme.ToolTipText = form1.inmylanguage("Kalender schliessen")
Label1(2).Caption = form1.inmylanguage("Zeige nur")
Label1(1).Caption = form1.inmylanguage("bis")
Label1(0).Caption = form1.inmylanguage("von")
menu_proj.Caption = form1.inmylanguage("Projekte")
menu_proj_open.Caption = form1.inmylanguage("öffne Projekt")
menu_proj_open_f.Caption = form1.inmylanguage("öffne Projekt - Finanzen")
menu_proj_open_ue.Caption = form1.inmylanguage("öffne Projekt - Übersicht")
Show

End Sub

Private Sub Form_Resize()
'd2infile = "pplan": d2insub = "Form_Resize"
If Width < 2000 Then Width = 2000
If Height < 2000 Then Height = 2000
exme.Top = Me.Height - exme.Height - 460
allclose.Top = exme.Top
mallplus.Top = exme.Top
mallminus.Top = exme.Top
Label2.Top = exme.Top
p1.Width = Me.Width - 160 - p1.Left
p1.Height = exme.Top - p1.Top
p1.ScaleHeight = 1000
p1.ScaleWidth = 1000
Call refr_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "pplan": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)
exuld:
On Error GoTo 0

End Sub

Private Sub mallminus_Click()
'd2infile = "pplan": d2insub = "mallminus_Click"
Call Command1_Click_mminus
Call Command7_Click_mminus
Call refr_Click

End Sub

Private Sub mallplus_Click()
'd2infile = "pplan": d2insub = "mallplus_Click"
Command3_Click_mplus
Call Command6_Click_mplus
Call refr_Click
End Sub

Private Sub menu_proj_open_Click()
'd2infile = "pplan": d2insub = "menu_proj_open_Click"
Call p1_DblClick
End Sub

Private Sub menu_proj_open_f_Click()
'd2infile = "pplan": d2insub = "menu_proj_open_f_Click"
Call menu_proj_open_Click
DoEvents
Call tplan.fshow_Click
End Sub

Private Sub menu_proj_open_ue_Click()
'd2infile = "pplan": d2insub = "menu_proj_open_ue_Click"
Call menu_proj_open_Click
DoEvents
Call tplan.Command16_Click
End Sub

Private Sub p1_DblClick()
Dim tpid$

'd2infile = "pplan": d2insub = "p1_DblClick"
  tpid$ = trm(Label2.Caption)
  If Len(tpid$) <> 0 Then
    Load tplan
    Call tplan.rlists
    Call tplan.nulldsp
    Call tplan.showrec(tpid$)
    On Error Resume Next
    Call tplan.SetFocus
    On Error GoTo 0
  End If


End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'd2infile = "pplan": d2insub = "p1_MouseDown"
If Button = 2 Then
  PopupMenu menu_proj
End If

End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i%, mind As Double, dst As Double, dx As Double, dy As Double, mi%
Dim c$, r As ADODB.Recordset, tt$, rrr

Dim d2infile As String, d2insub As String
d2infile = "pplan": d2insub = "p1_MouseMove"
If p1mm = True Then Exit Sub
p1mm = True
mind = p1.ScaleHeight * p1.ScaleHeight + p1.ScaleWidth * p1.ScaleWidth
mi% = -1
For i% = 0 To hotcount% - 1
  'dx = x - hotx%(i%)
  dx = 0
  dy = Y - hoty%(i%)
  dst = dx * dx + dy * dy
  If dst < mind Then
    mind = dst
    mi% = i%
  End If
Next i%
If mi% > -1 Then
  If Label2.Caption <> hot$(mi%) Then
    Label2.Caption = hot$(mi%)
'    tt$ = hot$(mi%)
    tt$ = ""
    c$ = "select veranstalter,solist,orchester,dirigent,von,bis from tplan where id='" & hot$(mi%) & "';"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr = 0 Then
    If Not r.EOF Then
      c$ = trm(r!von): If c$ <> "" Then tt$ = datfromsql(c$)
      c$ = trm(r!bis): If c$ <> "" Then tt$ = tt$ & " - " & datfromsql(c$)
      c$ = trm(r!orchester): If c$ <> "" Then tt$ = tt$ & ", " & c$
      c$ = trm(r!dirigent): If c$ <> "" Then tt$ = tt$ & ", " & c$
      c$ = trm(r!veranstalter): If c$ <> "" Then tt$ = tt$ & ", " & c$
      c$ = trm(r!Solist): If c$ <> "" Then tt$ = tt$ & ", " & c$
    End If
    End If
    If trm(tt$) = "" Then tt$ = hot$(mi%)
    p1.ToolTipText = tt$
    DoEvents
  End If
End If
p1mm = False
End Sub

Private Sub refr_Click()
Dim sdt$, edt$, c$, r As ADODB.Recordset, Y%, M%, rrr, i%
Dim anzmd%, amd%, emd%, dmd%, y0%, anzr%, col As Long, whr$, rbis$, rvon$


Dim d2infile As String, d2insub As String
d2infile = "pplan": d2insub = "refr_Click"
sdt$ = datum2sql("01." & Text1(0).text)
On Error Resume Next
edt$ = datum2sql(mdays(Val(Text1(1).text)) & "." & Text1(1).text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
If IsDate(sdt$) = False Then Exit Sub
If IsDate(edt$) = False Then Exit Sub
anzmd% = (apyear(edt$) - apyear(sdt$)) * 12
anzmd% = (anzmd% + apmonth(edt$) - apmonth(sdt$)) + 1
If anzmd% <= 0 Then Exit Sub
dmd% = p1.ScaleWidth / anzmd%
whr$ = "where ((Hauptperson<>'Dekade') and ( " + _
     "((von>='" & sdt$ & "')  and (von<='" & edt$ & "')) " + _
     "or ((bis>='" & sdt$ & "')  and (bis<='" & edt$ & "')) " + _
     "or ((von<'" & sdt$ & "')  and (bis>'" & edt$ & "')) " + _
     ")) "
If trm(Text2.text) <> "" Then whr$ = whr$ & " and (instr(id,'" & Text2.text & "')>0)"
whr$ = whr$
c$ = "select count(*) as cnt from tplan " & whr$
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
anzr% = imin(r!cnt, 100)
hotcount% = 0
r.Close
If anzr% > 0 Then

p1.Cls: y0% = p1.ScaleHeight / (anzr% + 1)
c$ = "select id,von,bis,hauptperson from tplan " & whr$ & "order by id"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
i% = 0
On Error GoTo refr_legende
While Not r.EOF And hotcount% < 100
  rvon$ = trm(r!von): If rvon$ = "" Then rvon$ = datum2sql(Date)
  rbis$ = trm(r!bis): If rbis$ = "" Then rbis$ = rvon$
  amd% = (apyear(rvon$) - apyear(sdt$)) * 12
  amd% = (amd% + apmonth(rvon$) - apmonth(sdt$))
  emd% = (apyear(rbis$) - apyear(sdt$)) * 12
  emd% = (emd% + apmonth(rbis$) - apmonth(sdt$))
  If amd% < 0 Then amd% = 0
  col = form1.projektfarbe(trm(r!hauptperson))
  p1.Line ((emd% + 1) * dmd%, y0% * (i% + 2))-(amd% * dmd%, y0% * (i% + 1)), col, BF
  p1.Print r!id
  hotx%(hotcount%) = amd% * dmd% + Abs(((emd% + 1) * dmd%) - (amd% * dmd%)) / 2
  hoty%(hotcount%) = y0% * (i% + 1) + Abs((y0% * (i% + 2)) - (y0% * (i% + 1))) / 2
  hot$(hotcount%) = r!id
  hotcount% = hotcount% + 1
  i% = i% + 1
  r.MoveNext
Wend

End If
refr_legende:
On Error GoTo 0
Y% = apyear(sdt$)
M% = apmonth(sdt$)
For i% = 0 To anzmd%
  p1.Line (i% * dmd%, 0)-(i% * dmd%, p1.Height)
  p1.Line ((i% + 1) * dmd%, y0%)-(i% * dmd%, 0), col, BF
  p1.Print M%; "/"; Right(trm(Y%), 2)
  M% = M% + 1
  If M% > 12 Then
    M% = 1
    Y% = Y% + 1
  End If
Next i%
End Sub

Private Sub Text1_DblClick(Index As Integer)
Dim p$

'd2infile = "pplan": d2insub = "Text1_DblClick"
  p$ = Text1(Index).text
  With frmCalendar
    .init Text1(Index), Text1(Index).text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1(Index).text = Format(.SelectedDate, "mm.yyyy")
    End If
  End With
  Unload frmCalendar
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'd2infile = "pplan": d2insub = "Text2_KeyDown"
If KeyCode = 13 Then Call refr_Click

End Sub
