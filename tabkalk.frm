VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tabkalk 
   Caption         =   "Tabellenkalkulation"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   ScaleHeight     =   4980
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      Picture         =   "tabkalk.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "Beendet die Kalkulation"
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CSV"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   7920
      Picture         =   "tabkalk.frx":0C42
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Vorlage (neu) öffnen"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "tabkalk.frx":126C
      Style           =   1  'Grafisch
      TabIndex        =   20
      ToolTipText     =   "Neue Kalkulation anlegen"
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   6600
      Sorted          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Andere Vorlage laden oder Name der Vorlage beim ""Speichern als Vorlage"""
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Alle Änderungen verwerfen"
      Top             =   4320
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid fg4 
      Height          =   1455
      Left            =   4200
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2566
      _Version        =   393216
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
      Height          =   375
      Left            =   360
      TabIndex        =   16
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tabkalk.frx":15FE
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Als Vorlage speichern"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Picture         =   "tabkalk.frx":1AF0
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "löschen dieser Kalkulation"
      Top             =   3240
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid fg3 
      Height          =   2895
      Left            =   7320
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tabkalk.frx":2DC6
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Speichern"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tabkalk.frx":330A
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Neu berechnen, Ansicht aktualisieren"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   0
      Picture         =   "tabkalk.frx":3E70
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   4560
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   4800
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   3975
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      _Version        =   393216
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Datenherkunft zeigen"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label currcursor 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "tabkalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currx As Integer, curry As Integer, funcweiter$, brokenkalk As Boolean

Private Sub Check1_Click()

'd2infile = "tabkalk": d2insub = "Check1_Click"
fg3.Top = fg2.Top
fg3.Left = fg2.Left
If Check1.value = 1 Then
  fg3.Visible = True
Else
  fg3.Visible = False
End If

End Sub

Private Sub Check2_Click()

'd2infile = "tabkalk": d2insub = "Check2_Click"
If Check2.value = 0 Then
  delme.Enabled = False
Else
  delme.Enabled = True
End If

End Sub

Private Sub combo1_Change()
'd2infile = "tabkalk": d2insub = "combo1_Change"
If Combo1.Text = "Standard" Then Combo1.Text = ""

End Sub

Private Sub Combo1_Click()

'd2infile = "tabkalk": d2insub = "Combo1_Click"
Call combo1_Change
Call Command29_Click

End Sub

Private Sub Combo1_DropDown()
Dim tr, l%, fn$

'd2infile = "tabkalk": d2insub = "Combo1_DropDown"
fn$ = "tabkalk_" & trm(Label1.Caption) & "__" & Label3.Caption
l% = Len(fn$)
fn$ = form1.s0dir() & "\" & form1.getdbname() & ".rtf\" & fn$ & "*.tbl"
Combo1.Clear
tr = Dir(fn$)
While tr <> ""
  fn$ = Mid$(basename(trm(tr), ".tbl"), l% + 1): If fn$ = "" Then fn$ = "Standard"
  While Left$(fn$, 1) = "_": fn$ = Mid$(fn$, 2): Wend
  Combo1.AddItem fn$
  tr = Dir
Wend

End Sub

Public Sub Command1_Click()

'd2infile = "tabkalk": d2insub = "Command1_Click"
Unload Me

End Sub

Private Sub Command11_Click()
'd2infile = "tabkalk": d2insub = "Command11_Click"
Call nulldsp

End Sub

Public Sub Command2_Click()
Dim cmd$, tb0$, l$, X%, Y%, typ$, i%, V$, c$, tbl$, fld$, kid$
Dim anbase%

'd2infile = "tabkalk": d2insub = "Command2_Click"
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
If Label3.Caption = "" Then Exit Sub
kid$ = Label2.Caption
tb0$ = "tabkalk_" & typ$ & "_" & Label3.Caption
cmd$ = "delete from auftritthigru where auftrittsid='" & Label2.Caption & "' and auftrittstyp='" & tb0$ & "'"
Call form1.sqlqry(cmd$)
For Y% = 0 To fg2.Rows - 1
  l$ = ""
  For X% = 0 To fg2.Cols - 1
    If fg4.TextMatrix(Y%, X%) = "=blocked" Then
      l$ = l$ & "|"
    Else
    l$ = l$ & fg3.TextMatrix(Y%, X%) & "|"
    If Left(fg3.TextMatrix(Y%, X%), 1) = "=" And InStr(fg3.TextMatrix(Y%, X%), "__") > 0 And Not isfunc(fg3.TextMatrix(Y%, X%)) Then
      If fg4.TextMatrix(Y%, X%) <> fg3.TextMatrix(Y%, X%) Then
        V$ = Mid$(fg3.TextMatrix(Y%, X%), 2)
        i% = InStr(V$, "__")
        If i% > 0 Then
          tbl$ = Left$(V$, i% - 1)
          fld$ = Mid$(V$, i% + 2)
          Select Case tbl$
            Case "auftritthigru":
              c$ = "delete from auftritthigru where feldname='" & fld$ & "' and" & _
                   " auftrittstyp='" & typ$ & "' and" & _
                   " auftrittsid='" & kid$ & "'"
              Call form1.sqlqry(c$)
              c$ = "insert into auftritthigru (id,felddaten,feldname,auftrittstyp,auftrittsid) values('" & _
                   form1.newid("auftritthigru", "id", 52) & "','" & fg4.TextMatrix(Y%, X%) & _
                   "','" & fld$ & "','" & typ$ & "','" & kid$ & "')"
              Call form1.sqlqry(c$)
              c$ = "update usr_" & utabn(typ$) & " set " & fld$ & "='" & fg4.TextMatrix(Y%, X%) & "' where" & _
                   " id='" & kid$ & "'"
              Call form1.sqlqry(c$)
            Case "finanzen":
              c$ = "update finanzen set " & fld$ & "='" & fg4.TextMatrix(Y%, X%) & "' where id='" & kid$ & "'"
              Call form1.sqlqry(c$)
            Case "auftritt":
              c$ = "update auftritt set " & fld$ & "='" & fg4.TextMatrix(Y%, X%) & "' where id='" & kid$ & "'"
              Call form1.sqlqry(c$)
            Case Else:
          End Select
        End If
      End If
    End If
    End If
  Next X%
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
      form1.newid("auftritthigru", "id", 50) & "','" & _
      Label2.Caption & "','" & _
      tb0$ & "','" & _
      Format$(Y%, "000###") & "','" & _
      l$ & "')"
  Call form1.sqlqry(cmd$)
Next Y%
If typ$ <> "Projekt" Then
  cmd$ = "delete from auftritthigru where auftrittsid='" & Label2.Caption & "' and auftrittstyp='" & Label1.Caption & "' and feldname='" & Label3.Caption & "'"
  Call form1.sqlqry(cmd$)
  cmd$ = "insert into auftritthigru (id,feldname,felddaten,auftrittsid,auftrittstyp) values('" & _
                 form1.newid("auftritthigru", "id", 50) & "','" & _
                 Label3.Caption & "','" & _
                 Me.Caption & "','" & _
                 Label2.Caption & "','" & _
                 Label1.Caption & "')"
  Call form1.sqlqry(cmd$)

  cmd$ = "update usr_" & utabn(Label1.Caption) _
                 & " set " & Label3.Caption & "='" & Me.Caption _
                 & "' where id='" & Label2.Caption & "'"
  Call form1.sqlqry(cmd$)
  anbase% = auftritt.nbase
  Call auftritt.showrec(Label2.Caption, 0)
  Call auftritt.clearlabels
  Call auftritt.initfields(typ$, anbase%, 0)
Else
  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
      form1.newid("auftritthigru", "id", 50) & "','" & _
      Label2.Caption & "','" & _
      tb0$ & "'," & _
      "'999999','" & _
      Me.Caption & "')"
  Call form1.sqlqry(cmd$)
  Call tplan.rkalklist
End If

Me.BackColor = form1.cleancolor
On Error Resume Next
Me.SetFocus
On Error GoTo 0

End Sub

Private Sub Command29_Click()
Dim typ$, c1$, vorlage$

'd2infile = "tabkalk": d2insub = "Command29_Click"
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
c1$ = "": If trm(Combo1.Text) <> "" Then c1$ = "__" & trm(Combo1.Text)
If LCase(c1$) = "standard" Then c1$ = ""
vorlage$ = form1.s0dir() & "\" & form1.getdbname() & ".rtf\tabkalk_" & typ$ & "__" & Label3.Caption & c1$ & ".tbl"
If nexist(vorlage$) Then
  MsgBox transe("Die Vorlage existiert nicht.")
  Exit Sub
End If
Call nulldsp
Call openvorlage(trm(Combo1.Text))
Call recalc
Me.BackColor = form1.cleancolor

End Sub

Private Sub Command3_Click()
Dim vorlage$, o%, Y%, X%, l$, typ$, antw, c1$

'd2infile = "tabkalk": d2insub = "Command3_Click"
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
c1$ = "": If trm(Combo1.Text) <> "" Then c1$ = "__" & trm(Combo1.Text)
If LCase(c1$) = "standard" Then c1$ = ""
vorlage$ = form1.s0dir() & "\" & form1.getdbname() & ".rtf\tabkalk_" & typ$ & "__" & Label3.Caption & c1$ & ".tbl"
If Not nexist(vorlage$) Then
  antw = MsgBox(transe("Die Vorlage existiert bereits, überschreiben?"), vbYesNo + vbCritical + vbDefaultButton2, "Vorlage ersetzen?")
  If antw <> vbYes Then Exit Sub
End If
MousePointer = 11: DoEvents
o% = FreeFile
Open vorlage$ For Output As #o%
For Y% = 0 To fg2.Rows - 1
  l$ = ""
  For X% = 0 To fg2.Cols - 1
    If fg4.TextMatrix(Y%, X%) = "=blocked" Then
      l$ = l$ & "|"
    Else
      l$ = l$ & fg3.TextMatrix(Y%, X%) & "|"
    End If
  Next X%
  Print #o%, l$
Next Y%

Close #o%
MousePointer = 0

End Sub

Private Sub Command4_Click()
'd2infile = "tabkalk": d2insub = "Command4_Click"
Call recalc

End Sub

Private Sub Command5_Click()
'd2infile = "tabkalk": d2insub = "Command5_Click"
Call Label1_Change
End Sub

Private Sub Command6_Click()
Dim fn$, o%, X%, Y%

'd2infile = "tabkalk": d2insub = "Command6_Click"
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) = "" Then Exit Sub
o% = FreeFile
Open fn$ For Output As #o%
For Y% = 0 To fg2.Rows - 1
  If fg2.TextMatrix(Y%, 0) = "ExportEnde" Then Exit For
  For X% = 0 To fg2.Cols - 1
    If fg2.TextMatrix(0, X%) = "ExportEnde" Then Exit For
    Print #o%, """" & fg2.TextMatrix(Y%, X%) & """";
    If X% < fg2.Rows - 1 Then Print #o%, ",";
  Next X%
  Print #o%,
Next Y%
Close #o%

End Sub

Private Sub Command7_Click()
'd2infile = "tabkalk": d2insub = "Command7_Click"
brokenkalk = True
End Sub

Private Sub delme_Click()
Dim cmd$, tb0$, typ$

'd2infile = "tabkalk": d2insub = "delme_Click"
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
tb0$ = "tabkalk_" & typ$ & "_" & Label3.Caption
cmd$ = "delete from auftritthigru where auftrittsid='" & Label2.Caption & "' and auftrittstyp='" & tb0$ & "'"
Call form1.sqlqry(cmd$)
If typ$ = "Projekt" Then Call tplan.rkalklist

Unload Me

End Sub

Private Sub fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If currx < 0 Or curry < 0 Or currx >= fg2.Cols Or curry >= fg2.Rows Then
  Exit Sub
End If
fg2.col = currx
fg2.Row = curry
fg2.CellBackColor = RGB(0, 0, 0)

End Sub

Private Sub fg2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xm%, ym%, rrr, tx%, z$

'd2infile = "tabkalk": d2insub = "fg2_MouseUp"
z$ = Right(Text1.Text, 1)
If z$ <> "" And InStr(funcweiter$, z$) > 0 Then
  Text1.Text = Text1.Text & "[" & trm(fg2.MouseCol) & "," & trm(fg2.MouseRow) & "]"
  Call Text1.SetFocus
  Text1.SelStart = Len(Text1.Text)
  Exit Sub
End If
Text1.Enabled = True
Text1.Text = ""
xm% = fg2.MouseCol
ym% = fg2.MouseRow
fg2.col = xm%
fg2.Row = ym%
fg2.CellBackColor = RGB(192, 192, 192)
Label4.Caption = fg3.TextMatrix(ym%, xm%)
Text1.Text = fg2.TextMatrix(ym%, xm%)
currcursor.Caption = "S" & trm(xm%) & ",Z" & trm(ym%)
currx = xm%: curry = ym%
If Left(Label4.Caption, 1) = "=" Then
  Text1.Text = Label4.Caption
End If
If xm% = fg2.Cols - 1 Then
  fg2.Cols = fg2.Cols + 1
  fg3.Cols = fg3.Cols + 1
  fg4.Cols = fg4.Cols + 1
  fg2.ColWidth(fg2.Cols - 1) = fg2.ColWidth(fg2.Cols - 1) / 2
End If
If ym% = fg2.Rows - 2 Then
  fg2.Rows = fg2.Rows + 1
  fg3.Rows = fg3.Rows + 1
  fg4.Rows = fg4.Rows + 1
  For tx% = 0 To fg2.Cols - 1
    fg2.TextMatrix(fg2.Rows - 1, tx%) = fg2.TextMatrix(fg2.Rows - 2, tx%)
    fg2.TextMatrix(fg2.Rows - 2, tx%) = ""
    fg3.TextMatrix(fg2.Rows - 1, tx%) = fg3.TextMatrix(fg2.Rows - 2, tx%)
    fg3.TextMatrix(fg2.Rows - 2, tx%) = ""
    fg4.TextMatrix(fg2.Rows - 1, tx%) = fg4.TextMatrix(fg2.Rows - 2, tx%)
    fg4.TextMatrix(fg2.Rows - 2, tx%) = ""
  Next tx%
End If
On Error Resume Next
Call Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
On Error GoTo 0

End Sub

Private Sub Form_Load()

'd2infile = "tabkalk": d2insub = "Form_Load"
axsResizer1.SaveControlPositions

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Me.BackColor = form1.cleancolor()
funcweiter$ = "+*-/=(;"
tabkalk.Caption = transe("Tabellenkalkulation")
Command7.ToolTipText = transe("Beendet die Kalkulation")
Command6.Caption = transe("CSV")
Command29.ToolTipText = transe("Vorlage (neu) öffnen")
Command11.ToolTipText = transe("Neue Kalkulation anlegen")
Combo1.ToolTipText = transe("Andere Vorlage laden oder Name der Vorlage beim <Speichern als Vorlage>")
Command5.Caption = transe("Esc")
Command5.ToolTipText = transe("Alle Änderungen verwerfen")
Command18.Caption = transe("?")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command3.ToolTipText = transe("Als Vorlage speichern")
delme.ToolTipText = transe("löschen dieser Kalkulation")
Command2.ToolTipText = transe("Speichern")
Command4.ToolTipText = transe("Neu berechnen, Ansicht aktualisieren")
Command1.ToolTipText = transe("Formular schliessen")
Label5.Caption = transe("Datenherkunft zeigen")
Show

End Sub
Private Sub Form_Resize()
'd2infile = "tabkalk": d2insub = "Form_Resize"
axsResizer1.Resize
fg3.Top = fg2.Top
fg3.Left = fg2.Left
fg3.Width = fg2.Width
fg3.Height = fg2.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim antw

'd2infile = "tabkalk": d2insub = "Form_Unload"
If Me.BackColor = form1.dirtycolor() Then
  antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  If antw = vbYes Then Call Command2_Click
End If

Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Sub nulldsp()
Dim X%, Y%

'd2infile = "tabkalk": d2insub = "nulldsp"
fg2.Clear
fg3.Clear
fg4.Clear
fg2.Cols = 3
fg2.Rows = 3
fg3.Cols = 3
fg3.Rows = 3
fg3.TextMatrix(fg3.Rows - 1, 0) = Label3.Caption
fg4.Cols = 3
fg4.Rows = 3

For X% = 0 To fg2.Cols - 1: fg2.ColWidth(X%) = fg2.ColWidth(X%) / 2: Next X%

End Sub

Sub openvorlage(vn$)
Dim typ$, id$, vorlage$, o%, l$, ml%, tb0$, Y%, X%
Dim fg2cntx%, fg2cnty%, chsiz, cmd$, vx$

'd2infile = "tabkalk": d2insub = "openvorlage"
chsiz = 89
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
tb0$ = "tabkalk_" & typ$ & "_" & Label3.Caption
vx$ = "": If vn$ <> "" Then vx$ = "__" & vn$
  vorlage$ = form1.s0dir() & "\" & form1.getdbname() & ".rtf\tabkalk_" & typ$ & "__" & Label3.Caption & vx$ & ".tbl"
  If Not nexist(vorlage$) Then
    o% = FreeFile
    Open vorlage$ For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      tb0$ = trm(l$)
      If Y% = fg3.Rows Then
        fg3.Rows = fg3.Rows + 1
        fg4.Rows = fg4.Rows + 1
        fg2.Rows = fg2.Rows + 1
      End If
      While trm(tb0$) <> ""
        l$ = cut_d1(tb0$, "|")
        tb0$ = Mid$(tb0$, Len(l$) + 2)
        If X% = fg3.Cols Then
          fg3.Cols = fg3.Cols + 1
          fg4.Cols = fg4.Cols + 1
          fg2.Cols = fg2.Cols + 1
          fg2.ColWidth(fg2.Cols - 1) = fg2.ColWidth(fg2.Cols - 1) / 2
        End If
        fg3.TextMatrix(Y%, X%) = l$
        If Left(l$, 1) = "=" And InStr(l$, "__") > 0 And Not isfunc(l$) Then fg4.TextMatrix(Y%, X%) = l$
        X% = X% + 1
      Wend
      Y% = Y% + 1
      X% = 0
    Wend
    Close #o%
  End If

End Sub
Sub showrec()
Dim rrr
Dim typ$, id$, vorlage$, o%, l$, ml%, tb0$, Y%, X%
Dim fg2cntx%, fg2cnty%, chsiz, r As ADODB.Recordset, cmd$

Dim d2infile As String, d2insub As String
d2infile = "tabkalk": d2insub = "showrec"
chsiz = 89
typ$ = trm(Label1.Caption): If typ$ = "" Then Exit Sub
tb0$ = "tabkalk_" & typ$ & "_" & Label3.Caption
cmd$ = "select * from auftritthigru where auftrittsid='" & Label2.Caption & "' and auftrittstyp='" & tb0$ & "' order by feldname"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Y% = 0: X% = 0
  While Not r.EOF
    If r!feldname <> "999999" Then
    tb0$ = trm(r!felddaten)
    If Y% = fg3.Rows Then
      fg4.Rows = fg4.Rows + 1
      fg3.Rows = fg3.Rows + 1
      fg2.Rows = fg2.Rows + 1
    End If
    While trm(tb0$) <> ""
      l$ = cut_d1(tb0$, "|")
      tb0$ = Mid$(tb0$, Len(l$) + 2)
      If X% = fg3.Cols Then
        fg3.Cols = fg3.Cols + 1
        fg4.Cols = fg4.Cols + 1
        fg2.Cols = fg2.Cols + 1
        fg2.ColWidth(fg2.Cols - 1) = fg2.ColWidth(fg2.Cols - 1) / 2
      End If
      fg3.TextMatrix(Y%, X%) = l$
      If Left(l$, 1) = "=" And InStr(l$, "__") > 0 And Not isfunc(l$) Then fg4.TextMatrix(Y%, X%) = l$
      X% = X% + 1
    Wend
    Y% = Y% + 1
    X% = 0
    End If
    r.MoveNext
  Wend
Else
  Call openvorlage("")
End If
Call recalc
Me.BackColor = form1.cleancolor
End Sub

Private Sub Label1_Change()

'd2infile = "tabkalk": d2insub = "Label1_Change"
Call nulldsp
If Label1.Caption = "" Then Exit Sub
Label2.Visible = False
If Label1.Caption = "Projekt" Then
  Label2.Visible = True
End If

Call showrec

End Sub


Private Sub Label3_Change()

'd2infile = "tabkalk": d2insub = "Label3_Change"
If trm(Label3.Caption) = "" Then
  Command2.Enabled = False
Else
  Command2.Enabled = True
End If

End Sub

Private Sub Label5_Click()
'd2infile = "tabkalk": d2insub = "Label5_Click"
If Check1.value = 0 Then
  Check1.value = 1
Else
  Check1.value = 0
End If

End Sub

Private Sub Text1_Change()
'd2infile = "tabkalk": d2insub = "Text1_Change"
If InStr(Text1.Text, "|") > 0 Then
  Text1.Text = strrepl(Text1.Text, "|", ":")
  Text1.SelStart = Len(Text1.Text)
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

'd2infile = "tabkalk": d2insub = "Text1_KeyDown"
If KeyCode = 27 Then
  Call recalc
  Text1.Text = fg2.TextMatrix(curry, currx)
End If

If currx < 0 Or curry < 0 Then Exit Sub

If KeyCode = 13 Then
  fg2.col = currx
  fg2.Row = curry
  fg2.CellBackColor = RGB(0, 0, 0)
  Call entercelldata
End If
'Me.Caption = KeyCode & "-" & Shift
End Sub

Sub entercelldata()
Dim z$

'd2infile = "tabkalk": d2insub = "entercelldata"
z$ = Right(Text1.Text, 1)
If z$ <> "" And InStr(funcweiter$, z$) > 0 Then
  Exit Sub
End If
  If curry < 0 Or currx < 0 Then Exit Sub
  If fg3.TextMatrix(curry, currx) = Text1.Text Then Exit Sub
  Me.BackColor = form1.dirtycolor()
  If fg4.TextMatrix(curry, currx) = "" Then
    fg4.TextMatrix(curry, currx) = Text1.Text
    fg2.TextMatrix(curry, currx) = varwertermittlung(Mid(Text1.Text, 2))
  End If
  If Left(fg3.TextMatrix(curry, currx), 1) = "=" And InStr(fg3.TextMatrix(curry, currx), "__") > 0 And Not isfunc(fg3.TextMatrix(curry, currx)) Then
    fg4.TextMatrix(curry, currx) = Text1.Text
  Else
    fg3.TextMatrix(curry, currx) = Text1.Text
  End If
  If Left(Text1.Text, 1) <> "=" Then
    fg2.TextMatrix(curry, currx) = Text1.Text
  End If
  If Text1.Text = "=delete" Then
    fg2.TextMatrix(curry, currx) = ""
    fg3.TextMatrix(curry, currx) = ""
    fg4.TextMatrix(curry, currx) = ""
  End If
  Text1.Text = ""
  currx = -1: curry = -1
  Call recalc
End Sub

Function wertaus(formel$) As String
Dim rrr
Dim fkt$, dlist$, i%, f$, fp%, p%, fktnam$, fktargs$, c$
Dim arg1$, arg2$, X%, Y%, argx$, arg3$, rs As ADODB.Recordset
Dim ssum As Double

Dim d2infile As String, d2insub As String
d2infile = "tabkalk": d2insub = "wertaus"
wertaus = "N/A"
If Left$(formel$, 1) <> "=" Then
  wertaus = formel$
  Exit Function
End If
fkt$ = formel$

While Left$(fkt$, 1) = "=": fkt$ = Mid$(fkt$, 2): Wend
dlist$ = "+-*/(": i% = Len(dlist$)
f$ = ""
While i% > 0
  p% = InStr(fkt$, Mid$(dlist, i%, 1))
  If p% > 0 Then
    f$ = Mid$(dlist, i%, 1)
    fktnam$ = Left$(fkt$, p% - 1)
    fktargs$ = Mid$(fkt$, p% + 1)
    While Right$(fktargs$, 1) = ")": fktargs$ = Left$(fktargs$, Len(fktargs$) - 1): Wend
    fp% = p%
    i% = 0
  End If
  i% = i% - 1
Wend
If f$ = "" Then
  wertaus = varwertermittlung(fkt$)
  Exit Function
End If
If f$ = "(" Then
  If LCase(fktnam$) = "fix2" Then
    If InStr(fktargs$, "[") > 0 Then fktargs$ = deref(fktargs$)
    On Error Resume Next
    wertaus = trm(fixeur(var2dbl(word1(strrepl(fktargs$, ".", ",")))))
    On Error GoTo 0
    Exit Function
  End If
  If LCase(fktnam$) = "int" Then
    If InStr(fktargs$, "[") > 0 Then fktargs$ = deref(fktargs$)
    On Error Resume Next
    wertaus = trm(Int(var2dbl(word1(strrepl(fktargs$, ".", ",")))))
    On Error GoTo 0
    Exit Function
  End If
  If LCase(fktnam$) = "terminbyid" Then
    arg1$ = cut_d1(fktargs$, ";"): arg2$ = cut_d2bis(fktargs$, ";")
    arg2$ = cut_d1(arg2$, ";")
    If InStr(arg1$, "[") > 0 Then arg1$ = deref(arg1$)
    If InStr(arg2$, "[") > 0 Then arg2$ = deref(arg2$)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    c$ = "SELECT felddaten FROM auftritthigru where auftrittsid='" & arg2$ & "' and " & _
         "feldname='" & arg1$ & "';"
rrr = form1.adoopen(rs, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rs.EOF Then
      wertaus = trm(rs!felddaten)
      Exit Function
    End If
  End If
  If LCase(fktnam$) = "terminkopfbyid" Then
    arg1$ = cut_d1(fktargs$, ";"): arg2$ = cut_d2bis(fktargs$, ";")
    arg2$ = cut_d1(arg2$, ";")
    If InStr(arg1$, "[") > 0 Then arg1$ = deref(arg1$)
    If InStr(arg2$, "[") > 0 Then arg2$ = deref(arg2$)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    c$ = "SELECT " & arg1$ & " as felddaten FROM auftritt where id='" & arg2$ & "';"
rrr = form1.adoopen(rs, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rs.EOF Then
      wertaus = trm(rs!felddaten)
      Exit Function
    End If
  End If
  If LCase(fktnam$) = "terminliste" Then
    If InStr(fktargs$, "[") > 0 Then fktargs$ = deref(fktargs$)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    c$ = "SELECT auftritt.ID as aid FROM auftritt INNER JOIN tplan ON auftritt.TourneeplanID = tplan.ID " & _
         "Where (((tplan.id)='" & Label2.Caption & "') And ((auftritt.auftrittstyp)='" & fktargs$ & "'))" & _
         "ORDER BY auftritt.Datum;"
rrr = form1.adoopen(rs, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    fktargs$ = ""
    While Not rs.EOF
      If Len(fktargs$) > 0 Then fktargs$ = fktargs$ & "|"
      fktargs$ = fktargs$ & rs!aid
      rs.MoveNext
    Wend
    wertaus = "Liste::" & fktargs$
    Exit Function
  End If
End If
If LCase(fktnam$) = "cut" Then
  'was cutten
  arg1$ = cut_d1(fktargs$, ";"): arg3$ = cut_d2bis(fktargs$, ";")
  'wo cutten
  arg2$ = cut_d1(arg3$, ";")
  'wievielte stelle
  arg3$ = cut_d2bis(arg3$, ";")

  arg1$ = varwertermittlung(arg1$)
  i% = Val(varwertermittlung(arg3$))
  While i% > 0 And arg1$ <> ""
    argx$ = cut_d1(arg1$, arg2$)
    arg1$ = cut_d2bis(arg1$, arg2$)
    i% = i% - 1
  Wend
  wertaus = argx$
  Exit Function
End If
If LCase(fktnam$) = "spaltensumme" Then
  arg1$ = cut_d1(fktargs$, ";")
  arg2$ = Mid$(fktargs$, Len(arg1$) + 2)
  X% = zellwert1(arg1$)
  i% = zellwert2(arg1$)
  Y% = zellwert2(arg2$)
  ssum = 0
  While i% <= Y%
      On Error Resume Next
      argx$ = word1(strrepl(fg2.TextMatrix(i%, X%), ".", "")): If trm(argx$) = "" Then argx$ = "0"
      ssum = ssum + var2dbl(argx$)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        wertaus = "N/A"
        Exit Function
      End If
      i% = i% + 1
  Wend
  wertaus = ssum
  Exit Function
End If
If LCase(fktnam$) = "zeilensumme" Then
  arg1$ = cut_d1(fktargs$, ";")
  arg2$ = Mid$(fktargs$, Len(arg1$) + 2)
  Y% = zellwert2(arg1$)
  i% = zellwert1(arg1$)
  X% = zellwert1(arg2$)
  ssum = 0
  While i% <= X%
      On Error Resume Next
      argx$ = word1(strrepl(fg2.TextMatrix(Y%, i%), ".", ",")): If trm(argx$) = "" Then argx$ = "0"
      ssum = ssum + var2dbl(argx$)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        wertaus = "N/A"
        Exit Function
      End If
      i% = i% + 1
  Wend
  wertaus = ssum
  Exit Function
End If

  arg1$ = Left$(fkt$, InStr(fkt$, f$) - 1)
  arg2$ = Mid$(fkt$, InStr(fkt$, f$) + 1)
  arg1$ = varwertermittlung(arg1$)
  arg2$ = varwertermittlung(arg2$)
  Select Case f$
    Case "+"
        ssum = var2dbl(word1(strrepl(arg1$, ".", ""))) + var2dbl(word1(strrepl(arg2$, ".", "")))
    Case "*"
        ssum = var2dbl(word1(strrepl(arg1$, ".", ""))) * var2dbl(word1(strrepl(arg2$, ".", "")))
    Case "-"
        ssum = var2dbl(word1(strrepl(arg1$, ".", ""))) - var2dbl(word1(strrepl(arg2$, ".", "")))
    Case "/"
        ssum = var2dbl(word1(strrepl(arg1$, ".", ""))) / var2dbl(word1(strrepl(arg2$, ".", "")))
    Case Else
        ssum = "0"
  End Select
  wertaus = ssum

End Function

Function deref(V$) As String
Dim X%, Y%

'd2infile = "tabkalk": d2insub = "deref"
deref = V$
X% = zellwert1(V$)
Y% = zellwert2(V$)
deref = fg2.TextMatrix(Y%, X%)

End Function
Function zellwert1(V$) As Integer
Dim z$
'd2infile = "tabkalk": d2insub = "zellwert1"
While Left$(V$, 1) = "[": V$ = Mid$(V$, 2): Wend
z$ = Left$(V$, InStr(V$, ",") - 1)
zellwert1 = Val(z$)
End Function

Function zellwert2(V$) As Integer
Dim z$
'd2infile = "tabkalk": d2insub = "zellwert2"
z$ = Mid$(V$, InStr(V$, ",") + 1)
While Right$(V$, 1) = "]": V$ = Left$(V$, Len(V$) - 1): Wend
zellwert2 = Val(z$)
End Function

Sub recalc()
Dim X%, Y%, neuwert As String, chgfl As Boolean
Dim chsiz, ml%, lstc%, c$, tx%, rereadlists As Boolean

'd2infile = "tabkalk": d2insub = "recalc"
chsiz = 88
fg3.TextMatrix(fg3.Rows - 1, 0) = Label3.Caption
Command7.Top = Command2.Top
Command7.Left = Command2.Left
Command7.Visible = True
brokenkalk = False: rereadlists = True
Do
chgfl = False
For Y% = 0 To fg3.Rows - 1
  If fg3.TextMatrix(Y%, 0) = "insert" Then
    fg3.TextMatrix(Y%, 0) = ""
    Call insline(Y%)
    chgfl = True
    Exit For
  End If
  If fg3.TextMatrix(Y%, 0) = "delete" Then
    Call delline(Y%)
    chgfl = True
    Exit For
  End If
  For X% = 0 To fg3.Cols - 1
    If fg3.TextMatrix(0, X%) = "insert" Then
      fg3.TextMatrix(0, X) = ""
      Call inscol(X%)
      chgfl = True
      Exit For
    End If
    If fg3.TextMatrix(0, X%) = "delete" Then
      Call delcol(X%)
      chgfl = True
      Exit For
    End If
    fg2.col = X%
    fg2.Row = Y%
    If trm(fg3.TextMatrix(Y%, X%)) <> "" Then
      Call form1.dbg2f("recalc: (" & trm(X%) & "," & trm(Y%) & ")" & fg3.TextMatrix(Y%, X%))
    End If
    If Left(fg3.TextMatrix(Y%, X%), 1) = "=" And InStr(fg3.TextMatrix(Y%, X%), "__") > 0 And Not isfunc(fg3.TextMatrix(Y%, X%)) Then
      fg2.CellForeColor = RGB(0, 0, 255)
    Else
      fg2.CellForeColor = RGB(0, 0, 0)
    End If
    DoEvents
    If Left(fg3.TextMatrix(Y%, X%), 1) = "=" And InStr(fg3.TextMatrix(Y%, X%), "__") > 0 And Not isfunc(fg3.TextMatrix(Y%, X%)) And _
         fg3.TextMatrix(Y%, X%) <> fg4.TextMatrix(Y%, X%) Then
      neuwert = fg4.TextMatrix(Y%, X%)
    Else
      neuwert = wertaus(fg3.TextMatrix(Y%, X%))
      If Left$(neuwert, 7) = "Liste::" Then
        If rereadlists Then
          lstc% = 1
          While Y% + lstc% < fg2.Rows
            fg2.TextMatrix(Y% + lstc%, X%) = ""
            fg3.TextMatrix(Y% + lstc%, X%) = ""
            fg4.TextMatrix(Y% + lstc%, X%) = ""
            lstc% = lstc% + 1
          Wend
        End If
        lstc% = 1
        neuwert = Mid$(neuwert, 8)
        While trm(neuwert) <> ""
          c$ = cut_d1(neuwert, "|")
          neuwert = cut_d2bis(neuwert, "|")
          fg2.TextMatrix(Y% + lstc%, X%) = c$
          fg3.TextMatrix(Y% + lstc%, X%) = c$
          fg4.TextMatrix(Y% + lstc%, X%) = "=blocked"
          lstc% = lstc% + 1
If Y% + lstc% = fg2.Rows - 2 Then
  fg2.Rows = fg2.Rows + 1
  fg3.Rows = fg3.Rows + 1
  fg4.Rows = fg4.Rows + 1
  For tx% = 0 To fg2.Cols - 1
    fg2.TextMatrix(fg2.Rows - 1, tx%) = fg2.TextMatrix(fg2.Rows - 2, tx%)
    fg2.TextMatrix(fg2.Rows - 2, tx%) = ""
    fg3.TextMatrix(fg2.Rows - 1, tx%) = fg3.TextMatrix(fg2.Rows - 2, tx%)
    fg3.TextMatrix(fg2.Rows - 2, tx%) = ""
    fg4.TextMatrix(fg2.Rows - 1, tx%) = fg4.TextMatrix(fg2.Rows - 2, tx%)
    fg4.TextMatrix(fg2.Rows - 2, tx%) = ""
  Next tx%
End If
        Wend
        neuwert = "Liste: " & Mid$(fg3.TextMatrix(Y%, X%), InStr(fg3.TextMatrix(Y%, X%), "(") + 1)
        neuwert = Left(neuwert, Len(neuwert) - 1)
      End If
    End If
    ml% = (Len(neuwert) + 1) * chsiz
    If ml% > fg2.ColWidth(X%) Then
      fg2.ColWidth(X%) = ml%
    End If
    If neuwert <> fg2.TextMatrix(Y%, X%) Then
      fg2.TextMatrix(Y%, X%) = neuwert
      chgfl = True
      DoEvents
    End If
  Next X%
Next Y%
rereadlists = False
Loop Until chgfl = False Or brokenkalk

If brokenkalk Then chgfl = False
Command7.Visible = False
neuwert = ""
For X% = 1 To fg2.Cols - 1
  If fg2.TextMatrix(fg2.Rows - 1, X%) <> "" Then
    neuwert = neuwert & fg2.TextMatrix(fg2.Rows - 1, X%)
  End If
Next X%
Me.Caption = neuwert
End Sub

Private Function varwertermittlung(V$) As String
Dim i%, tbl$, fld$, X%, Y%, c$, kid$, ktyp$, r As ADODB.Recordset, rrr, op1
Dim ssum

Dim d2infile As String, d2insub As String
d2infile = "tabkalk": d2insub = "varwertermittlung"
kid$ = Label2.Caption
ktyp$ = Label1.Caption
If Left$(V$, 1) = "[" Then
  varwertermittlung = deref(V$)
  Exit Function
End If
If V$ = "rnd" Then
  varwertermittlung = trm(Int(Rnd * 32768))
  Exit Function
End If
i% = InStr(V$, "__")
If i% > 0 Then
  tbl$ = Left$(V$, i% - 1)
  fld$ = Mid$(V$, i% + 2)
  c$ = ""
  Select Case tbl$
    Case "auftritt":
      c$ = "select " & fld$ & " from " & tbl$ & " where id='" & kid$ & "'"
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
        'varwertermittlung = word1(strrepl(form1.ohnewaehrung(trm(r.Fields(0).value)), ".", ""))
        varwertermittlung = trm(r.Fields(0).value)
        Exit Function
      Else
        varwertermittlung = "0"
        Exit Function
      End If
    End If
  End If
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

Private Sub Text1_LostFocus()
'd2infile = "tabkalk": d2insub = "Text1_LostFocus"
Call entercelldata
End Sub

Function isfunc(arg$) As Boolean
Dim dlist$, i%, p%

'd2infile = "tabkalk": d2insub = "isfunc"
isfunc = False
dlist$ = "+-*/(": i% = Len(dlist$)
While i% > 0
  p% = InStr(arg$, Mid$(dlist, i%, 1))
  If p% > 0 Then i% = 0
  i% = i% - 1
Wend
If p% > 0 Then isfunc = True

End Function

Sub delline(nr%)
Dim i%, X%, erg$

'd2infile = "tabkalk": d2insub = "delline"
For i% = 0 To fg3.Rows - 1
  For X% = 0 To fg3.Cols - 1
    If fg3.TextMatrix(i%, X%) <> "" And i% <> nr% Then
      erg$ = fixyformellinedel(fg3.TextMatrix(i%, X%), nr%)
      If erg$ <> fg3.TextMatrix(i%, X%) Then
        fg3.TextMatrix(i%, X%) = erg$
      End If
    End If
  Next X%
Next i%
For i% = nr% To fg3.Rows - 2
  For X% = 0 To fg3.Cols - 1
    fg3.TextMatrix(i%, X%) = fg3.TextMatrix(i% + 1, X%)
    fg4.TextMatrix(i%, X%) = fg4.TextMatrix(i% + 1, X%)
  Next X%
Next i%
fg3.Rows = fg3.Rows - 1
fg4.Rows = fg4.Rows - 1
fg2.Rows = fg2.Rows - 1
End Sub

Sub insline(nr%)
Dim i%, X%, erg$

'd2infile = "tabkalk": d2insub = "insline"
fg3.Rows = fg3.Rows + 1
fg4.Rows = fg4.Rows + 1
fg2.Rows = fg2.Rows + 1
For i% = fg3.Rows - 1 To 0 Step -1
  For X% = fg3.Cols - 1 To 0 Step -1
    If fg3.TextMatrix(i%, X%) <> "" Then
      erg$ = fixyformellineins(fg3.TextMatrix(i%, X%), nr%)
      If erg$ <> fg3.TextMatrix(i%, X%) Then
        fg3.TextMatrix(i%, X%) = erg$
      End If
    End If
  Next X%
Next i%
For i% = fg3.Rows - 1 To nr% + 1 Step -1
  For X% = 0 To fg3.Cols - 1
    fg3.TextMatrix(i%, X%) = fg3.TextMatrix(i% - 1, X%)
    fg4.TextMatrix(i%, X%) = fg4.TextMatrix(i% - 1, X%)
  Next X%
Next i%
For X% = 0 To fg3.Cols - 1
  fg3.TextMatrix(nr%, X%) = ""
  fg4.TextMatrix(nr%, X%) = ""
Next X%
End Sub
Sub inscol(nr%)
Dim i%, Y%, erg$

'd2infile = "tabkalk": d2insub = "inscol"
fg3.Cols = fg3.Cols + 1
fg4.Cols = fg4.Cols + 1
fg2.Cols = fg2.Cols + 1
For i% = fg3.Cols - 1 To 0 Step -1
  For Y% = 0 To fg3.Rows - 1
    If fg3.TextMatrix(Y%, i%) <> "" Then
      erg$ = fixxformellineins(fg3.TextMatrix(Y%, i%), nr%)
      If erg$ <> fg3.TextMatrix(Y%, i%) Then
        fg3.TextMatrix(Y%, i%) = erg$
      End If
    End If
  Next Y%
Next i%

For i% = fg3.Cols - 1 To nr% + 1 Step -1
  For Y% = 0 To fg3.Rows - 1
    fg3.TextMatrix(Y%, i%) = fg3.TextMatrix(Y%, i% - 1)
    fg4.TextMatrix(Y%, i%) = fg4.TextMatrix(Y%, i% - 1)
  Next Y%
Next i%
For Y% = 0 To fg3.Rows - 1
  fg3.TextMatrix(Y%, nr%) = ""
  fg4.TextMatrix(Y%, nr%) = ""
Next Y%

End Sub


Sub delcol(nr%)
Dim i%, Y%, erg$

'd2infile = "tabkalk": d2insub = "delcol"
For i% = 0 To fg3.Cols - 1
  For Y% = 0 To fg3.Rows - 1
    If fg3.TextMatrix(Y%, i%) <> "" And i% <> nr% Then
      erg$ = fixxformellinedel(fg3.TextMatrix(Y%, i%), nr%)
      If erg$ <> fg3.TextMatrix(Y%, i%) Then
        fg3.TextMatrix(Y%, i%) = erg$
      End If
    End If
  Next Y%
Next i%

For i% = nr% To fg3.Cols - 2
  For Y% = 0 To fg3.Rows - 1
    fg3.TextMatrix(Y%, i%) = fg3.TextMatrix(Y%, i% + 1)
    fg4.TextMatrix(Y%, i%) = fg4.TextMatrix(Y%, i% + 1)
  Next Y%
Next i%
fg3.Cols = fg3.Cols - 1
fg4.Cols = fg4.Cols - 1
fg2.Cols = fg2.Cols - 1
End Sub

Function fixyformellinedel(l$, n%) As String
Dim p%, l1$

'd2infile = "tabkalk": d2insub = "fixyformellinedel"
l1$ = l$

      p% = InStr(l1$, "[")
      If p% > 0 Then
        'l1$ = strrepl(l1$, "," & trm(n%) & "]", "," & trm(n% - 1) & "]")
        For p% = n% To fg3.Rows - 1
          l1$ = strrepl(l1$, "," & trm(p%) & "]", "," & trm(p% - 1) & "]")
        Next p%
      End If

fixyformellinedel = l1$

End Function

Function fixyformellineins(l$, n%) As String
Dim p%, l1$

'd2infile = "tabkalk": d2insub = "fixyformellineins"
l1$ = l$

      p% = InStr(l1$, "[")
      If p% > 0 Then
        'l1$ = strrepl(l1$, "," & trm(n%) & "]", "," & trm(n% - 1) & "]")
        For p% = fg3.Rows - 2 To n% Step -1
          l1$ = strrepl(l1$, "," & trm(p%) & "]", "," & trm(p% + 1) & "]")
        Next p%
      End If

fixyformellineins = l1$

End Function

Function fixxformellinedel(l$, n%) As String
Dim p%, l1$

'd2infile = "tabkalk": d2insub = "fixxformellinedel"
l1$ = l$

      p% = InStr(l1$, "[")
      If p% > 0 Then
        'l1$ = strrepl(l1$, "," & trm(n%) & "]", "," & trm(n% - 1) & "]")
        For p% = n% To fg3.Cols - 1
          l1$ = strrepl(l1$, "[" & trm(p%) & ",", "[" & trm(p% - 1) & ",")
        Next p%
      End If

fixxformellinedel = l1$

End Function

Function fixxformellineins(l$, n%) As String
Dim p%, l1$

'd2infile = "tabkalk": d2insub = "fixxformellineins"
l1$ = l$

      p% = InStr(l1$, "[")
      If p% > 0 Then
        'l1$ = strrepl(l1$, "," & trm(n%) & "]", "," & trm(n% - 1) & "]")
        For p% = fg3.Cols - 2 To n% Step -1
          l1$ = strrepl(l1$, "[" & trm(p%) & ",", "[" & trm(p% + 1) & ",")
        Next p%
      End If

fixxformellineins = l1$

End Function

