VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form kc 
   Caption         =   "Kalenderkontrollfeld"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton yvw 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1440
      Picture         =   "kc.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   29
      ToolTipText     =   "Jahresübersicht"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox priv 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox immer1 
      Height          =   255
      Left            =   1920
      TabIndex        =   25
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox prz 
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1080
      Picture         =   "kc.frx":0462
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Kalender öffnen"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CheckBox dkz 
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   2220
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "A"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      ToolTipText     =   "Alle Filter löschen, alles anzeigen"
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6840
      Picture         =   "kc.frx":05EC
      Style           =   1  'Grafisch
      TabIndex        =   18
      ToolTipText     =   "Kalender schliessen"
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6480
      Picture         =   "kc.frx":083C
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Schliesst alle Formulare"
      Top             =   2520
      Width           =   375
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
      TabIndex        =   16
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command7 
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
      Height          =   195
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   495
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
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1920
      Picture         =   "kc.frx":0E66
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Zum 1. des jetzigen Monats"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command5 
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
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command4 
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
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   6840
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
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
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Adresse dazu"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Beenden 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Picture         =   "kc.frx":0F66
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Kalender schliessen"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      Picture         =   "kc.frx":11B6
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
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
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Adresse entfernen"
      Top             =   1680
      Width           =   255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   1680
      TabIndex        =   6
      Top             =   0
      Width           =   5535
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
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   375
   End
   Begin VB.ListBox selct 
      Height          =   645
      Index           =   0
      Left            =   0
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.ListBox selct 
      Height          =   645
      Index           =   1
      Left            =   0
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox selct 
      Height          =   840
      Index           =   2
      Left            =   240
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "priv."
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "immer am 1."
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proj."
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dekaden"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   2220
      Width           =   855
   End
End
Attribute VB_Name = "kc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mnams$(1 To 12), wdays$(7), wkspl%
Public yyyy0%

Private Sub Beenden_Click()
'd2infile = "kc": d2insub = "Beenden_Click"
Unload kc
End Sub

Public Function getwho() As String
Dim s$, i%

'd2infile = "kc": d2insub = "getwho"
s$ = ""
For i% = 0 To selct(1).ListCount - 1
  If selct(1).Selected(i%) = True Then
    If Len(s$) = 0 Then
      s$ = "|" + selct(1).List(i%) + "|"
    Else
      s$ = s$ + selct(0).List(i%) + "|"
    End If
  End If
Next i%

getwho = s$

End Function

Public Function getwhere() As String
Dim s$, i%, shwpriv As Boolean, typ$
'd2infile = "kc": d2insub = "getwhere"
s$ = ""
shwpriv = False
For i% = 0 To selct(0).ListCount - 1
  typ$ = transo(selct(0).List(i%))
  If selct(0).Selected(i%) = True Then
    If Len(s$) = 0 Then
      s$ = "where ((auftritt.auftrittstyp='" + typ$ + "') "
    Else
      s$ = s$ + "or (auftritt.auftrittstyp='" + typ$ + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"

getwhere = s$

End Function

Private Sub Combo2_Click()
Dim d As Variant, rrr, tg As Integer

ttform.Hide
DoEvents
If Combo2.ListIndex >= 0 And Combo3.ListIndex >= 0 And Val(Text1.text) > 0 Then
  tg = Val(Text1.text)
Call form1.dbg2f("tg=" + trm(tg))
  Do
    If form1.getusersetting("datumsformat", "de") = "de" Then
      On Error Resume Next
      d = CDate(trm(tg) & "." & (Combo2.ListIndex + 1) & "." & (Combo3.ListIndex + yyyy0))
      rrr = Err
      On Error GoTo 0
    Else
      On Error Resume Next
      d = CDate(trm(tg) & "/" & (Combo2.ListIndex + 1) & "/" & (Combo3.ListIndex + yyyy0))
      rrr = Err
      On Error GoTo 0
    End If
Call form1.dbg2f("d=" + trm(d))
    If rrr <> 0 And tg > 28 Then tg = tg - 1
  Loop Until rrr = 0 Or tg < 29
  If apyear(d) > 9000 Then
    tg = Val(Text1.text) - 1
    If tg <= 0 Then tg = 1
    Text1.text = trm(tg)
    DoEvents
    Exit Sub
  End If
  If rrr = 0 Then
    List1.Clear
Call form1.dbg2f("calling settag0:" + trm(d) + " (" + datum2sql(d) + ")")
    Call k3.settag0(datum2sql(d))
' das ist ein don't!
'    On Error Resume Next
'    Call k3.SetFocus
'    On Error GoTo 0
  End If
End If
End Sub

Private Sub Combo3_Click()
'd2infile = "kc": d2insub = "Combo3_Click"
Call Combo2_Click
End Sub

Public Sub Command1_Click()
'd2infile = "kc": d2insub = "Command1_Click"
Call Combo2_Click
End Sub

Private Sub Command15_Click()
'd2infile = "kc": d2insub = "Command15_Click"
Call settag0(Date)
Text1.text = 1
End Sub

Private Sub Command18_Click()

'd2infile = "kc": d2insub = "Command18_Click"
Call form1.handbuchcall("10-Termine.htm")

End Sub

Private Sub Command2_Click()
'd2infile = "kc": d2insub = "Command2_Click"
On Error Resume Next
selct(2).RemoveItem selct(2).ListIndex
On Error GoTo 0
End Sub


Private Sub Command27_Click()
'd2infile = "kc": d2insub = "Command27_Click"
Call form1.Form_DblClick

End Sub

Private Sub Command29_Click()
Dim d

'd2infile = "kc": d2insub = "Command29_Click"
If form1.getusersetting("datumsformat", "de") = "de" Then
  d = CDate(Text1.text & "." & (Combo2.ListIndex + 1) & "." & (Combo3.ListIndex + yyyy0))
Else
  d = CDate(Text1.text & "/" & (Combo2.ListIndex + 1) & "/" & (Combo3.ListIndex + yyyy0))
End If
Load dayvw
On Error Resume Next
Call dayvw.SetFocus
On Error GoTo 0
dayvw.Text1.text = trm(d)

End Sub

Private Sub Command3_Click()
Dim sel$

Load adrselect
adrselect.SetFocus
Call adrselect.sel_init("", "")
MousePointer = 11
While adrselect.sel_valid() = 0: DoEvents: Wend
MousePointer = 0
sel$ = adrselect.sel_getselected()
selct(2).AddItem sel$

End Sub

Public Sub Command4_Click()
Dim M%, i%, Y

'd2infile = "kc": d2insub = "Command4_Click"
For i% = 1 To 12
  If mnams$(i%) = Combo2.text Then
    M% = i%
    i% = 12
  End If
Next i%
M% = M% - 1
If M% < 1 Then
  M% = 12
  If Combo3.ListIndex + 1 > 0 Then Combo3.ListIndex = Combo3.ListIndex - 1
End If
Combo2.ListIndex = M% - 1
Combo2.text = mnams$(M%)
'Call Combo2_Click
End Sub

Public Sub Command5_Click()
Dim M%, i%, Y

'd2infile = "kc": d2insub = "Command5_Click"
For i% = 1 To 12
  If mnams$(i%) = Combo2.text Then
    M% = i%
    i% = 12
  End If
Next i%
M% = M% + 1
If M% > 12 Then
  M% = 1
  If Combo3.ListIndex + 1 < Combo3.ListCount Then Combo3.ListIndex = Combo3.ListIndex + 1
End If
Combo2.ListIndex = M% - 1
Combo2.text = mnams$(M%)
'Call Combo2_Click

End Sub

Public Sub Command6_Click()
If Combo3.ListIndex + 1 < Combo3.ListCount Then Combo3.ListIndex = Combo3.ListIndex + 1
End Sub

Public Sub Command7_Click()

If Combo3.ListIndex + 1 > 1 Then Combo3.ListIndex = Combo3.ListIndex - 1

End Sub

Private Sub Command8_Click()
Call Beenden_Click
End Sub

Private Sub Command9_Click()

Dim i%
For i% = 0 To selct(0).ListCount - 1
  selct(0).Selected(i%) = False
Next i%
For i% = 0 To selct(1).ListCount - 1
  selct(1).Selected(i%) = False
Next i%
While selct(2).ListCount > 0
  selct(2).RemoveItem 0
Wend
Call Command1_Click
End Sub

Private Sub dkz_Click()

If dkz.value = 0 Then
  Call form1.setusersetting("Dekadenzeigen", "nein")
Else
  Call form1.setusersetting("Dekadenzeigen", "ja")
End If
Call Combo2_Click

End Sub

Private Sub Form_Load()
Dim t$, i%, s%
Dim mew, meh

axsResizer1.SaveControlPositions

wkspl% = 1
List1.Clear
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Left = form1.Left + form1.Width + 20
End If
mew = form1.mylastwidth(Me.name, 0)
meh = form1.mylastheight(Me.name, 0)
If meh > 0 And mew > 0 Then
  Me.Width = mew
  Me.Height = meh
End If
Call form1.formpos(Me)
s% = form1.myfontsize()
List1.Font.Size = s%
selct(2).Font.Size = s%

t$ = form1.inmylanguage(form1.myfirstdayofweek())
mnams$(1) = form1.inmylanguage("Januar")
mnams$(2) = form1.inmylanguage("Februar")
mnams$(3) = form1.inmylanguage("März")
mnams$(4) = form1.inmylanguage("April")
mnams$(5) = form1.inmylanguage("Mai")
mnams$(6) = form1.inmylanguage("Juni")
mnams$(7) = form1.inmylanguage("Juli")
mnams$(8) = form1.inmylanguage("August")
mnams$(9) = form1.inmylanguage("September")
mnams$(10) = form1.inmylanguage("Oktober")
mnams$(11) = form1.inmylanguage("November")
mnams$(12) = form1.inmylanguage("Dezember")

wdays$(0) = form1.inmylanguage("Mo")
wdays$(1) = form1.inmylanguage("Di")
wdays$(2) = form1.inmylanguage("Mi")
wdays$(3) = form1.inmylanguage("Do")
wdays$(4) = form1.inmylanguage("Fr")
wdays$(5) = form1.inmylanguage("Sa")
wdays$(6) = form1.inmylanguage("So")
Me.Caption = form1.inmylanguage("Kalenderkontrollfeld")
yvw.ToolTipText = transe("Jahresübersicht")
Beenden.ToolTipText = transe("Kalender schliessen")
Command8.ToolTipText = transe("Kalender schliessen")
Command4.ToolTipText = transe("einen") + " " + transe("Monat") + " " + transe("zurück")
Command5.ToolTipText = transe("einen") + " " + transe("Monat") + " " + transe("voraus")
Command7.ToolTipText = transe("ein") + " " + transe("Jahr") + " " + transe("zurück")
Command27.ToolTipText = transe("Schliesst alle Formulare")
Command6.ToolTipText = transe("ein") + " " + transe("Jahr") + " " + transe("vor")
Command1.ToolTipText = transe("Ansicht aktualisieren")
Command15.ToolTipText = transe("Zum 1. des jetzigen Monats")
Command3.ToolTipText = transe("Adresse dazu")
Command2.ToolTipText = transe("Adresse entfernen")
Command9.ToolTipText = transe("Alle Filter löschen, alles zeigen")
Command29.ToolTipText = transe("Tageskalender öffnen")
Label2.Caption = transe("immer am 1.")
Show
Combo2.Clear
Combo2.text = ""
For i% = 1 To 12
  Combo2.AddItem mnams$(i%)
Next i%
If form1.getusersetting("Dekadenzeigen", "nein") = "ja" Then
  dkz.value = 1
Else
  dkz.value = 0
End If
If form1.getusersetting("Projektezeigen", "nein") = "ja" Then
  prz.value = 1
Else
  prz.value = 0
End If
If form1.getusersetting("Privateszeigen", "nein") = "ja" Then
  priv.value = 1
Else
  priv.value = 0
End If
If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then
  immer1.value = 1
Else
  immer1.value = 0
End If
Combo3.Clear
Combo3.text = ""
yyyy0 = 1940
For i% = yyyy0 To 2070
  Combo3.AddItem i%
Next i%

Call k3.setnogoto(1)
Load k3
Call rlist1
Call k3.setnogoto(0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i$, t%, Index As Integer, tw%

Unload k3
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)

Index = 0
For t% = 0 To selct(Index).ListCount - 1
  i$ = selct(Index).List(t%)
  tw% = selct(Index).Selected(t%)
  Call form1.setmylastFormVar(Me.name, transo(i$), trm(tw%))
Next t%

exuld:
On Error GoTo 0
End Sub
Sub rlist1()
Dim rtmp As ADODB.Recordset, sme%, i%, rrr

Dim d2infile As String, d2insub As String
d2infile = "kc": d2insub = "rlist1"
selct(0).Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftrittstypen order by sortierung", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  selct(0).AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend
For i% = 0 To selct(0).ListCount - 1
  sme% = CInt(form1.mylastFormVar(Me.name, transo(selct(0).List(i%)), "0"))
  If sme% <> -1 Then sme% = 0
  selct(0).Selected(i%) = sme%
Next i%

selct(1).Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppenindex", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not rtmp.EOF
  selct(1).AddItem rtmp!id
  rtmp.MoveNext
Wend

End Sub

Public Sub settag0(d$)
'd2infile = "kc": d2insub = "settag0"
  Call form1.dbg2f("kc.settag(" + d$ + ")")
  Call form1.dbg2f("cdate(d)=" + trm(CDate(d$)))
  Combo2.ListIndex = apmonth(CDate(d$)) - 1
  On Error Resume Next
  Combo3.ListIndex = apyear(CDate(d$)) - yyyy0
  On Error GoTo 0

  Text1.text = apday(CDate(d$))
End Sub

Private Sub immer1_Click()
'd2infile = "kc": d2insub = "immer1_Click"
If immer1.value = 0 Then
  Call form1.setusersetting("kalenderimmeramersten", "nein")
Else
  Call form1.setusersetting("kalenderimmeramersten", "ja")
End If
Call Combo2_Click

End Sub

Private Sub Label1_Click()
'd2infile = "kc": d2insub = "Label1_Click"
If prz.value = 0 Then
  prz.value = 1
Else
  prz.value = 0
End If

End Sub

Private Sub Label2_Click()
'd2infile = "kc": d2insub = "Label2_Click"
If immer1.value = 0 Then
  immer1.value = 1
Else
  immer1.value = 0
End If

End Sub

Private Sub Label3_Click()
If priv.value = 0 Then
  priv.value = 1
Else
  priv.value = 0
End If

End Sub

Private Sub Label7_Click()
'd2infile = "kc": d2insub = "Label7_Click"
If dkz.value = 0 Then
  dkz.value = 1
Else
  dkz.value = 0
End If

End Sub

Public Sub List1_DblClick()
Dim id$, prj As Boolean

'd2infile = "kc": d2insub = "List1_DblClick"
prj = False
If List1.ListIndex < 0 Then Exit Sub
id$ = List1.List(List1.ListIndex)
If InStr(id$, " Projekt: ") > 0 Then prj = True
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)

If Not prj Then
  Unload auftritt
  DoEvents
  Load auftritt
  On Error Resume Next
  Call auftritt.SetFocus
  On Error GoTo 0
  Call auftritt.showrec(id$, 0)
Else
        On Error Resume Next
        Unload tplan
        On Error GoTo 0
        DoEvents
        Load tplan
        Call tplan.rlists
        Call tplan.nulldsp
        Call tplan.showrec(id$)
        On Error Resume Next
        Call tplan.SetFocus
        On Error GoTo 0
End If

End Sub

Private Sub priv_Click()
If priv.value = 0 Then
  Call form1.setusersetting("Privateszeigen", "nein")
Else
  Call form1.setusersetting("Privateszeigen", "ja")
End If
Call Combo2_Click

End Sub

Private Sub prz_Click()

If prz.value = 0 Then
  Call form1.setusersetting("Projektezeigen", "nein")
Else
  Call form1.setusersetting("Projektezeigen", "ja")
End If
Call Combo2_Click

End Sub

Private Sub selct_Click(Index As Integer)
Dim i$, t%

'd2infile = "kc": d2insub = "selct_Click"
If Index = 0 Then
  t% = selct(0).ListIndex
  If t% < 0 Then Exit Sub
  i$ = selct(0).List(t%)
  t% = selct(0).Selected(t%)
  Call form1.setmylastFormVar(Me.name, transo(i$), trm(t%))
End If

End Sub

Private Sub Text1_Change()
'd2infile = "kc": d2insub = "Text1_Change"
Call Combo2_Click
End Sub
Private Sub Form_Resize()
'd2infile = "kc": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub yvw_Click()
Load ky
On Error Resume Next
Call ky.SetFocus
On Error GoTo 0
End Sub
