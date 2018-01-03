VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form dayvw 
   Caption         =   "Tageskalender"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "dayvw.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "Terminliste senden"
      Top             =   720
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
      Left            =   480
      TabIndex        =   22
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   6360
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6360
      Picture         =   "dayvw.frx":00B2
      Style           =   1  'Grafisch
      TabIndex        =   20
      ToolTipText     =   "Kalender öffnen"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   960
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4920
      TabIndex        =   10
      Top             =   0
      Width           =   615
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
      Left            =   5640
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      Picture         =   "dayvw.frx":01B2
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Kalender öffnen"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      Picture         =   "dayvw.frx":02B2
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox kalimmer 
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Beenden 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "dayvw.frx":0E18
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Kalender schliessen"
      Top             =   6360
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   5415
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9551
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   20
      Left            =   240
      Picture         =   "dayvw.frx":1068
      ToolTipText     =   "drucken"
      Top             =   1320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label statlbl 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   6480
      Width           =   5055
   End
   Begin VB.Label gtoheute 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Heute"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Label jtag 
      BackStyle       =   0  'Transparent
      Caption         =   "Tt"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "beim Starten öffnen"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   6060
      Width           =   1815
   End
   Begin VB.Label wtag 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
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
      Left            =   4200
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6255
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6495
   End
   Begin VB.Menu neu 
      Caption         =   "Neu"
      Visible         =   0   'False
      Begin VB.Menu opnterm 
         Caption         =   "Öffnen"
      End
      Begin VB.Menu subneu 
         Caption         =   "Neu"
         Begin VB.Menu neuterm 
            Caption         =   "Termin für alle"
         End
         Begin VB.Menu neutermprivate 
            Caption         =   "Termin nur für den gewählten Benutzer"
         End
         Begin VB.Menu neutodo 
            Caption         =   "ToDo"
         End
         Begin VB.Menu neuproj 
            Caption         =   "Projekt"
         End
      End
      Begin VB.Menu shwopts 
         Caption         =   "Anzeige"
         Begin VB.Menu thisnoshow 
            Caption         =   "Markierte Termine für gewählten Benutzer nicht zeigen"
         End
      End
   End
End
Attribute VB_Name = "dayvw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim abeg As Integer, aend As Integer, lastcreatedid$, atsnd$
Const LVCTYP = 1
Const LVCBEZ = 2
Const LVCID = 3
Const LVCSTATUS = 4

Private Sub Beenden_Click()

Unload Me

End Sub

Private Sub combo1_Change()

Call Command4_Click
End Sub

Private Sub Combo1_Click()

Call combo1_Change
End Sub

Private Sub Combo1_DropDown()
Dim rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "dayvw": d2insub = "Combo1_DropDown"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM benutzerdaten order by id", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Combo1.Clear
While Not rtmp.EOF
  Combo1.AddItem rtmp!id
  rtmp.MoveNext
Wend

End Sub

Private Sub Command1_Click()

Text1.text = CDate(Text1.text) - 1
End Sub

Private Sub Command15_Click()

Load kc
Call kc.settag0(Text1.text)
'If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then kc.Text1.Text = 1
On Error Resume Next
Call kc.SetFocus
Call k3.SetFocus
On Error GoTo 0

End Sub

Private Sub Command16_Click()
Dim dtg$, dn$, fn$, o%, rrr, X, wot As String, i%, beginn$, ende$, typ As String

dtg$ = datum2sql(Text1.text)
dn$ = atsnd$ + "\" + dtg$ + "-00-00"
fn$ = form1.mymailaddress()
o% = FreeFile
On Error Resume Next
MkDir dn$
Open dn$ + "\" + fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, "Agencyprof Erinnerung Tagestermine " + trm(Text1.text)
  For i% = 1 To gd1.ListItems.Count
    typ = gd1.ListItems(i%).SubItems(1)
    If typ <> "" Then
      wot = gd1.ListItems(i%).SubItems(2)
      beginn$ = gd1.ListItems(i%).text
      If i% < gd1.ListItems.Count Then
        ende$ = gd1.ListItems(i% + 1).text
      Else
        ende$ = "24:00"
      End If
      Print #o%, beginn$ + " - " + ende$ + ": " + typ + ": " + wot
    End If
  Next i%
  Close #o%
  X = Shell("explorer.exe " + dn$, vbNormalFocus)
  X = Shell("notepad.exe " + dn$ + "\" + fn$, vbNormalFocus)
End If

End Sub

Private Sub Command18_Click()

Call form1.handbuchcall("10-Termine.htm")

End Sub

Private Sub Command2_Click()

Call Command15_Click
End Sub

Public Sub Command4_Click()

Call Text1_Change
If Combo1.text <> form1.getuserid() Then
  Me.BackColor = RGB(255, 196, 196)
  statlbl.Caption = transe("NICHT IHR KALENDER")
Else
  Me.BackColor = form1.cleancolor()
  statlbl.Caption = ""
End If

End Sub

Private Sub Command5_Click()

Text1.text = CDate(Text1.text) + 1
End Sub

Private Sub Form_Load()
Dim mew As Integer, meh As Integer
Dim colHeader
Dim lvitem, c$


axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
mew = form1.mylastwidth(Me.name, 0)
meh = form1.mylastheight(Me.name, 0)
If meh > 0 And mew > 0 Then
  Me.Width = mew
  Me.Height = meh
End If
Call form1.formpos(Me)
lastcreatedid$ = ""
Me.Caption = transe("Tageskalender")
Text2.ToolTipText = transe("Kalenderteilung in Minuten")
Label7.Caption = transe("beim Starten öffnen")
If Not nexist(form1.mylocaldatadir() + "\positions\dayvw.aut") Then kalimmer.value = 1
c$ = onlynums(form1.getusersetting("arbeitsbeginn", "0800"))
If Len(c$) <> 4 Then c$ = "0800"
abeg = Val(Left$(c$, 2)) * 60 + Val(Mid$(c$, 3, 2))
c$ = onlynums(form1.getusersetting("arbeitsende", "1800"))
If Len(c$) <> 4 Then c$ = "1800"
aend = Val(Left$(c$, 2)) * 60 + Val(Mid$(c$, 3, 2))
form1.dayvopen = True
Combo1.text = form1.getuserid()

gd1.View = lvwReport
Set colHeader = gd1.ColumnHeaders.add(, , transe("Zeit"), 800)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Typ"), 800)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Bezeichnung"), 3000)
Set colHeader = gd1.ColumnHeaders.add(, , "ID", 1)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Status"), 1000)
atsnd$ = form1.getusersetting("atsend", "")
If atsnd$ = "" Then Command16.Visible = False
gtoheute.ForeColor = RGB(0, 0, 255)
gtoheute.Caption = form1.Label1.Caption
c$ = form1.getusersetting("dayvwteilung", "30")
Text2.text = c$
Command15.ToolTipText = transe("Kalender öffnen")
Command18.ToolTipText = transe("Hilfeseite öfnen")
Command4.ToolTipText = transe("Ansicht aktualisieren")
Command2.ToolTipText = transe("Kalender öffnen")
Beenden.ToolTipText = transe("Kalender schliessen")

Show

End Sub
Private Sub Form_Resize()

axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

form1.dayvopen = False
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)

exuld:
On Error GoTo 0

End Sub

Private Sub gd1_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim id$, c As String, frt As String, tt As String, caltyp As String


If Cancel = 0 Then
  caltyp = cut_d1(gd1.SelectedItem.SubItems(LVCID), ":")
  id$ = cut_d2bis(gd1.SelectedItem.SubItems(LVCID), ":")
  If id$ <> "" Then
    frt = trm(NewString): tt = ""
    If InStr(frt, "-") > 0 Then
      tt = cut_d2bis(frt, "-")
      frt = cut_d1(frt, "-")
    End If
    Select Case caltyp:
      Case "(auftritt":
        c = "update auftritt set zeit='" + frt + "' where id='" + id$ + "';"
        Call form1.sqlqry(c)
        If tt <> "" Then
          c = "delete from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysez' and auftrittstyp='" + form1.auftrittstyp(id$) + "';"
          Call form1.sqlqry(c)
          If tt <> "" Then
            c = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values(" & _
              "'" + form1.newid("auftritthigru", "id", 50) + "','" + _
              id$ + "','" + form1.auftrittstyp(id$) + "','zzzsysez','" + _
              trm(tt) + "');"
            Call form1.sqlqry(c)
          End If
        End If
      Case "(todo"
        c = "update todolist set zeit='" + frt + "' where id='" + id$ + "';"
        Call form1.sqlqry(c)
      Case Else:
        MsgBox "Sorry, nicht möglich für " + caltyp
    End Select
  End If
End If
DoEvents
Call Command4_Click

End Sub

Private Sub gd1_DblClick()
Dim id$, idtyp$

'd2infile = "dayvw": d2insub = "gd1_DblClick"
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem.SubItems(LVCID)
If id$ <> "" Then
  idtyp$ = cut_d1(id$, ":")
  id$ = cut_d2bis(id$, ":")
  Select Case idtyp$
    Case "(auftritt"
        Unload auftritt
        DoEvents
        Load auftritt
        Call auftritt.SetFocus
        Call auftritt.showrec(id$, 0)
    Case "(todo"
        Load todolist
        Call todolist.SetFocus
    Case Else:
        MsgBox idtyp$ + ": " + id$ + vbCrLf + ":-( noch nicht programmiert."
  End Select
End If
End Sub

Private Sub gd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'd2infile = "dayvw": d2insub = "gd1_MouseDown"
If Button = 2 Then
  PopupMenu neu
End If

End Sub

Private Sub gtoheute_Click()
'd2infile = "dayvw": d2insub = "gtoheute_Click"
dayvw.Text1.text = word1(gtoheute.Caption)

End Sub

Private Sub Image1_Click(Index As Integer)

'd2infile = "dayvw": d2insub = "Image1_Click"
If Index = 20 Then Call dayprint

End Sub

Private Sub jtag_Click(Index As Integer)

'd2infile = "dayvw": d2insub = "jtag_Click"
If Index <> 3 Then Text1.text = CDate(Text1.text) - 3 + Index

End Sub

Private Sub kalimmer_Click()
'd2infile = "dayvw": d2insub = "kalimmer_Click"
Call form1.setmyautoopen(Me.name, kalimmer.value)

End Sub

Private Sub neuproj_Click()
Dim neuid As String, s$

'd2infile = "dayvw": d2insub = "neuproj_Click"
neuid = CDate(Text1.text)
neuid = " " & Mid(neuid, 4, 2) & " " & apyear(neuid)
neuid = trm(InputBox(transe("Neue Projekt-ID:"), transe("Neues Projekt erstellen"), neuid))
If trm(neuid) = "" Then Exit Sub

s$ = "insert into tplan (id,kuerzel,hauptperson) values('" + neuid$ + "','" + Left$(neuid$, 4) + "','Orchester')"
Call form1.sqlqry(s$)

On Error Resume Next
Load tplan
tplan.SetFocus
On Error GoTo 0
DoEvents

tplan.Text2.text = neuid$

End Sub

Private Sub neuterm_Click()
Dim d As Variant, id$, i%, beginn$, ende$
Dim s$

'd2infile = "dayvw": d2insub = "neuterm_Click"
  d = CDate(Text1.text)
  beginn$ = "": ende$ = ""
  For i% = 1 To gd1.ListItems.Count
    If gd1.ListItems(i%).Selected Then
      If beginn$ = "" Then beginn$ = gd1.ListItems(i%).text
      If i% < gd1.ListItems.Count Then
        ende$ = gd1.ListItems(i% + 1).text
      Else
        ende$ = "24:00"
      End If
    End If
  Next i%
  If beginn$ = "" Then beginn$ = "00:00"
  If ende$ = "" Then ende$ = "01:00"
  id$ = form1.newid("auftritt", "id", 20)
  lastcreatedid$ = id$
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 id$ + "','-1'" + _
                 ",'Neuer Auftritt','" + transe("Neuer Auftritt") + "','" + _
                 datum2sql(CDate(d)) + "')")
  form1.sqlqry ("update auftritt set zeit='" + beginn$ + "' where id='" + id$ + "';")
  s$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values(" & _
          "'" + form1.newid("auftritthigru", "id", 50) + "','" + _
          id$ + "','Neuer Auftritt','zzzsysez','" + _
          ende$ + "');"
  Call form1.sqlqry(s$)
  Unload auftritt
  DoEvents
  Load auftritt
  Call auftritt.SetFocus
  Call auftritt.showrec(id$, 0)

End Sub

Private Sub neutermprivate_Click()
Dim c$, fnam$

'd2infile = "dayvw": d2insub = "neutermprivate_Click"
Call neuterm_Click
If lastcreatedid$ <> "" Then
  fnam$ = "zzzsysisviz"
  c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + _
       form1.newid("auftritthigru", "id", 18) + "','" + _
       lastcreatedid$ + "','" + _
       form1.auftrittstyp(lastcreatedid$) + "','" + _
       fnam$ + "','" + Combo1.text + "')"
    Call form1.sqlqry(c$)
End If
End Sub

Private Sub neutodo_Click()
Dim neuid As String, d As Variant

'd2infile = "dayvw": d2insub = "neutodo_Click"
neuid = CDate(Text1.text)
Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), "" _
             , "", neuid, Left(Time, 5))
create2do.Text1(1).Enabled = False
Call create2do.SetFocus
create2do.Combo1.text = gd1.SelectedItem.text

End Sub

Private Sub opnterm_Click()
'd2infile = "dayvw": d2insub = "opnterm_Click"
Call gd1_DblClick
End Sub

Private Sub Text1_Change()
'd2infile = "dayvw": d2insub = "Text1_Change"
Call reinit
Call reshow

End Sub

Private Sub Text1_DblClick()

'd2infile = "dayvw": d2insub = "Text1_DblClick"
With frmCalendar
    .init Text1, Text1.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
End With
Unload frmCalendar

End Sub

Private Sub Text2_Change()
Dim stp As String

'd2infile = "dayvw": d2insub = "Text2_Change"
Call reinit
Call reshow
stp = trm(Text2.text)
Call form1.setusersetting("dayvwteilung", stp)

End Sub

Function isworkhour(minute As Integer) As Boolean

'd2infile = "dayvw": d2insub = "isworkhour"
isworkhour = False
If minute >= abeg And minute <= aend Then isworkhour = True

End Function

Function indexbyminute(minute As String) As Integer
Dim i%, trgt As Integer
'd2infile = "dayvw": d2insub = "indexbyminute"
indexbyminute = -1

trgt = time2minutes(minute)
For i% = 1 To gd1.ListItems.Count
  If time2minutes(gd1.ListItems(i%).text) >= trgt Then
   indexbyminute = i%
   Exit Function
  End If
Next i%
indexbyminute = gd1.ListItems.Count + 1
End Function

Sub reinit()
Dim i%, hh$, mm$, stp%, rrr, lvitem
Dim sindex As Integer

'd2infile = "dayvw": d2insub = "reinit"
MousePointer = 11: DoEvents
sindex = -1
On Error Resume Next
stp% = currstp()
gd1.ListItems.Clear
For i% = 0 To 1439 Step stp%
  hh$ = Int(i% / 60): If Len(hh$) = 1 Then hh$ = "0" + hh$
  mm$ = i% Mod 60: If Len(mm$) = 1 Then mm$ = "0" + mm$
  Set lvitem = gd1.ListItems.add(, , trm(hh$ + ":" + mm$))
  If isworkhour(i%) Then
    lvitem.Bold = True
  End If
Next i%
'sindex = (abeg + (aend - abeg) / 2) / stp%
sindex = indexbyminute(Left$(trm(Time), 5))
sindex = sindex + (((aend - abeg) / 2) / stp%)
If sindex >= 0 Then
  If sindex > gd1.ListItems.Count Then sindex = gd1.ListItems.Count
  On Error Resume Next
  Call gd1.ListItems(sindex).EnsureVisible
  On Error GoTo 0
End If
MousePointer = 0

End Sub

Sub reshow()
Dim rtmp As ADODB.Recordset, c$, i%, deflen As Integer, rrr, currcolor As Long
Dim lvitem, eidx As Integer, aidx As Integer, termterm As Integer
Dim currid As String, stp%, currbez As String, currtyp As String, currstat As String
Dim currtime As String, ctm As Integer, wd%, td%

Dim d2infile As String, d2insub As String
d2infile = "dayvw": d2insub = "reshow"
If trm(Text1.text) = "" Then Exit Sub
MousePointer = 11: DoEvents
wd% = Weekday(Text1.text)
wtag.Caption = transe(form1.tagesname(wd%))
jtag(3).FontBold = True
For i% = -3 To 3
  td% = wd% + i%
  If td% < 1 Then td% = td% + 7
  If td% > 7 Then td% = td% - 7
  If td% = 1 Then
    jtag(3 + i%).ForeColor = RGB(255, 0, 0)
  Else
    jtag(3 + i%).ForeColor = RGB(0, 0, 255)
  End If
  jtag(3 + i%).Caption = transe(form1.tagesname(td%))
Next i%
jtag(3).ForeColor = RGB(0, 0, 0)
c$ = "SELECT * FROM auftritt where datum='" + datum2sql(Text1.text) + "' order by zeit"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
While Not rtmp.EOF
  If form1.terminisviz4me(rtmp!id, Combo1.text) And Not form1.terminisinviz4me(rtmp!id, Combo1.text) Then
  c$ = trm(rtmp!zeit)
  If Len(c$) = 3 Then c$ = "0" + c$
  If Len(c$) = 4 Then c$ = Left$(c$, 2) + ":" + Right$(c$, 2)
  If Mid$(c$, 2, 1) = ":" Then c$ = "0" + c$
  aidx = time2minutes(c$): If aidx < 0 Then aidx = 0
  currtime = c$
  currcolor = 0
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    currcolor = Val(trm0(rtmp!optkalcolor))
  End If
  If currcolor = 0 Then currcolor = form1.get_eventcolor(rtmp!auftrittstyp)
  currbez = trm(rtmp!ort) + ": " + trm(rtmp!bezeichnung)
  currid = "(auftritt:" + rtmp!id
  currtyp = form1.get_atabkz(trm(rtmp!auftrittstyp))
  currstat = form1.get_eventstatusname(rtmp!astatus)
  termterm = time2minutes(form1.auftrittsende(rtmp!id, ""))
  If termterm < 0 Then
    On Error Resume Next
    deflen = Val(form1.getusersetting("termindauer", "60"))
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then deflen = 60
    termterm = time2minutes(c$) + deflen
  End If
  eidx = termterm - 1
  If eidx < aidx Then eidx = aidx


  stp% = currstp()
  For i% = aidx To eidx Step stp%
    c$ = minutes2time(i%)
    Set lvitem = gd1.ListItems.add(, , c$)
    lvitem.SubItems(LVCTYP) = currtyp
    lvitem.SubItems(LVCBEZ) = currbez
    lvitem.ForeColor = currcolor
    lvitem.SubItems(LVCID) = currid
    lvitem.SubItems(LVCSTATUS) = currstat
    If isworkhour(i%) Then lvitem.Bold = True
  Next i%
  aidx = indexbyminute(currtime): If aidx < 1 Then aidx = 1
  eidx = indexbyminute(minutes2time(termterm)) - 1
  If eidx < aidx Then eidx = aidx
  i% = eidx
  While i% >= aidx
    If gd1.ListItems(i%).SubItems(LVCTYP) = "" Then gd1.ListItems.Remove (i%)
    i% = i% - 1
  Wend



  End If
  rtmp.MoveNext
Wend
Call addtodos

currtime = Left(word2bis(trm(now)), 5)
ctm = indexbyminute(currtime) - 1
On Error Resume Next
gd1.ListItems(ctm).Selected = True
On Error GoTo 0
If ctm > 1 Then gd1.ListItems(1).Selected = False
MousePointer = 0

End Sub

Function indexbyid(ByVal id$) As Integer
Dim i%

'd2infile = "dayvw": d2insub = "indexbyid"
indexbyid = -1
For i% = 1 To gd1.ListItems.Count
  If InStr(gd1.ListItems(i%).SubItems(LVCID), ":" + id$) > 0 Then
    indexbyid = i%
    Exit Function
  End If
Next i%

End Function

Function minutes2time(mnte As Integer) As String
Dim h$, M$

'd2infile = "dayvw": d2insub = "minutes2time"
minutes2time = ""
h$ = trm(Int(mnte / 60)): If Len(h$) = 1 Then h$ = "0" + h$
M$ = trm(mnte Mod 60): If Len(M$) = 1 Then M$ = "0" + M$
minutes2time = h$ + ":" + M$

End Function

Sub addtodos()
Dim rrr
Dim cmd$, r As ADODB.Recordset, c$, lvitem
Dim wer As String

Dim d2infile As String, d2insub As String
d2infile = "dayvw": d2insub = "addtodos"
wer = trm(Combo1.text)
If wer = "" Then wer = form1.getuserid()
cmd$ = "select * from todolist where An='" + wer + "' and datum='" + datum2sql(Text1.text) + "' order by Datum,Zeit"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  c$ = trm(r!zeit)
  If Len(c$) = 3 Then c$ = "0" + c$
  If Len(c$) = 4 Then c$ = Left$(c$, 2) + ":" + Right$(c$, 2)
  If Mid$(c$, 2, 1) = ":" Then c$ = "0" + c$
  Set lvitem = gd1.ListItems.add(, , c$)
  If isworkhour(time2minutes(c$)) Then lvitem.Bold = True
  lvitem.ForeColor = RGB(196, 0, 0)
  lvitem.SubItems(LVCTYP) = "TODO"
  lvitem.SubItems(LVCBEZ) = trm(r!betreff)
  lvitem.SubItems(LVCID) = "(todo:" + r!id
  lvitem.SubItems(LVCSTATUS) = r!Status
  r.MoveNext
Wend

End Sub
Function currstp() As Integer
Dim rrr, stp%

'd2infile = "dayvw": d2insub = "currstp"
stp% = Val(Text2.text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then stp% = 30
If stp% <= 0 Then stp% = 30
currstp = stp%

End Function

Private Sub thisnoshow_Click()
Dim fnam$, c$, i As Integer, id As String, idtyp As String, didlist As String

'd2infile = "dayvw": d2insub = "thisnoshow_Click"
For i = 1 To gd1.ListItems.Count
  If gd1.ListItems(i).Selected Then
    id = gd1.ListItems(i).SubItems(LVCID)
    If id <> "" Then
      idtyp = cut_d1(id, ":")
      If idtyp = "(auftritt" Then
        id = cut_d2bis(id, ":")
        If InStr(didlist, "|" + id + "|") = 0 Then
          fnam$ = "zzzsysisinviz"
          c$ = "delete from auftritthigru where auftrittsid='" + id + "' and feldname='" + fnam$ + "' and felddaten='" + Combo1.text + "';"
          Call form1.sqlqry(c$)
          c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + _
             form1.newid("auftritthigru", "id", 18) + "','" + _
             id + "','" + _
             form1.auftrittstyp(id) + "','" + _
             fnam$ + "','" + Combo1.text + "')"
          Call form1.sqlqry(c$)
          didlist = didlist + "|" + id + "|"
        End If
      End If
    End If
  End If
Next i
Call Command4_Click

End Sub

Sub dayprint()
Dim o%, p%, rrr, fn$, l$, q%, t$, orgt$, t0$, marke$, rev$, ttest$
Dim vorlage$, inf As String, i%, uv$, erg$, ln$, pb%
Dim bkmstart$, bkmend$, udat As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "dayvw": d2insub = "dayprint"
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
vorlage = "tageskalender.rtf"
inf = form1.vorlagendir() + "\" + vorlage$
If exist(inf) = 0 Then
  MsgBox "Vorlage unbekannt: " + vorlage$
  Exit Sub
End If

o% = FreeFile
Open inf For Input As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Vorlage " & vorlage$ & " kann nicht geöffnet werden."
  Exit Sub
End If
fn$ = form1.myuniquedocname("")
If Len(trm(fn$)) = 0 Then Exit Sub
Set udat = New ADODB.Recordset
udat.CursorLocation = adUseServer
rrr = form1.adoopen(udat, "SELECT * FROM benutzerdaten where id ='" + form1.getuserid() + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
p% = FreeFile

MousePointer = 11: DoEvents
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  While Len(l$) > 0
    q% = InStr(l$, bkmstart$)
    If q% > 0 Then
      t$ = Mid$(l$, q% + Len(bkmstart$))
      t$ = Left$(t$, InStr(t$, "}") - 1)
      orgt$ = t$
      t0$ = t$
      t$ = LCase(t0$)
      If Left$(t0$, 5) = "MARKE" Then
        If isdigit(Mid$(t0$, 6, 1)) <> 0 Then
          If Mid$(t0$, 7, 1) = "_" Then
            marke$ = Left$(t0$, 7)
            t$ = Mid$(t$, 8)
          End If
        End If
      End If
      If Left$(t0$, 1) = "M" And Mid$(t0$, 3, 1) = "_" Then
        If Mid$(t0$, 3, 1) = "_" Then
          marke$ = Left$(t0$, 3)
          t$ = Mid$(t$, 4)
        End If
      End If
      If InStr(t$, "__") > 0 Then
        rev$ = Mid$(t$, InStr(t$, "__") + 2)
        ttest$ = Left$(t$, InStr(t$, "__") - 1)
      End If
      If ttest$ = "dvw" Then
        Select Case LCase(rev$)
          Case "datum": Print #p%, Left$(l$, q% - 1);: Print #p%, trm(Text1.text);
          Case "zeit": Print #p%, Left$(l$, q% - 1);: Print #p%, gd1.SelectedItem.text;
          Case Default: Print #p%, Left$(l$, q% - 1);: Print #p%, "(" + ttest$ + "__" + rev$ + ")";
        End Select
      End If
      If ttest$ = "user" Then
        If Not udat.EOF Then
          For i% = 0 To 21
            If Len(udat.Fields(i%).name) = Len(rev$) - 1 Then   'aliase ermitteln
              If isdigit(Right$(rev$, 1)) <> 0 Then rev$ = Left$(rev$, Len(rev$) - 1)
            End If
            If LCase(udat.Fields(i%).name) = LCase(rev$) Then
              uv$ = ""
              If Not IsNull(udat.Fields(i%).value) Then uv$ = udat.Fields(i%).value
              Print #p%, Left$(l$, q% - 1);:  Print #p%, strrepl(uv$, "\", "\\");
              i% = 33
            End If
          Next i%
          If i% < 30 Then
            erg$ = form1.getusersetting(rev$)
            If erg$ <> "" Then
              Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(erg$, "\", "\\");
            End If
          End If
        End If
      End If
      If ttest$ = "system" Then
        Select Case LCase(rev$)
          Case "datum": Print #p%, Left$(l$, q% - 1);: Print #p%, Date;
          Case "zeit": Print #p%, Left$(l$, q% - 1);: Print #p%, Left(Time, 5);
          Case Default:
        End Select
      End If

      ln$ = Mid$(l$, q% + 1)
      Do
        pb% = InStr(LCase(ln$), bkmend$ + LCase(orgt$))
        If pb% = 0 Then Line Input #o%, ln$
      Loop Until pb% > 0
        ln$ = Mid$(ln$, pb%)
      If InStr(ln$, "}") = 0 Then
        l$ = ""
      Else
        l$ = Mid$(ln$, InStr(ln$, "}") + 1)
      End If
    Else
      Print #p%, l$
      l$ = ""
    End If
    ttest$ = ""
    rev$ = ""

  Wend
Wend
Close #o%
Close #p%
MousePointer = 0
Call form1.openthisdoc(fn$, "")

End Sub
