VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form create2do 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Neue Aufgabe verfassen - AgencyProf"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox pin 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   45
      ToolTipText     =   "Sie benötigen Ihre PIN nur für die Änderung Ihrer persönlichen Daten."
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox wk_mail 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   44
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox wk_tel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   43
      Top             =   1380
      Width           =   2415
   End
   Begin VB.TextBox wk_name 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   42
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox wk_txt 
      Enabled         =   0   'False
      Height          =   975
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      Picture         =   "create2do.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   40
      ToolTipText     =   "Auftrag starten"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6840
      TabIndex        =   38
      ToolTipText     =   "Was ist das?"
      Top             =   240
      Width           =   375
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "create2do.frx":0544
      Left            =   4440
      List            =   "create2do.frx":0563
      TabIndex        =   36
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   975
      IntegralHeight  =   0   'False
      Left            =   4440
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00000000&
      Picture         =   "create2do.frx":058D
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "Telefon testen"
      Top             =   1360
      Width           =   375
   End
   Begin VB.CommandButton sndrec 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   240
      Picture         =   "create2do.frx":0717
      Style           =   1  'Grafisch
      TabIndex        =   32
      ToolTipText     =   "Nachricht aufnehmen"
      Top             =   3000
      Width           =   735
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
      TabIndex        =   31
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   1560
      Width           =   255
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   720
      Top             =   480
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   30
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   1
      ItemData        =   "create2do.frx":0B59
      Left            =   2880
      List            =   "create2do.frx":0B69
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      ItemData        =   "create2do.frx":0B89
      Left            =   1920
      List            =   "create2do.frx":0BAE
      TabIndex        =   28
      Top             =   4800
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   2760
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   960
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1320
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      ItemData        =   "create2do.frx":0BDA
      Left            =   2400
      List            =   "create2do.frx":0BEA
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      ItemData        =   "create2do.frx":0C0E
      Left            =   1440
      List            =   "create2do.frx":0C30
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "create2do.frx":0C53
      Left            =   2760
      List            =   "create2do.frx":0C9F
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "create2do.frx":0D4B
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Speichern"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "create2do.frx":128F
      Style           =   1  'Grafisch
      TabIndex        =   14
      ToolTipText     =   "Abbrechen / schlissen"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1005
      Index           =   4
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   9
      Text            =   "create2do.frx":13F3
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label pinlbl 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   46
      ToolTipText     =   "Aktuelles Datum mit Uhrzeit"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minuten vorher anrufen"
      Height          =   435
      Index           =   9
      Left            =   5280
      TabIndex        =   39
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefonische Erinnerung"
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
      Index           =   8
      Left            =   4440
      TabIndex        =   37
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "kommende Alarme"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   35
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5175
      Left            =   4320
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "mal, alle"
      Height          =   255
      Left            =   1320
      TabIndex        =   27
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "danach"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   4800
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4080
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Zeitraum"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Termin"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Team(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "in"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Uhr"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "um"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   12
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "am"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   10
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nachricht"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Betreff"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "An"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Von"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1935
      Left            =   840
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1335
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   3960
      Width           =   4095
   End
End
Attribute VB_Name = "create2do"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nflds As Integer, nores As Boolean
Dim d0 As Variant, delta As Long, deltaunit$
Dim perdeltau$
Dim dbn As String, dbu As String, dbh As String, dbp As String
Dim wk_n As String, wk_e As String, wk_t As String

Dim perdelta As Integer

Private Sub Combo2_Change(Index As Integer)
'd2infile = "create2do": d2insub = "Combo2_Change"
Call deltaset(Index)
End Sub

Private Sub Combo2_Click(Index As Integer)
'd2infile = "create2do": d2insub = "Combo2_Click"
Call deltaset(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)

'd2infile = "create2do": d2insub = "Combo2_LostFocus"
Call deltaset(Index)

End Sub

Private Sub Combo3_Change(Index As Integer)
'd2infile = "create2do": d2insub = "Combo3_Change"
Call deltaset(Index)
End Sub

Private Sub Combo3_Click(Index As Integer)
'd2infile = "create2do": d2insub = "Combo3_Click"
Call deltaset(Index)
End Sub

Private Sub Combo3_LostFocus(Index As Integer)

'd2infile = "create2do": d2insub = "Combo3_LostFocus"
Call deltaset(Index)

End Sub

Private Sub Command1_Click()
'd2infile = "create2do": d2insub = "Command1_Click"
Unload create2do
End Sub

Private Sub Command18_Click()
'd2infile = "create2do": d2insub = "Command18_Click"
Call form1.handbuchcall("05-Wiedervorlagen.htm")
End Sub

Private Sub Command2_Click()
Dim r As ADODB.Recordset, i%, j%, cmd$, rrr

Dim d2infile As String, d2insub As String
d2infile = "create2do": d2insub = "Command2_Click"
For i% = 0 To List1.ListCount - 1
  Call form1.dbg2f(List1.List(i%) & " " & List1.Selected(i%))
  If List1.Selected(i%) = True Then
    Call form1.new2do(Text1(1).text, List1.List(i%), Text1(3).text, Text1(4).text, datum2sql(Text1(5).text), Combo1.text, perdelta, perdeltau$, Val(Text2.text))
    List1.Selected(i%) = False
    DoEvents
  End If
Next i%

For j% = 0 To List3.ListCount - 1
  Call form1.dbg2f(List3.List(j%) & " " & List3.Selected(j%))
  If List3.Selected(j%) = True Then
    List3.Selected(j%) = False
    cmd$ = "select userid from benutzergruppen where groupid='" + List3.List(j%) + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
      For i% = 0 To List1.ListCount - 1
        If List1.List(i%) = r!userid Then
          List1.Selected(i%) = True
          i% = List1.ListCount
        End If
      Next i%
      r.MoveNext
    Wend
  End If

  For i% = 0 To List1.ListCount - 1
    Call form1.dbg2f(List1.List(i%) & " " & List1.Selected(i%))
    If List1.Selected(i%) = True Then
      Call form1.new2do(Text1(1).text, List1.List(i%), Text1(3).text, Text1(4).text, datum2sql(Text1(5).text), Combo1.text, perdelta, perdeltau$, Val(Text2.text))
      List1.Selected(i%) = False
      DoEvents
    End If
  Next i%
Next j%
Me.BackColor = form1.cleancolor()
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.dochistisopen Then
  If dochist2.topics.ListIndex > 0 Then Call dochist2.topics_Click
End If
Unload create2do
End Sub

Private Sub Command3_Click()
MsgBox ("no info yet")
End Sub

Private Sub Command4_Click()
Dim c$, nid$, trgdtg As Double, hrdtg
Dim Dt As Variant, d1tg$, d2tg$, dtg$, cmd$, d1t

If Not form1.alertdbok Then Exit Sub

hrdtg = CDate(Text1(5).text + " " + Combo1.text)
Dt = DateValue(hrdtg) - DateValue("1.1.1970 0:00:00")
Dt = Dt * 24 * 60 * 60
d1tg$ = trm(Time)
d1t = 3600 * Val(cut_d1(d1tg$, ":")): d1tg = cut_d2bis(d1tg$, ":")
d1t = d1t + 60 * Val(cut_d1(d1tg$, ":")): d1tg = cut_d2bis(d1tg$, ":")
d1t = d1t + Val(cut_d1(d1tg$, ":"))
Dt = Dt + d1t - form1.tzoffset
d1tg$ = Trim(str$(Dt))

nid = form1.alertdbuid + trm(d1tg$) + strrepl(trm(Rnd), ",", ".")
c$ = "insert into ruf (id,trgdtg,callnum,textmsg,uid) values('" + nid$ + "','" + trm(d1tg$) + "','" + trm(wk_tel) + "','" + trm(wk_txt) + "','" + trm(form1.alertdbuid) + "')"
Call form1.alrtdbsqlqry(c$)
Call rlist2

End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset, rrr, cmd$

'd2infile = "create2do": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
nores = False
nflds = 6
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
create2do.Caption = transe("Neue Aufgabe verfassen")
sndrec.ToolTipText = transe("Nachricht aufnehmen")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command2.ToolTipText = transe("Speichern")
Command1.ToolTipText = transe("Formular schliessen")
Label9.Caption = transe("mal, alle")
Label8.Caption = transe("danach")
Label7.Caption = transe("Zeitraum")
Label6.Caption = transe("Termin")
Label5.Caption = transe("Team(s)")
Label4.Caption = transe("Benutzer")
Label3.Caption = transe("in")
Label2.Caption = transe("Uhr")
Label1(6).Caption = transe("um")
Label1(5).Caption = transe("am")
Label1(4).Caption = transe("Nachricht")
Label1(3).Caption = transe("Betreff")
Label1(2).Caption = transe("An")
Label1(1).Caption = transe("Von")
wk_n = "": wk_e = "": wk_t = ""
If Not form1.alertdbok Then
  Command35.Enabled = False
Else
  cmd$ = "select tel,email,name from user where id='" + form1.alertdbuid + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.alertdbo, adOpenDynamic, adLockReadOnly)
  If rrr = 0 Then
    If Not r.EOF Then
      wk_n = trm(r!name): wk_e = trm(r!email): wk_t = trm(r!tel)
    Else
      wk_n = form1.getusersetting("name", "")
      wk_e = form1.getusersetting("email", "")
      wk_t = form1.getusersetting("tel", "")
    End If
  Else
    form1.alertdbok = False
  End If
End If
wk_name.text = wk_n
wk_tel.text = wk_t
wk_mail.text = wk_e
Show
delta = 0
Call rlist1
If form1.alertdbok Then Call rlist2
End Sub
Sub rlist1()
Dim rtmp As ADODB.Recordset
Dim r As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "create2do": d2insub = "rlist1"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

List1.Clear
While Not rtmp.EOF
  List1.AddItem rtmp!id
  rtmp.MoveNext
Wend

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM gruppennamen", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
List3.Clear
While Not r.EOF
  List3.AddItem r!gid
  r.MoveNext
Wend

End Sub

Private Sub Form_Resize()
'd2infile = "create2do": d2insub = "Form_Resize"
If Not nores Then axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "create2do": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub

Public Sub initmsg(von$, an$, betreff$, nachricht$, am$, um$)
Dim i%

'd2infile = "create2do": d2insub = "initmsg"
Text1(1).text = von$
Text1(2).text = an$
Text1(3).text = betreff$
Text1(4).text = nachricht$
Text1(5).text = am$
wk_txt.text = trm(betreff + " " + nachricht$)
d0 = Date
Combo1.text = um$

For i% = 0 To List1.ListCount - 1
  If LCase(List1.List(i%)) = LCase(an$) Then List1.Selected(i%) = True
Next i%
Me.BackColor = form1.dirtycolor()

End Sub
Sub deltaset(i%)
Dim df As Double, rrr

'd2infile = "create2do": d2insub = "deltaset"
On Error Resume Next
delta = Val(Combo2(i%).text)
rrr = Err
On Error GoTo 0
deltaunit$ = Combo3(i%).text
If trm(deltaunit$) = "" Or rrr <> 0 Then Exit Sub
Select Case LCase$(Left$(deltaunit$, 2))
  Case "ta": df = 1
  Case "wo": df = 7
  Case "mo": df = 30
  Case "ja": df = 365.25
  Case Default: df = 0
End Select
On Error Resume Next
If i% = 0 Then
  Text1(5).text = Left(d0 + CDate(delta * df), 10)
Else
  perdeltau$ = LCase$(Left$(deltaunit$, 2))
  perdelta = delta
End If
On Error GoTo 0
End Sub

Private Sub List1_Click()
Me.BackColor = form1.dirtycolor()
End Sub

Private Sub List3_Click()
Me.BackColor = form1.dirtycolor()
End Sub

Private Sub sndrec_Click()
Dim id$, tn$, X
Dim p_in%, sfn$

'd2infile = "create2do": d2insub = "sndrec_Click"
sfn$ = ""
p_in% = InStr(Text1(3).text, "[Wiedervorlage] ")
If p_in% > 0 Then
  tn$ = Mid$(Text1(3).text, p_in% + 16)
  id$ = Mid$(tn$, InStr(tn$, ":") + 1)
  tn$ = trm(LCase$(Left$(tn$, InStr(tn$, ":") - 1)))
  If tn$ = "voicemail" And exist(id$) <> 0 Then sfn$ = id$
End If

If sfn$ = "" Then
  sfn$ = form1.myuniquedocname("", ".wav")
  If sfn$ = "" Then Exit Sub
  Text1(3).text = trm(Text1(3).text) & " [Wiedervorlage] Voicemail:" & sfn$
  Call FileCopy(form1.s0dir() & "\wav\s0.wav", sfn$)
  Me.BackColor = form1.dirtycolor()
End If
If sfn$ <> "" Then
  X = Shell("sndrec32.exe " & sfn$, 1)
End If

End Sub

Private Sub Text1_Change(Index As Integer)
Me.BackColor = form1.dirtycolor()
End Sub

Private Sub Text1_DblClick(Index As Integer)
'd2infile = "create2do": d2insub = "Text1_DblClick"
If Index = 3 Then Exit Sub
  With frmCalendar
    .Init Text1(5), Text1(5).text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1(5).text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar
  Me.BackColor = form1.dirtycolor()

End Sub

Private Sub savecheck()
Dim antw As Integer

If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command2_Click
  End If
End If
BackColor = form1.cleancolor()

End Sub

Sub rlist2()
Dim rtmp As ADODB.Recordset
Dim rrr
Dim Dt As Variant, d1tg$, d2tg$, dtg$, cmd$, d1t

Dim d2infile As String, d2insub As String
d2infile = "create2do": d2insub = "rlist2"

Dt = DateValue(Date) - DateValue("1.1.1970 0:00:00")
Dt = Dt * 24 * 60 * 60
d1tg$ = trm(Time)
d1t = 3600 * Val(cut_d1(d1tg$, ":")): d1tg = cut_d2bis(d1tg$, ":")
d1t = d1t + 60 * Val(cut_d1(d1tg$, ":")): d1tg = cut_d2bis(d1tg$, ":")
d1t = d1t + Val(cut_d1(d1tg$, ":"))
Dt = Dt + d1t

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM ruf where uid='" + form1.alertdbuid + "' order by trgdtg", form1.alertdbo, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
List2.Clear
While Not rtmp.EOF
  List2.AddItem trm(Int((CDbl(rtmp!trgdtg) - Dt) / 86400)) + " Tg. " + trm(rtmp!textmsg)
  DoEvents
  rtmp.MoveNext
Wend

End Sub

