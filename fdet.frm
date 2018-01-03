VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form fdet 
   Caption         =   "Finanzdetails"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form2"
   ScaleHeight     =   2655
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox anmerk 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   26
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Netto aus Brutto"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      TabIndex        =   25
      Top             =   1545
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Picture         =   "fdet.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   24
      ToolTipText     =   "Eintrag ins Kassenbuch"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   3600
      Top             =   360
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      Picture         =   "fdet.frx":018A
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "Auftritt speichern"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ComboBox anid 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   4200
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox vonid 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   360
      Top             =   240
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "fdet.frx":0531
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Formular schliessen"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox nettobet 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4200
      TabIndex        =   17
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox mwstw 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   15
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox waehr 
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
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox anz 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox nettobet 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox anid 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox vonid 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox mwst 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label xpara 
      Caption         =   "xpara"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Anmerkung"
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Kontakt"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Brutto Endbetrag:"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   20
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
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
      Left            =   4320
      TabIndex        =   19
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Summe Netto"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   " MwSt"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Währung"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Anzahl / Dauer"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nettobetrag Einzelpr."
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "An"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Von"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "% MwSt"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label fid 
      Caption         =   "fid"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   3975
   End
End
Attribute VB_Name = "fdet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currkid$
Dim lchg1 As Double, srchit%
Dim lchg2 As Double
Dim achg1 As Double
Dim achg2 As Double

Private Sub anid_Change(Index As Integer)

'd2infile = "fdet": d2insub = "anid_Change"
If srchit% = 0 Then Exit Sub
If Index = 0 Then achg1 = Date + Time
If Index = 1 Then achg2 = Date + Time

End Sub

Private Sub anid_Click(Index As Integer)
Dim von$, srchit%, p%, c$, voni$

'd2infile = "fdet": d2insub = "anid_Click"
srchit% = 0
von$ = trm(anid(Index).text)
p% = InStr(von$, "(") - 1
If p% > 0 Then von$ = trm(Left$(von$, p%))
If Index = 0 Then
  c$ = "update finanzen set an='" & Left(txt2db(von$), 70) & "' where id='" & fid.Caption & "'"
Else
  c$ = "update finanzen set kan='" & Left(txt2db(von$), 70) & "' where id='" & fid.Caption & "'"
  Call form1.sqlqry(c$)
  von$ = trm(anid(Index).text)
  p% = InStr(von$, "ID:") + 3
  voni$ = trm(Mid$(von$, p%))
  c$ = "update finanzen set an='" & Left(txt2db(form1.getadridbykontaktid(voni$)), 70) & "' where id='" & fid.Caption & "'"
  anid(0).Clear: DoEvents
  anid(0).text = form1.getadridbykontaktid(voni$)
End If
Call form1.sqlqry(c$)
DoEvents
srchit% = 1
Call vonid(1).SetFocus


End Sub

Private Sub Command1_Click()
'd2infile = "fdet": d2insub = "Command1_Click"
Unload Me
'Unload auftritt
End Sub


Public Sub Command10_Click()
Dim r As ADODB.Recordset, p%, anm$, fldName$, id$, trgid$, c$
Dim fldnam$, typ$, rrr, net$, snet$, wae$, nanz$, prototyp$
Dim c0$, c0a$, c1$, c1a$, c2$, c3$, c4$, c5$, c6$, c7$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "fdet": d2insub = "Command10_Click"
fldName$ = ""
id$ = fid.Caption
p% = InStr(fid.Caption, "(ID:")
If p% > 0 Then
  id$ = Mid$(fid.Caption, p% + 4)
  fldName$ = Left$(fid.Caption, p% - 1)
End If
If auftritt.Text1(0).text <> id$ Then
  Call auftritt.SetFocus
  Call auftritt.showrec(fid.Caption, 0)
End If
trgid$ = fid.Caption
c$ = "update finanzen set mwst=" & trm(Int(100 * var2dbl(trm(mwst.text)))) & " where id='" & trgid$ & "'"
Call form1.sqlqry(c$)
c$ = "update finanzen set anz=" & trm0(d2db(anz.text)) & " where id='" & trgid$ & "'"
Call form1.sqlqry(c$)
fldnam$ = "honorar": If fldName$ <> "" Then fldnam$ = fldName$
typ$ = ""
c$ = "select auftrittstyp from auftritt where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If r.EOF Then Exit Sub                                          'cannot be error
  typ$ = r!auftrittstyp
  cmd$ = LCase(typ$ + fldnam$ + "setztmwst")
  If form1.getusersetting(cmd$, "") = "ja" Then
    c$ = "update finanzen set mwst=" & trm(Int(100 * var2dbl(trm(mwst.text)))) & " where id='" & id$ & "'"
    Call form1.sqlqry(c$)
  End If

net$ = fixeur(trm0(nettobet(0).text))
c$ = "update finanzen set netto=" & strrepl(var2dbl(net$), ",", ".") & " where id='" & trgid$ & "'"
Call form1.sqlqry(c$)
snet$ = fixeur(trm0(nettobet(1).text))
wae$ = " " & waehr.text
c$ = "update finanzen set waehrung='" & wae$ & "' where id='" & trgid$ & "'"
Call form1.sqlqry(c$)
nanz$ = anz.text
c0$ = "": c0a$ = "": c1$ = "": c1a$ = "": c2$ = "": c3$ = "": c4$ = "": c5$ = "": c6$ = "": c7$ = ""
anm$ = trm(anmerk.text): If anm$ <> "" Then anm$ = " " + anm$
prototyp$ = typ$
If LCase(prototyp$) = "promo" Then prototyp$ = "Künstlerauftritt"
If LCase(prototyp$) = "perfartist" Then prototyp$ = "Künstlerauftritt"
If LCase(prototyp$) = "deal" Then prototyp$ = "Künstlerauftritt"
Select Case LCase(prototyp$)
  Case "orchesterauftritt"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & anm$ & "')"
    'c1$ = "update auftritthigru set felddaten='" & net$ & wae$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c2$ = "update usr_" & utabn(typ$) & " set " & fldnam$ & "='" & net$ & wae$ & anm$ & "' where id='" & id$ & "'"
    c3$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "komposition"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & anm$ & "')"
    c2$ = "update usr_" & utabn(typ$) & " set " & fldnam$ & "='" & net$ & wae$ & anm$ & "' where id='" & id$ & "'"
    c3$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "künstlerauftritt"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & anm$ & "')"
    c2$ = "update usr_" & utabn(typ$) & " set " & fldnam$ & "='" & net$ & wae$ & anm$ & "' where id='" & id$ & "'"
    c3$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "dirigentenauftritt"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & anm$ & "')"
    c2$ = "update usr_" & utabn(typ$) & " set " & fldnam$ & "='" & net$ & wae$ & anm$ & "' where id='" & id$ & "'"
    c3$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "chorauftritt"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & anm$ & "')"
    c2$ = "update usr_" & utabn(typ$) & " set " & fldnam$ & "='" & net$ & wae$ & anm$ & "' where id='" & id$ & "'"
    c3$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "dienstleistung"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & "')"
    c1a$ = "update auftritthigru set felddaten='" & net$ & wae$ & anm$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='Betrag_pro_Stunde'"
    c2$ = "update usr_" & utabn(typ$) & " set Betrag_pro_Stunde='" & net$ & wae$ & "' where id='" & id$ & "'"
    c3$ = "update auftritthigru set felddaten='" & nanz$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='Dauer'"
    c4$ = "update usr_" & utabn(typ$) & " set dauer='" & nanz$ & "' where id='" & id$ & "'"
    c5$ = "update auftritthigru set felddaten='" & snet$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='honorar'"
    c6$ = "update usr_" & utabn(typ$) & " set honorar='" & snet$ & "' where id='" & id$ & "'"
    c7$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case "verkauf"
    c0$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldnam$ & "'"
    c1$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & form1.newid("auftritthigru", "id", 40) + _
        "','" & id$ + _
        "','" & typ$ + _
        "','" & fldnam$ + _
        "','" & net$ & wae$ & "')"
'    c1$ = "update auftritthigru set felddaten='" & net$ & wae$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='Einzelpreis'"
    c2$ = "update usr_" & utabn(typ$) & " set Einzelpreis='" & net$ & wae$ & "' where id='" & id$ & "'"
    c3$ = "update auftritthigru set felddaten='" & nanz$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='Anzahl'"
    c4$ = "update usr_" & utabn(typ$) & " set anzahl='" & nanz$ & "' where id='" & id$ & "'"
    c5$ = "update auftritthigru set felddaten='" & snet$ & "' where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='Gesamtpreis'"
    c6$ = "update usr_" & utabn(typ$) & " set Gesamtpreis='" & snet$ & "' where id='" & id$ & "'"
    c7$ = "delete FROM auftritthigru where auftrittsid='" + id$ + "' and feldname='" + fldnam$ + "' and auftrittstyp='kalku_" + typ$ + "'"
  Case Else
End Select
If c0$ <> "" Then Call form1.sqlqry(c0$)
If c1$ <> "" Then Call form1.sqlqry(c1$)
If c1a$ <> "" Then Call form1.sqlqry(c1a$)
If c2$ <> "" Then Call form1.sqlqry(c2$)
If c3$ <> "" Then Call form1.sqlqry(c3$)
If c4$ <> "" Then Call form1.sqlqry(c4$)
If c5$ <> "" Then Call form1.sqlqry(c5$)
If c6$ <> "" Then Call form1.sqlqry(c6$)
If c7$ <> "" Then Call form1.sqlqry(c6$)
'150728: removed (possibly client standort nicht verf...
'Unload auftritt
'Call auftritt.SetFocus
Me.BackColor = form1.cleancolor()
Call auftritt.showrec(id$, 0)
DoEvents
Call auftritt.Command10_Click
End If              'rrr=0

Unload fdet
End Sub

Private Sub Command2_Click()
Dim c$, r As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "fdet": d2insub = "Command2_Click"
c$ = "select * from auftritt where id='" & fid.Caption & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
c$ = "insert into kassenbuch (id,thema,dtg,vorgang,zahlstatus,vonid," + _
                             "kontaktname,anzahl,Bezeichnung,epreisnetto,mwst) values('" + _
      form1.newid("kassenbuch", "id", 30) & "','Auftritt','" + _
      datum2sql(Date) & " " & Time & "','" + _
      r!id & "','" + _
      "berechnet" & "','" + _
      vonid(0).text & "','" + _
      "" & "'," & d2db(anz.text) & ",'" + _
      trm(r!ort) & " " & r!datum & " " & vonid(0).text & " " & anid(0).text & "'," + _
      d2db(nettobet(1).text) & "," + _
      d2db(mwstw.text) & ")"
'Call form1.sqlqry(c$)
End If
End Sub

Private Sub Command3_Click()
Dim dbrt As Double, dnet As Double, mwstm As Double

'd2infile = "fdet": d2insub = "Command3_Click"
dbrt = var2dbl(word1(strrepl(trm(nettobet(0).text), ".", "")))
mwstm = var2dbl(word1(strrepl(trm(mwst.text), ".", "")))
dnet = dbrt / (mwstm + 100) * 100
nettobet(0).text = fixeur("0" & dnet)
End Sub

Private Sub fid_Change()
Dim r As ADODB.Recordset, auf As ADODB.Recordset, c$, rrr, ran$
Dim danz As Double, dnet As Double, mwstm As Double, typ$
Dim id$, fldName$, rnet$, p%, fldan$
Dim d2infile As String, d2insub As String

d2infile = "fdet": d2insub = "fid_Change"
srchit% = 0
id$ = fid.Caption
p% = InStr(fid.Caption, "(ID:")
If p% > 0 Then
  id$ = Mid$(fid.Caption, p% + 4)
  fldName$ = Left$(fid.Caption, p% - 1)
  fldan$ = Mid$(fldName$, 8)
End If
typ$ = ""
c$ = "select * from auftritt where id='" & id$ & "'"
Set auf = New ADODB.Recordset
auf.CursorLocation = adUseServer
rrr = form1.adoopen(auf, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not auf.EOF Then typ$ = auf!auftrittstyp

c$ = "select * from finanzen where id='" & fid.Caption & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fdet.Caption = transe("Finanzdetails") & ": " & r!bezeichnung
  vonid(0).text = trm(r!von)
  anid(0).text = trm(r!an)
  If trm(r!kvon) <> "" Then vonid(1).text = r!kvon
  If trm(r!kan) <> "" Then anid(1).text = r!an
  anz.text = trm(r!anz)
  rnet$ = trm(r!netto): If rnet$ = "" Then rnet$ = "0"
  nettobet(0).text = trm(fixeur("" & rnet$))
  waehr.text = trm(r!waehrung)
  mwstm = 0: If Not IsNull(r!mwst) Then mwstm = r!mwst / 100
  mwst.text = trm(fixeur(mwstm))
  If LCase(typ$) = "hotelaufenthalt" Then
    anz.Enabled = False
    nettobet(0).Enabled = False
    waehr.Enabled = False
  End If
Else
  fdet.Caption = "Finanzdetails: " & auf!bezeichnung & " " & fldName$
  vonid(0).text = "Veranstalter"
  ran$ = "select * from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & typ$ & "' and feldname='" & fldan$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, ran$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    ran$ = r!felddaten
  Else
    ran$ = ""
  End If
  anid(0).text = trm(ran$)
  vonid(1).text = "Veranstalter"
  anid(1).text = ran$
  anz.text = "1"
  nettobet(0).text = trm(fixeur("0"))
  waehr.text = transe("€")
  mwstm = fixeur(Val("0" & form1.getusersetting("auftrittsmwst", form1.getusersetting("mwst", 1900))) / 100)
  mwst.text = trm(fixeur(mwstm))
  If LCase(typ$) = "hotelaufenthalt" Then
    anz.Enabled = False
    nettobet(0).Enabled = False
    waehr.Enabled = False
  End If
  c$ = "insert into finanzen (id) values('" & fid.Caption & "')": Call form1.sqlqry(c$)
End If
BackColor = form1.cleancolor()
srchit% = 1
End Sub

Private Sub Form_Load()
Dim s%, lchg As Integer, i%

'd2infile = "fdet": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
s% = form1.myfontsize()
anid(0).Font.Size = s%
anid(1).Font.Size = s%
vonid(0).Font.Size = s%
vonid(1).Font.Size = s%
nettobet(0).Font.Size = s%
nettobet(1).Font.Size = s%
waehr.Font.Size = s%
mwstw.Font.Size = s%
mwst.Font.Size = s%
anz.Font.Size = s%
anz.Font.Size = s%
anz.Font.Size = s%
anz.Font.Size = s%
anz.Font.Size = s%
anmerk.Font.Size = s%
lchg = 0
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
fdet.Caption = transe("Finanzdetails")
Command3.Caption = transe("Netto aus Brutto")
Command2.ToolTipText = transe("Eintrag ins Kassenbuch")
Command10.ToolTipText = transe("Auftritt speichern")
Command1.ToolTipText = transe("Formular schliessen")
Label10.Caption = transe("Kontakt")
Label8(1).Caption = transe("Brutto Endbetrag:")
Label8(0).Caption = transe("Summe Netto")
Label7.Caption = transe("MwSt")
Label6.Caption = transe("Währung")
Label5.Caption = transe("Anzahl / Dauer")
Label4.Caption = transe("Nettobetrag Einzelpr.")
Label3.Caption = transe("An")
Label2.Caption = transe("Von")
Label1.Caption = transe("% MwSt")
Label11.Caption = transe("Anmerkung")

fdet.Caption = transe("Finanzdetails")
Command3.Caption = transe("Netto aus Brutto")
Command2.ToolTipText = transe("Eintrag ins Kassenbuch")
Command10.ToolTipText = transe("Auftritt speichern")
Command1.ToolTipText = transe("Formular schliessen")
Label10.Caption = transe("Kontakt")
Label8(1).Caption = transe("Brutto Endbetrag:")
Label9.Caption = transe("--,--")
Label8(0).Caption = transe("Summe Netto")
Label7.Caption = " " + transe("MwSt")
Label6.Caption = transe("Währung")
Label5.Caption = transe("Anzahl / Dauer")
Label4.Caption = transe("Nettobetrag Einzelpr.")
Label3.Caption = transe("An")
Label2.Caption = transe("Von")
Label1.Caption = transe("% MwSt")
Label11.Caption = transe("Anmerkung")
Show
waehr.Clear
For i% = 0 To form1.waehrungen.ListCount - 1
  waehr.AddItem cut_d1(form1.waehrungen.List(i%), ":")
Next i%

End Sub

Private Sub Form_Resize()
'd2infile = "fdet": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

'd2infile = "fdet": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld

Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Private Sub anz_Change()
Dim mwstm As Double, danz As Double, dnet As Double

'd2infile = "fdet": d2insub = "anz_Change"
  If waehr.text = "" Then waehr.text = "€"
  If trm(nettobet(0).text) = "" Then Exit Sub
  If trm(anz.text) = "" Then Exit Sub
  If trm(mwst.text) = "" Then Exit Sub
  BackColor = form1.dirtycolor()
  On Error Resume Next
  danz = var2dbl(word1(strrepl(trm(anz.text), ".", "")))
  dnet = var2dbl(word1(strrepl(trm(nettobet(0).text), ".", "")))
  mwstm = var2dbl(word1(strrepl(trm(mwst.text), ".", "")))
  nettobet(1).text = fixeur(danz * dnet)
  mwstm = danz * dnet * mwstm / 100
  mwstw.text = fixeur(mwstm)
  Label9.Caption = fixeur(mwstm + danz * dnet) & " " & waehr.text
  On Error GoTo 0
End Sub

Private Sub mwst_Change()
'd2infile = "fdet": d2insub = "mwst_Change"
Call anz_Change
End Sub

Private Sub nettobet_Change(Index As Integer)
'd2infile = "fdet": d2insub = "nettobet_Change"
Call anz_Change
End Sub
Sub savecheck()
Dim antw As Integer

'd2infile = "fdet": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command10_Click
  End If
End If
BackColor = form1.cleancolor()

End Sub

Private Sub Timer1_Timer()
Dim rtmp As ADODB.Recordset, Index As Integer, s$
Dim nw As Double, cmd$, rrr, fcnt%

Dim d2infile As String, d2insub As String
d2infile = "fdet": d2insub = "Timer1_Timer"
Call form1.dbg2f("fdet Timer1 start")
nw = Date + Time
If lchg1 > 0 Or achg1 > 0 Then
If (nw - lchg1) * 86400000 > 900 Or (nw - achg1) * 86400000 > 900 Then
  Index = 0
  If lchg1 > 0 Then
    s$ = trm(vonid(Index).text)
    vonid(Index).Clear
  Else
    s$ = trm(anid(Index).text)
    anid(Index).Clear
  End If
  If s$ = "" Then
    Call form1.dbg2f("fdet Timer1 exit")
    Exit Sub
  End If
  cmd$ = "SELECT * FROM adresse where ( (" + _
       "instr(lcase(strasse),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(ort),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(plz),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(id),'" + LCase(s$) + "')>0) or (" + _
       "instr(telfaxhandy,'" + s$ + "')>0) or (" + _
       "instr(lcase(url),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(email),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(name),'" + LCase(s$) + "')>0) )"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    rtmp.MoveFirst
    fcnt% = 0
    While Not rtmp.EOF And fcnt% < 50
      If trm(rtmp!id) <> trm(rtmp!name) Then
        fcnt% = fcnt% + 1
        If lchg1 > 0 Then
          vonid(Index).AddItem rtmp!id & "(" + rtmp!name + ")"
        Else
          anid(Index).AddItem rtmp!id & "(" + rtmp!name + ")"
        End If
      End If
      rtmp.MoveNext
    Wend
  End If
End If
If lchg1 > 0 Then
  lchg1 = 0
Else
  achg1 = 0
End If
End If
If lchg2 > 0 Or achg2 > 0 Then
Index = 1
If (nw - lchg2) * 86400000 > 900 Or (nw - achg2) * 86400000 > 900 Then
  If lchg2 > 0 Then
    s$ = trm(vonid(Index).text)
    vonid(Index).Clear
  Else
    s$ = trm(anid(Index).text)
    anid(Index).Clear
  End If
  If s$ = "" Then
    Call form1.dbg2f("fdet Timer1 exit")
    Exit Sub
  End If
  cmd$ = "SELECT id,name,position FROM kontakt where (" + _
    " (instr(lcase(name),'" + LCase(s$) + "')>0) or " + _
    " (instr(lcase(email),'" + LCase(s$) + "')>0) or " + _
    " (instr(telfaxhandy,'" + s$ + "')>0) )"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    rtmp.MoveFirst
    While Not rtmp.EOF
      If lchg2 > 0 Then
        vonid(Index).AddItem rtmp!name & "(" & rtmp!Position & ")" & Space$(80) & "(ID:" & rtmp!id
      Else
        anid(Index).AddItem rtmp!name & "(" & rtmp!Position & ")" & Space$(80) & "(ID:" & rtmp!id
      End If
      rtmp.MoveNext
    Wend
  End If
End If
If lchg2 > 0 Then
  lchg2 = 0
Else
  achg2 = 0
End If
End If
Call form1.dbg2f("fdet Timer1 exit")
End Sub

Private Sub vonid_Change(Index As Integer)

'd2infile = "fdet": d2insub = "vonid_Change"
If srchit% = 0 Then Exit Sub
If Index = 0 Then lchg1 = Date + Time
If Index = 1 Then lchg2 = Date + Time

End Sub

Private Sub vonid_Click(Index As Integer)
Dim von$, p%, c$, voni$

'd2infile = "fdet": d2insub = "vonid_Click"
srchit% = 0
von$ = trm(vonid(Index).text)
p% = InStr(von$, "(") - 1
If p% > 0 Then von$ = trm(Left$(von$, p%))
If Index = 0 Then
  c$ = "update finanzen set von='" & txt2db(von$) & "' where id='" & fid.Caption & "'"
Else
  c$ = "update finanzen set kvon='" & txt2db(von$) & "' where id='" & fid.Caption & "'"
  Call form1.sqlqry(c$)
  von$ = trm(vonid(Index).text)
  p% = InStr(von$, "ID:") + 3
  voni$ = trm(Mid$(von$, p%))
  c$ = "update finanzen set von='" & txt2db(form1.getadridbykontaktid(voni$)) & "' where id='" & fid.Caption & "'"
  vonid(0).Clear: DoEvents
  vonid(0).text = form1.getadridbykontaktid(voni$)
End If
Call form1.sqlqry(c$)
DoEvents
srchit% = 1
Call anid(1).SetFocus

End Sub

Private Sub waehr_Change()
'd2infile = "fdet": d2insub = "waehr_Change"
BackColor = form1.dirtycolor()
End Sub

Private Sub waehr_Click()
'd2infile = "fdet": d2insub = "waehr_Click"
BackColor = form1.dirtycolor()
End Sub

Private Sub xpara_Change()
Dim p%, a$

a$ = trm(xpara.Caption)
p% = InStr(a$, waehr.text)
If p% > 0 Then
  p% = p% + Len(waehr.text)
  If p% >= Len(a$) Then
    a$ = ""
  Else
    a$ = trm(Mid(a$, p%))
  End If
End If
anmerk.text = a$

End Sub

