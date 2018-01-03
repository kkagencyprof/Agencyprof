VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form verwaltung 
   Caption         =   "Verwaltung"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command35 
      Caption         =   "txt2eml"
      Height          =   375
      Left            =   8400
      TabIndex        =   41
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "msg2eml"
      Height          =   375
      Left            =   7560
      TabIndex        =   40
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command34 
      Caption         =   "drop || from checklists"
      Height          =   495
      Left            =   2880
      TabIndex        =   39
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command33 
      Caption         =   "anymdb2sql"
      Height          =   375
      Left            =   4560
      TabIndex        =   38
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Mussfelder setzen"
      Height          =   375
      Left            =   2160
      TabIndex        =   37
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6120
      TabIndex        =   36
      Text            =   "Text3"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command32 
      Caption         =   "<-"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      Caption         =   "->"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4080
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Kontakte o. Adr."
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   31
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command23 
      Caption         =   "temp"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "teste adreesstyp"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command19 
      Caption         =   "CRLF entfernen"
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Command27 
      Caption         =   "spamdoms lesen"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Postanr., Trim(Name)"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command25 
      Caption         =   "doppelte Kontakte löschen"
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Kontakt -> Person"
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Export Handy"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command29 
      Caption         =   "msfe.erst. konv."
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command21 
      Caption         =   "KSTL...->Person, TelFaxHandy"
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Finanzen"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "create dbstamp"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "create apdemo.rtf.publish"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   120
      Width           =   2295
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   840
      Top             =   120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command18 
      Caption         =   "checklists"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "leere Kontakte u. Hinweise löschen"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "tfh löschen"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "TelFaxHandy"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "pgdump-Werke"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "A&larmlisten"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Hintergrund löschen"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "STOP"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   840
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7440
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "0"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "aufbauen"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&GO"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Künstler,... -> Person"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tests"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "0"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "verwaltung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stopfl

Private Sub Command1_Click()

Hide
Unload verwaltung
End Sub

Private Sub Command10_Click()
Load alarmlist

End Sub

Private Sub Command11_Click()


If exist("vdkw.dump") Then Kill "vdkw.dump"

Call form1.pg_xp("k_loc", "vdkw.dump")
Call form1.pg_xp("w_loc", "vdkw.dump")
Call form1.pg_xp("sbz_loc", "vdkw.dump")

End Sub

Private Sub Command12_Click()
Dim rtmp As Recordset

MousePointer = 11: DoEvents
List1.Clear

sel$ = "SELECT * FROM kontakt"
Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)

While Not rtmp.EOF And List1.ListCount < 16000
  List1.AddItem "delete from adresstyp where vid='" + rtmp!vid + "' and typ='Person' and kid='" & rtmp!id & "'"
  List1.AddItem "insert into adresstyp (id,vid,kid,typ) values('" + form1.newid("adresstyp", "id", 20) + "','" + rtmp!vid + "','" & rtmp!id & "','Person')"
  rtmp.MoveNext
Wend
MousePointer = 0
End Sub

Private Sub Command13_Click()
Dim r As Recordset, s As Recordset, t As Recordset, a As Recordset
Dim net1l As Double, netl As Double, fi As ADODB.Recordset

stopfl = 0
Set t = form1.sqla.OpenRecordset( _
    "SELECT * FROM auftritt where datum>='" + trm(Text1.text) + "' order by datum", dbOpenDynaset, dbReadOnly)

While Not t.EOF And List1.ListCount < 12000 And stopfl = 0
  typ$ = t!auftrittstyp
  anz$ = "1"
  If typ$ <> "Neuer Auftritt" Then
    id$ = t!id
    bez$ = t!bezeichnung
    If id$ <> "" Then
      Text1.text = t!datum
      DoEvents
      Set r = form1.sqla.OpenRecordset( _
        "SELECT * FROM finanzen where id='" + id$ + "'", dbOpenDynaset, dbReadOnly)
      If r.EOF Then
      Set r = form1.sqla.OpenRecordset( _
        "SELECT * FROM usr_" & utabn(typ$) + " where id='" + id$ + "'", dbOpenDynaset, dbReadOnly)
      If Not r.EOF Then

        Select Case LCase(typ$)
          Case "deal"
            an$ = trm(r!Lieferant)
            von$ = trm(r!Kunde)
            net$ = form1.ohnewaehrung(trm(r!Honorar))
            wae$ = form1.nurdiewaehrung(trm(r!Honorar))
            tut$ = "Honorar " & trm(bez$)
          Case "orchesterauftritt"
            an$ = trm(r!orchester)
            von$ = trm(r!veranstalter)
            net$ = form1.ohnewaehrung(trm(r!Honorar))
            wae$ = form1.nurdiewaehrung(trm(r!Honorar))
            tut$ = "Honorar " & trm(bez$)
          Case "künstlerauftritt"
            an$ = trm(r!künstler)
            von$ = trm(r!veranstalter)
            net$ = form1.ohnewaehrung(trm(r!Honorar))
            wae$ = form1.nurdiewaehrung(trm(r!Honorar))
            tut$ = "Honorar " & trm(bez$)
          Case "dienstleistung"
            an$ = trm(r!wer)
            von$ = trm(r!Kunde)
            net$ = form1.ohnewaehrung(trm(r!betrag_pro_stunde))
            wae$ = form1.nurdiewaehrung(trm(r!betrag_pro_stunde))
            anz$ = word1(trm(r!Dauer))
            tut$ = trm(bez$)
          Case Else
            an$ = ""
            von$ = ""
            net$ = ""
        End Select
        If Len(von$ + an$ + net$) > 0 Then
        c$ = "insert into finanzen (id,mwst,anz) values('" & trm(id$) & "'," & form1.getusersetting("auftrittsmwst", form1.getusersetting("mwst", 1900)) & ",1)"
        Call form1.sqlqry(c$)
        ad$ = " where id='" & id$ & "'": ad1$ = "'" + ad$
        w$ = trm(an$)
        If w$ <> "" Then
          c$ = "update finanzen set an='" & Left(txt2db(w$), 70) & ad1$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(von$)
        If w$ <> "" Then
          c$ = "update finanzen set von='" & Left(txt2db(w$), 70) & ad1$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(net$): w$ = strrepl(strrepl(w$, ".", ""), ",", ".")
        If w$ <> "" Then
          c$ = "update finanzen set netto=" & d2db(word1(w$)) & ad$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(wae$)
        If w$ <> "" Then
          c$ = "update finanzen set waehrung='" & trm(w$) & ad1$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(anz$): w$ = strrepl(strrepl(w$, ".", ""), ",", ".")
        If w$ <> "" Then
          c$ = "update finanzen set anz=" & d2db(w$) & ad$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(tut$)
        If w$ <> "" Then
          c$ = "update finanzen set bezeichnung='" & trm(w$) & ad1$
          Call form1.sqlqry(c$)
        End If
        w$ = trm(typ$)
        If w$ <> "" Then
          c$ = "update finanzen set typ='" & trm(w$) & ad1$
          Call form1.sqlqry(c$)
        End If
        End If
      End If
      Else
        Select Case LCase(typ$)
          Case "dienstleistung"
            net$ = "0": anz = "0"
            an$ = "select * from auftritthigru where auftrittsid='" & id$ & "' and feldname='betrag_pro_stunde'"
            Set a = form1.sqla.OpenRecordset(an$, dbOpenDynaset, dbReadOnly)
            If Not a.EOF Then net$ = form1.ohnewaehrung(trm(a!felddaten))
            If trm(net$) = "" Then net$ = "0"
            On Error Resume Next
            netl = var2dbl(net$)
            rrr = Err
            On Error GoTo 0
            If rrr <> 0 Then netl = 0
            an$ = "select * from auftritthigru where auftrittsid='" & id$ & "' and feldname='dauer'"
            Set a = form1.sqla.OpenRecordset(an$, dbOpenDynaset, dbReadOnly)
            If Not a.EOF Then anz$ = form1.ohnewaehrung(trm(a!felddaten))
            net1$ = form1.ohnewaehrung(trm(r!netto))
            net1l = var2dbl(net1$)
            wae1$ = form1.nurdiewaehrung(trm(r!netto))
            anz1$ = word1(trm(r!anz))
            If Val(anz$) <> Val(anz1$) Then
              If Val("0" & anz$) = 0 Then
                c$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and feldname='dauer'"
                Call form1.sqlqry(c$)
                c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
                     form1.newid("auftritthigru", "id", 30) & "','" & id$ & "','" & _
                     "Dienstleistung','Dauer','" & anz1$ & "')"
                Call form1.sqlqry(c$)
'              Else
'                Call form1.new2do("www", "www", "Finanzen [Wiedervorlage] Auftritt:" & id$, "prüfen", datum2sql(Date), "0:00", 0, 0, 0)
              End If
            End If
            If netl <> net1l Then
              If net1l = "0" Then
                net2$ = strrepl(strrepl(net2$, ".", ""), ",", ".")
                c$ = "update finanzen set netto=" & d2db(net2$) & " where id='" & id$ & "'"
                Call form1.sqlqry(c$)
              Else
                If netl = 0 Then
                  c$ = "delete from auftritthigru where auftrittsid='" & id$ & "' and feldname='Betrag_pro_Stunde'"
                  Call form1.sqlqry(c$)
                  c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
                     form1.newid("auftritthigru", "id", 30) & "','" & id$ & "','" & _
                     "Dienstleistung','Betrag_pro_Stunde','" & net1$ & "')"
                  Call form1.sqlqry(c$)
'                Else
'                  Call form1.new2do("www", "www", "Finanzen [Wiedervorlage] Auftritt:" & id$, "prüfen", datum2sql(Date), "0:00", 0, 0, 0)
                End If
              End If
            End If
          Case Else
            'Call form1.new2do("www", "www", "Finanzen [Wiedervorlage] Auftritt:" & id$, "prüfen", datum2sql(Date), "0:00", 0, 0, 0)
        End Select
      End If
    End If
  End If
Set fi = New ADODB.Recordset
fi.CursorLocation = adUseServer
c$ = "SELECT * FROM finanzen where id='" + trm(t!id) + "'"
rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
If Not fi.EOF Then
Set r = form1.sqla.OpenRecordset( _
  "SELECT * FROM usr_" & utabn(typ$) + " where id='" + trm(t!id) + "'", dbOpenDynaset, dbReadOnly)

Select Case LCase(typ$)
  Case "künstlerauftritt"
    If trm(fi!an) = "" Then
      On Error Resume Next
      an$ = trm(r!künstler)
      On Error GoTo 0
    End If
    If trm(fi!von) = "" Then von$ = trm(r!veranstalter)
    net$ = form1.ohnewaehrung(trm(fi!netto * fi!anz))
    wae$ = form1.nurdiewaehrung(trm0(fi!netto))
    tut$ = "Honorar " & trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & ",1)": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    net2$ = strrepl(strrepl(net2$, ".", ""), ",", ".")
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case "dienstleistung"
    an$ = trm(r!wer)
    von$ = trm(r!Kunde)
    net$ = form1.ohnewaehrung(trm(r!Honorar))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    anz$ = word1(trm(r!Dauer))
    tut$ = trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & ",1)": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    net$ = strrepl(strrepl(net$, ".", ""), ",", ".")
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case Else
End Select
End If
  t.MoveNext
Wend
End Sub

Private Sub Command14_Click()
Dim rtmp As Recordset
MousePointer = 11
DoEvents
List1.Clear

sel$ = "SELECT id,tel,fax,handy "
sel$ = sel$ + "FROM adresse "
sel$ = sel$ + "WHERE telfaxhandy='nix'"

Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)

While Not rtmp.EOF
  tfh$ = trm(onlynums("" & rtmp!tel) & " " & onlynums("" & rtmp!fax) & " " & onlynums("" & rtmp!handy))
  If tfh$ <> "" Then List1.AddItem "update adresse set telfaxhandy='" & tfh$ & "' where id='" & rtmp!id & "'"
  rtmp.MoveNext
Wend
Call form1.sqlqry("update adresse set telfaxhandy=' ' where telfaxhandy='nix';")

sel$ = "SELECT id,tel,fax,handy "
sel$ = sel$ + "FROM kontakt "
sel$ = sel$ + "WHERE telfaxhandy='nix'"

Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)
While Not rtmp.EOF
  tfh$ = trm(onlynums("" & rtmp!tel) & " " & onlynums("" & rtmp!fax) & " " & onlynums("" & rtmp!handy))
  If tfh$ <> "" Then List1.AddItem "update kontakt set telfaxhandy='" & tfh$ & "' where id='" & rtmp!id & "'"
  rtmp.MoveNext
Wend
Call form1.sqlqry("update kontakt set telfaxhandy=' ' where telfaxhandy='nix';")
MousePointer = 0
End Sub

Private Sub Command15_Click()
MousePointer = 11
DoEvents
form1.sqlqry ("update adresse set telfaxhandy='nix'")
form1.sqlqry ("update kontakt set telfaxhandy='nix'")
MousePointer = 0

End Sub

Private Sub Command16_Click()
MousePointer = 11
form1.sqlqry ("delete from kontakt where ( isnull(name)=true and isnull(tel)=true and isnull(fax)=true and isnull(email)=true and handy='0' )")
form1.sqlqry ("delete from kontakt where ( trim(name)='' and trim(tel)='' and trim(fax)='' and trim(email)='' and (trim(handy='' or trim(handy='0') )")
form1.sqlqry ("delete from auftritthigru where ( isnull(felddaten)=true )")
form1.sqlqry ("delete from kontakt where ( trim(name)="""" and trim(tel)="""" and trim(fax)="""" and trim(email)="""" and handy='0' )")
form1.sqlqry ("delete from auftritthigru where ( trim(felddaten)="""" )")
MousePointer = 0
End Sub


Private Sub Command17_Click()
Dim r As Recordset, i%, o%, s As Recordset

o% = FreeFile
Open form1.s0dir() & "\handyex.csv" For Output As #o%
Print #o%, """Name"",""weitere"""
c$ = "SELECT adresse.id as id,name,tel,handy FROM adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid WHERE (((kid='-1') or isnull(kid)) and (( ((adresstyp.typ)='nokia') )))  order by adresse.PLZ"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  num$ = trm(r!tel)
  nam$ = ""
  c$ = "select felddaten from auftritthigru where auftrittstyp='nokia' and feldname='NokiaName' and auftrittsid='" & r!id & "'"
  Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If Not s.EOF Then
    nam$ = trm(s!felddaten)
  End If
  If nam$ = "" Then nam$ = r!name
  If num$ <> "" Then Print #o%, """" & nam$ & """,""" & num$ & """"
  nam$ = nam$ & " " & "Handy"
  num$ = trm(r!handy)
  If num$ <> "" Then Print #o%, """" & nam$ & """,""" & num$ & """"
  r.MoveNext
Wend

c$ = "SELECT kontakt.name as name ,kontakt.id as id,kontakt.vid as vid ,kontakt.tel as tel,kontakt.handy as handy FROM (kontakt INNER JOIN adresstyp ON kontakt.id = adresstyp.kid) INNER JOIN adresse ON kontakt.vid = adresse.ID where instr(lcase(kontakt.name),'')>0 and ( ((adresstyp.typ)='nokia') )  order by adresse.PLZ"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  num$ = trm(r!tel)
  nam$ = ""
  c$ = "select felddaten from auftritthigru where auftrittstyp='nokia' and feldname='NokiaName' and auftrittsid='" & r!vid & r!id & "'"
  Set s = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
  If Not s.EOF Then
    nam$ = trm(s!felddaten)
  End If
  If nam$ = "" Then nam$ = r!name
  If num$ <> "" Then Print #o%, """" & nam$ & """,""" & num$ & """"
  nam$ = nam$ & " " & "Handy"
  num$ = trm(r!handy)
  If num$ <> "" Then Print #o%, """" & nam$ & """,""" & num$ & """"
  r.MoveNext
Wend

Close #o%
X = Shell("notepad.exe " & form1.s0dir() & "\handyex.csv", 1)

End Sub

Function mkvalid(r$) As String
  
  rid$ = strrepl(r$, Chr$(34), "")
  rid$ = strrepl(rid$, Chr$(13), "")
  rid$ = strrepl(rid$, Chr$(10), " ")
  rid$ = strrepl(rid$, ")", " ")
  rid$ = strrepl(rid$, "(", " ")
mkvalid = rid$

End Function

Private Sub Command18_Click()
Dim r As Recordset, s As Recordset, t As Recordset
stopfl = 0
List1.Clear
Set r = form1.sqla.OpenRecordset( _
    "SELECT * FROM auftritt where datum>='" + trm(Text1.text) + "' order by datum", dbOpenDynaset, dbReadOnly)

While Not r.EOF And stopfl = 0
  typ$ = r!auftrittstyp
  If typ$ <> "Neuer Auftritt" Then
    id$ = r!id
    If id$ <> "" Then
      Text1.text = r!datum
      DoEvents
      Call form1.check_tst(id$)
    End If
  End If
  r.MoveNext
Wend

List1.Clear

End Sub

Private Sub Command19_Click()
Dim r As Recordset, pl$
Dim s As Recordset, l$, n$

Set r = form1.sqla.OpenRecordset( _
    "SELECT id,name,strasse,ort FROM adresse order by id", dbOpenDynaset, dbReadOnly)
While Not r.EOF
  Text1.text = trm(r!id): DoEvents
  l$ = trm(r!name): n$ = crlfremove(l$)
  If l$ <> n$ Then
    List1.AddItem "update adresse set name='" + n$ + "' where id='" + r!id + "'"
  End If
  l$ = trm(r!ort): n$ = crlfremove(l$)
  If l$ <> n$ Then
    List1.AddItem "update adresse set ort='" + n$ + "' where id='" + r!id + "'"
  End If
  l$ = trm(r!strasse): n$ = crlfremove(l$)
  If l$ <> n$ Then
    List1.AddItem "update adresse set strasse='" + n$ + "' where id='" + r!id + "'"
  End If
  r.MoveNext
Wend
Set r = form1.sqla.OpenRecordset( _
    "SELECT id,name,strasse,ort FROM kontakt order by name", dbOpenDynaset, dbReadOnly)
While Not r.EOF
  l$ = trm(r!name): n$ = crlfremove(l$)
  Text1.text = n$: DoEvents
  If l$ <> n$ Then
    List1.AddItem "update kontakt set name='" + n$ + "' where id='" + r!id + "'"
  End If
  l$ = trm(r!ort): n$ = crlfremove(l$)
  If l$ <> n$ Then
    List1.AddItem "update kontakt set ort='" + n$ + "' where id='" + r!id + "'"
  End If
  l$ = trm(r!strasse): n$ = crlfremove(l$)
  If l$ <> n$ Then
    List1.AddItem "update kontakt set strasse='" + n$ + "' where id='" + r!id + "'"
  End If
  r.MoveNext
Wend
End Sub

Function crlfremove(k$) As String
Dim t$, d As Boolean, d1 As Boolean

t$ = k$
If t$ <> "" Then
  Do
    d = False
    Do
      d1 = False
      If t$ <> "" Then
        If Asc(Right(t$, 1)) = 10 Then
          t = Left(t$, Len(t$) - 1)
          d = True
          d1 = True
        End If
      End If
    Loop Until d1 = False
    Do
      d1 = False
      If t$ <> "" Then
        If Asc(Right(t$, 1)) = 13 Then
          t = Left(t$, Len(t$) - 1)
          d = True
          d1 = True
        End If
      End If
    Loop Until d1 = False
  Loop Until Not d
End If
crlfremove = t$

End Function
Private Sub Command2_Click()
Dim r As ADODB.Recordset, p$, prv$, c$, kid$, n%

Dim s As ADODB.Recordset, fnd As Boolean, ofn As Integer

MousePointer = 11: DoEvents
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM opt_allenummern;"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
prv$ = "": fnd = True: n% = 0
ofn = FreeFile
Open "c:\temp\liste.txt" For Output As ofn
While Not r.EOF
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  kid$ = trm(r!kid)
  If kid$ <> "-1" Then
    rrr = form1.adoopen(s, "SELECT name FROM kontakt where id='" + kid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    If s.EOF And rrr = 0 Then
      c$ = "delete FROM opt_allenummern where vid='" + r!vid + "' and kid='" + kid$ + "'"
      Debug.Print n%; " - "; c$
      Print #ofn, c$
      Call form1.sqlqry(c$)
      n% = n% + 1
    End If
  End If
  DoEvents
  r.MoveNext
Wend
Debug.Print n%
End If
Close ofn
MousePointer = 0

End Sub

Private Sub Command20_Click()
Dim r As ADODB.Recordset, p$, prv$, c$
Dim s As ADODB.Recordset, fnd As Boolean, ofn As Integer

MousePointer = 11: DoEvents
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM auftritthigru WHERE auftrittstyp='Freundeskreis' ORDER BY auftrittsid, FeldName;"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
prv$ = "": fnd = True
ofn = FreeFile
Open "L:\tmp\ttt\liste.txt" For Output As ofn
While Not r.EOF
  If prv$ <> r!auftrittsid Or (Not fnd) Then
    If (Not fnd) And prv$ = r!auftrittsid Then
      Debug.Print r!feldname + "=" + r!felddaten
      Print #ofn, r!feldname + "=" + r!felddaten
    Else
      prv$ = r!auftrittsid
      fnd = True
      Text1.text = r!auftrittsid
      DoEvents
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
      cmd$ = "select * from adresstyp where vid='" + r!auftrittsid + "'"
      rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
      If rrr = 0 Then
      If s.EOF Then
        Set s = New ADODB.Recordset
        s.CursorLocation = adUseServer
        cmd$ = "select * from kontakt where instr('" + r!auftrittsid + "',vid)=1 and instr('" + r!auftrittsid + "',id)>01"
        rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
        fnd = True
        If rrr = 0 Then
        If s.EOF Then
        Print #ofn, vbCrLf + r!auftrittsid
          fnd = False
        End If
        End If
      End If
      End If
    End If
  End If
  r.MoveNext
Wend
End If
Close ofn
MousePointer = 0
Text1.text = "0"
End Sub

Private Sub Command21_Click()
Call Command3_Click
Call Command15_Click
Do
  Call Command14_Click
  n% = List1.ListCount
  Call Command4_Click
Loop Until n% = 0

End Sub

Private Sub Command22_Click()
Dim rtmp As Recordset, hgr As Recordset, c$
Dim stmp As ADODB.Recordset
Dim r As ADODB.Recordset, mf$, kat$, fld$

MousePointer = 11
DoEvents
List1.Clear


Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_mussfeld')>0 or instr(owner,'sysvar_system_mussfeld')>0"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
While Not r.EOF
  
  mf$ = trm(r!wert)
  typ$ = cut_d1(mf$, "|")
  fld$ = cut_d2bis(mf$, "|")
  List1.AddItem typ$ + "->" + fld$

  sel$ = "SELECT id "
  sel$ = sel$ + "FROM adresse order by id"
  Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)
  While Not rtmp.EOF
    rid$ = trm(rtmp!id)
    Command22.Caption = rid$: DoEvents
    If form1.isoftype(rid$, typ$) <> "-1" Then
      c$ = higruget(rid$, "-1", typ$, fld$)
      If c$ = "(null)" Then
        Call form1.higruinsert(rid$, typ$, fld$, "(null)")
        List1.AddItem rid$ + " gesetzt": List1.ListIndex = List1.ListCount - 1: DoEvents
      End If
    End If
    rtmp.MoveNext
  Wend
Command22.Caption = "Mussfelder setzen: DoEvents"
  
  sel$ = "SELECT id,vid,name "
  sel$ = sel$ + "FROM kontakt order by vid,name"
  Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)
  While Not rtmp.EOF
    rid$ = trm(rtmp!vid) + trm(rtmp!id)
    Command22.Caption = trm(rtmp!name) + " " + trm(rtmp!id): DoEvents
      
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
    rrr = form1.adoopen(stmp, "SELECT id,wert FROM adresstyp where vid='" & trm(rtmp!vid) & "' and typ='" & typ$ & "' and kid='" & trm(rtmp!id) & "'", form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not stmp.EOF Then
      c$ = higruget(trm(rtmp!vid), trm(rtmp!id), typ$, fld$)
      If c$ = "(null)" Then
        Call form1.higruinsert(rid$, typ$, fld$, "(null)")
        List1.AddItem rid$ + " gesetzt": List1.ListIndex = List1.ListCount - 1: DoEvents
      End If
    End If
    rtmp.MoveNext
  Wend
  
  
  
  r.MoveNext
Wend
End If
Command22.Caption = "Mussfelder setzen: DoEvents"

MousePointer = 0

End Sub

Function validatefntmp(f) As String
Dim i%, r$, z$, l$, bsfn$

validatefntmp = f
r$ = ""
l$ = f
If LCase(Right$(l$, 4)) = ".msg" Then
  bsfn$ = Left(l$, Len(l$) - 4)
Else
  validatefntmp = ""
  Exit Function
End If
For i% = 1 To Len(bsfn$)
  z$ = Mid$(f, i%, 1)
  If z$ = "-" Or (LCase(z$) >= "a" And LCase(z$) <= "z") Or (LCase(z$) >= "0" And LCase(z$) <= "9") Then
      r$ = r$ + z$
  Else
      r$ = r$ + "_"
  End If
Next i%
validatefntmp = r$ + ".msg"
End Function


Private Sub Command23_Click()
   Dim r As Recordset, pl$
 
    cmd$ = "delete from auftritthigru where auftrittstyp='Tournee'"
    Call form1.sqlqry(cmd$)
Set r = form1.sqla.OpenRecordset( _
    "SELECT * FROM usr_tournee order by id", dbOpenDynaset, dbReadOnly)
While Not r.EOF
  c0$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,Feldname,FeldDaten) values('"
  List1.AddItem trm(r!id)
  List1.ListIndex = List1.ListCount - 1
  DoEvents
  If trm(r!enddatum) <> "" Then
    c$ = c0$ + form1.newid("auftritthigru", "id", 17) + "','" + trm(r!id) + "','Tournee','Enddatum','" + trm(r!enddatum) + "'"
    Call form1.sqlqry(c$)
  End If
  If trm(r!Solist) <> "" Then
    c$ = c0$ + form1.newid("auftritthigru", "id", 17) + "','" + trm(r!id) + "','Tournee','Solist','" + trm(r!Solist) + "'"
    Call form1.sqlqry(c$)
  End If
  If trm(r!orchester) <> "" Then
    c$ = c0$ + form1.newid("auftritthigru", "id", 17) + "','" + trm(r!id) + "','Tournee','Orchester','" + trm(r!orchester) + "'"
    Call form1.sqlqry(c$)
  End If
  If trm(r!veranstalter) <> "" Then
    c$ = c0$ + form1.newid("auftritthigru", "id", 17) + "','" + trm(r!id) + "','Tournee','Veranstalter','" + trm(r!veranstalter) + "'"
    Call form1.sqlqry(c$)
  End If
  If trm(r!dirigent) <> "" Then
    c$ = c0$ + form1.newid("auftritthigru", "id", 17) + "','" + trm(r!id) + "','Tournee','Dirigent','" + trm(r!dirigent) + "'"
    Call form1.sqlqry(c$)
  End If
  r.MoveNext
Wend
End Sub

Private Sub Command24_Click()
Dim r As Recordset, pl$
Dim s As Recordset, cmd$

Set r = form1.sqla.OpenRecordset( _
    "SELECT id,name,vid FROM kontakt order by vid", dbOpenDynaset, dbReadOnly)
While Not r.EOF
  Command24.Caption = trm(r!vid): DoEvents
  cmd$ = "select id from adresse where id='" + r!vid + "'"
  Set s = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
  If s.EOF Then
    Debug.Print r!name; " in adresse " + r!vid + " fehlt -> " + r!id + " wird gelöscht."
    cmd$ = "delete from kontakt where id='" + r!id + "'"
    Call form1.sqlqry(cmd$)
  End If
  r.MoveNext
Wend

Set r = form1.sqla.OpenRecordset( _
    "SELECT id,kid,vid,typ FROM adresstyp order by vid,kid,typ", dbOpenDynaset, dbReadOnly)
While Not r.EOF
  Command24.Caption = trm(r!vid): DoEvents
  cmd$ = "select id from adresse where id='" + r!vid + "'"
  Set s = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
  If s.EOF Then
    Debug.Print "Adresse " + r!vid + " fehlt (" + r!typ + ") -> " + r!id + " wird gelöscht."
    cmd$ = "delete from adresstyp where id='" + r!id + "'"
    Call form1.sqlqry(cmd$)
    cmd$ = "delete from auftritthigru where auftrittsid='" + r!vid + "' and auftrittstyp='" + r!typ + "'"
Debug.Print cmd$
    Call form1.sqlqry(cmd$)
  End If
  If r!kid <> "-1" Then
    cmd$ = "select id from kontakt where id='" + r!kid + "'"
    Set s = form1.sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
    If s.EOF Then
      Debug.Print "Kontakt " + r!kid + " in Adresse " + r!vid + " fehlt (" + r!typ + ") -> " + r!id + " wird gelöscht."
      cmd$ = "delete from adresstyp where id='" + r!id + "'"
      Call form1.sqlqry(cmd$)
      cmd$ = "delete from auftritthigru where auftrittsid='" + r!vid + r!kid + "' and auftrittstyp='" + r!typ + "'"
Debug.Print cmd$
      Call form1.sqlqry(cmd$)
    End If
  End If
  r.MoveNext
Wend
End Sub

Private Sub Command25_Click()
Dim r As Recordset, pl$
Dim s As Recordset

Set r = form1.sqla.OpenRecordset( _
    "SELECT id,name,vid,tel,fax,handy,email,url,position FROM kontakt order by name,vid,tel,fax,handy,email,url,position", dbOpenDynaset, dbReadOnly)
pl$ = ""
lkid = ""
While Not r.EOF
  l$ = trm(r!name) & trm(r!vid) & trm(r!tel) & trm(r!fax) & trm(r!handy) & trm(r!email) & trm(r!url) & trm(r!Position)
  Debug.Print l$
  If pl$ = l$ Then
    List1.AddItem "delete from adresstyp where kid='" & r!id & "'"
    List1.AddItem "delete from kontakt where id='" & r!id & "'"
    Set s = form1.sqla.OpenRecordset( _
       "SELECT * FROM dochist where kontakt='" & r!id & "'", dbOpenDynaset, dbReadOnly)
    While Not s.EOF
      List1.AddItem "update dochist set kontakt='" & lkid & "' where id='" & s!id & "';"
      s.MoveNext
    Wend
    List1.ListIndex = List1.ListCount - 1
    DoEvents
  Else
    pl$ = l$
    lkid = r!id
  End If
  r.MoveNext
Wend

End Sub

Private Sub Command26_Click()
c$ = "update kontakt set postanrede='Ms.' where instr(name,'Ms. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,4) where instr(name,'Ms. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set postanrede='Mr.' where instr(name,'Mr. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,4) where instr(name,'Mr. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set postanrede='Mrs.' where instr(name,'Mrs. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,5) where instr(name,'Mrs. ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set postanrede='Frau' where instr(name,'Frau ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,6) where instr(name,'Frau ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set postanrede='Herr' where instr(name,'Herr ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,6) where instr(name,'Herr ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set postanrede='Herrn' where instr(name,'Herrn ')=1;": Call form1.sqlqry(c$)
c$ = "update kontakt set name=mid(name,7) where instr(name,'Herrn ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set postanrede='Frau' where instr(name,'Frau ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set name=mid(name,6) where instr(name,'Frau ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set postanrede='Herr' where instr(name,'Herr ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set name=mid(name,6) where instr(name,'Herr ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set postanrede='Herrn' where instr(name,'Herrn ')=1;": Call form1.sqlqry(c$)
c$ = "update adresse set name=mid(name,7) where instr(name,'Herrn ')=1;": Call form1.sqlqry(c$)

c$ = "update adresse set name=trim(name);": Call form1.sqlqry(c$)
c$ = "update kontakt set name=trim(name);": Call form1.sqlqry(c$)


End Sub


Private Sub Command27_Click()
Dim frm$, c$, r As ADODB.Recordset, u$, o%, i As Integer, p%

p% = FreeFile
Open "adrlist.txt" For Input As #p%
u$ = form1.getuserid()

While Not EOF(p%)
  Line Input #p%, frm$
  On Error GoTo 0
  If frm$ <> "" Then
Debug.Print frm$
    frm$ = LCase(frm$)
    c$ = "select * from sysvars where owner='" & u$ & "' and wert='" & frm$ & "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    Call form1.dbg2f("Frmmain.Command11_Click:" & c$)
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If r.EOF Then
      c$ = "insert into sysvars (id,owner,wert) values('" & form1.newid("sysvars", "id", 36) & _
         "','blacklistdom:" & u$ & _
         "','" & frm$ & "')"
      Call form1.sqlqry(c$)
      c$ = "insert into phpop3clean_received_domains (domain) values('" & frm$ & "');"
      o% = FreeFile
      Open form1.mydir() + "\spamdoms.txt" For Append As #o%
      Print #o%, c$
      Close #o%
    End If
  End If
Wend
Close #p%

End Sub

Private Sub Command28_Click()
Dim rtmp As ADODB.Recordset, r2 As ADODB.Recordset, c$, ne$, rrr, fn$, r2id$

List1.Clear
c$ = "select * from dochist where lcase(right(docname,4))='.msg' and lcase(left(docname,2))='" + LCase(Left$(form1.s0dir(), 2)) + "' limit 0,10000;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
List1.Clear
While Not rtmp.EOF
  If nexist(rtmp!docname) Then
      Debug.Print "not found: " + rtmp!docname
      List1.AddItem rtmp!docname + " not found"
      List1.ListIndex = List1.ListCount - 1
      DoEvents
      If List1.ListCount > 1000 Then List1.Clear
  Else
  If InStr(LCase(rtmp!docname), LCase(form1.s0dir)) = 1 Then
    c$ = "select id from mailsafe where Message like '%" + strrepl(rtmp!docname, "\", "\\\\") + "%' or Message like '%" + strrepl(rtmp!docname, "\", "\\") + "%'"
    Set r2 = New ADODB.Recordset
    r2.CursorLocation = adUseServer
    r2.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
    r2id$ = ""
    If Not r2.EOF Then
      r2id$ = r2!id
    End If
    Debug.Print rtmp!docname + " (" + r2id$ + ")"
    fne$ = rtmp!docname + ".eml"
    rrr = 0
    If nexist(fne$) Then
      On Error Resume Next
      Call FileCopy(rtmp!docname, fne$)
      rrr = Err
      On Error GoTo 0
    End If
    If rrr = 0 Then
      c$ = "update dochist set docname='" + fne$ + "' where id='" + rtmp!id + "'"
      Call form1.sqlqry(c$)
      List1.AddItem c$
      List1.ListIndex = List1.ListCount - 1
      DoEvents
      If List1.ListCount > 1000 Then List1.Clear
      If r2id$ <> "" Then
        c$ = "update mailsafe set message='" + fne$ + "' where id='" + r2!id + "'"
        Call form1.sqlqry(c$)
      End If
      If Not nexist(fne$) Then
        On Error Resume Next
        Kill rtmp!docname
        On Error GoTo 0
      End If
    End If
  End If
  End If
  rtmp.MoveNext
Wend
End Sub

Private Sub Command29_Click()
Dim r As Recordset, c$

c$ = "select id,erstellt from mailsafe where erstellt like '__.__%';"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  p% = InStr(r!erstellt, " ") + 1
  If p% < Len(r!erstellt) Then t$ = trm(Mid$(r!erstellt, p%))
  c$ = "update mailsafe set erstellt='" & datum2sql(r!erstellt) & " " & t$ & "' where id='" & r!id & "';"
  List1.AddItem c$
  r.MoveNext
Wend
End Sub

Private Sub Command3_Click()
Dim rtmp As Recordset

List1.Clear

sel$ = "SELECT * "
sel$ = sel$ + "FROM (adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid) "
sel$ = sel$ + "INNER JOIN adresstypen ON adresstyp.typ = adresstypen.id "
sel$ = sel$ + "WHERE (((adresstyp.typ)='Künstler') "
sel$ = sel$ + "OR ((adresstyp.typ)='Orchester') "
sel$ = sel$ + "OR ((adresstyp.typ)='Partner') "
sel$ = sel$ + "OR ((adresstyp.typ)='Tourneeleitung') "
sel$ = sel$ + "OR ((adresstyp.typ)='Dirigent') );"

Set rtmp = form1.sqla.OpenRecordset(sel$, dbOpenDynaset, dbReadOnly)

While Not rtmp.EOF And List1.ListCount < 16000
  List1.AddItem "delete from adresstyp where vid='" + rtmp!vid + "' and typ='Person'"
  List1.AddItem "insert into adresstyp (id,vid,typ) values('" + form1.newid("adresstyp", "id", 20) + "','" + rtmp!vid + "','Person')"
  rtmp.MoveNext
Wend


End Sub


Private Sub Command31_Click()
Text4.text = encrypt(Text2.text, Text3.text)
End Sub

Private Sub Command32_Click()
Text2.text = decrypt(Text4.text, Text3.text)
End Sub

Private Sub Command33_Click()
Dim sqla As Database
Dim wrkJ As Workspace
Dim f_ldn$(0 To 199)
Dim t1yp$(0 To 199)

dbname$ = "anymdb.mdb"
If nexist(dbname$) Then
  MsgBox (dbname$ + " nicht gefunden")
  Exit Sub
End If
Set wrkJ = CreateWorkspace("", "Admin", "", dbUseJet)
Set sqla = wrkJ.OpenDatabase(dbname$, dbDriverCompleteRequired, False, dbpara$)
fn$ = form1.mydatadir() & "\" & dbname$ & "-sql.txt"
If exist(fn$) <> 0 Then
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
End If
o% = FreeFile
Open fn$ For Output As #o%
Debug.Print sqla.name; " mit"; sqla.TableDefs.Count - 1; " Tabellen"
For i = 0 To sqla.TableDefs.Count - 1
  flst$ = ""
  If Left$(LCase(sqla.TableDefs(i).name), 4) <> "msys" Then
    Debug.Print sqla.TableDefs(i).name
    List1.AddItem sqla.TableDefs(i).name
    List1.ListIndex = List1.ListCount - 1
    DoEvents
    For j = 0 To 99: f_ldn$(j) = "": Next j
    c$ = ""
    adl$ = ""
    For j = 0 To sqla.TableDefs(i).Fields.Count - 1
      ftyp = sqla.TableDefs(i).Fields(j).Type
      fnam = sqla.TableDefs(i).Fields(j).name
      isid = False
      If InStr(fnam, "Konzertklei") = 0 Then
        If LCase(fnam) = "id" Then
          adl$ = adl$ + vbCrLf + "alter table " + sqla.TableDefs(i).name + " ADD PRIMARY KEY (" + fnam + ");"
          isid = True
        Else
          If InStr(LCase(fnam), "id") > 0 Then
            adl$ = adl$ + vbCrLf + "alter table " + sqla.TableDefs(i).name + " ADD INDEX " + fnam + " (" + fnam + ");"
          End If
        End If
      End If
      fsiz = sqla.TableDefs(i).Fields(j).Size
      If flst$ <> "" Then flst$ = flst$ & ","
      flst$ = flst$ & fnam
      f_ldn$(j) = fnam
      Debug.Print ftyp; "; "; fnam; "; s="; fsiz
      If ftyp = 0 Then ftyp = 3
      t1yp$(j) = trm(ftyp)
      sqlt$ = ""
      Select Case ftyp
        Case 20: sqlt$ = "bigint"
        Case 10: sqlt$ = "varchar"
        Case 12: sqlt$ = "longtext"
        Case 8: sqlt$ = "timestamp": fsiz = 14
        Case 3: sqlt$ = "smallint": fsiz = 14
        Case 7: sqlt$ = "double": fsiz = 0
        Case 5: sqlt$ = "int"
        Case 4: sqlt$ = "int"
        Case 2: sqlt$ = "int"
        Case 1: sqlt$ = "int"
        Case Else
      End Select
      If sqlt$ = "" Then
        MsgBox "unbekannter typ: " & ftyp & "; " & fnam & "; s=" & fsiz
        End
      Else
        If c$ <> "" Then c$ = c$ & ", "
        c$ = c$ & fnam & " " & sqlt$
        If fsiz <> 0 Then
          c$ = c$ & "(" & trm(fsiz) & ")"
          If isid Then c$ = c$ + " NOT NULL"
        End If
      End If
    Next j
    c1$ = "create table " & sqla.TableDefs(i).name & " ( " & c$ & " ) TYPE = MYISAM;"
    Debug.Print c1$
    Print #o%, c1$
    If adl$ <> "" Then
      Debug.Print adl$
      Print #o%, adl$
    End If
    c$ = "select * from " & sqla.TableDefs(i).name
    Set r = sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    While Not r.EOF
      werte$ = ""
      For j = 0 To sqla.TableDefs(i).Fields.Count - 1
        'If trm("" & r.Fields(j).value) <> "" Then
        'ff$ = "'": If t1yp$(j) = "3" Or t1yp$(j) = "4" Or t1yp$(j) = "7" Then ff$ = ""
        ff$ = "'"
        If ff$ = "" Then
          werte$ = werte$ & strrepl(trm("" & r.Fields(j).value), ",", ".") & ","
        Else
          werte$ = werte$ & ff$ & r.Fields(j).value & ff$ & ","
        End If
        'End If
      Next j
      werte$ = trm(Left$(werte$, Len(werte$) - 1))
      While Right$(werte$, 1) = ",": werte$ = Left$(werte$, Len(werte$) - 1): Wend
      werte$ = strrepl(werte$, "\", "||--||")
      werte$ = strrepl(werte$, "||--||", "\\")
      c$ = "insert into " & sqla.TableDefs(i).name & " (" & flst$ & ") values(" & werte$ & ");"
'      Debug.Print c$
      Print #o%, c$
      r.MoveNext
    Wend
  End If
Next i
Close #o%
X = Shell("notepad.exe " & fn$, 1)

End Sub


Private Sub Command34_Click()
Dim c$

c$ = "update opt_checks set ownr='' where ownr='||'"
Call form1.sqlqry(c$)

End Sub

Private Sub Command35_Click()
Dim rtmp As ADODB.Recordset, r2 As ADODB.Recordset, c$, ne$, rrr, fn$, r2id$

List1.Clear
c$ = "select * from dochist where lcase(right(docname,4))='.txt' and doctyp='Emaileingang';"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
List1.Clear
While Not rtmp.EOF
  If nexist(rtmp!docname) Then
      Debug.Print "not found: " + rtmp!docname
      List1.AddItem rtmp!docname + " not found"
      List1.ListIndex = List1.ListCount - 1
      DoEvents
      If List1.ListCount > 1000 Then List1.Clear
  Else
  If InStr(LCase(rtmp!docname), LCase(form1.s0dir)) = 1 Then
    c$ = "select id from mailsafe where Message like '%" + strrepl(rtmp!docname, "\", "\\\\") + "%' or Message like '%" + strrepl(rtmp!docname, "\", "\\") + "%'"
    Set r2 = New ADODB.Recordset
    r2.CursorLocation = adUseServer
    r2.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
    r2id$ = ""
    If Not r2.EOF Then
      r2id$ = r2!id
    End If
    Debug.Print rtmp!docname + " (" + r2id$ + ")"
    List1.AddItem "(" + r2id$ + ") " + rtmp!docname
    List1.ListIndex = List1.ListCount - 1
    fne$ = rtmp!docname + ".eml"
    rrr = 0
    If nexist(fne$) Then
      On Error Resume Next
      Call FileCopy(rtmp!docname, fne$)
      rrr = Err
      On Error GoTo 0
    End If
    If rrr <> 0 Then
      List1.AddItem "copy failed (" + r2id$ + ") " + rtmp!docname
      List1.AddItem "ziel: (" + fne$
      List1.ListIndex = List1.ListCount - 1
      rtmp.MoveLast
    Else
      List1.AddItem "ok (" + r2id$ + ") " + rtmp!docname
      List1.ListIndex = List1.ListCount - 1
      c$ = "update dochist set docname='" + fne$ + "' where id='" + rtmp!id + "'"
      Call form1.sqlqry(c$)
      List1.AddItem c$
      List1.ListIndex = List1.ListCount - 1
      DoEvents
      If List1.ListCount > 1000 Then List1.Clear
      If r2id$ <> "" Then
        c$ = "update mailsafe set message='" + fne$ + "' where id='" + r2!id + "'"
        Call form1.sqlqry(c$)
      End If
      If Not nexist(fne$) Then
        On Error Resume Next
        Kill rtmp!docname
        On Error GoTo 0
      End If
    End If
  End If
  End If
  rtmp.MoveNext
Wend

End Sub

Private Sub Command4_Click()
stopfl = 0
If List1.ListCount > 0 Then List1.ListIndex = 0
While List1.ListCount > 0 And stopfl = 0
  form1.sqlqry (List1.List(0))
  List1.RemoveItem 0
  DoEvents
Wend

End Sub

Private Sub Command5_Click()
Dim o%, tr$, p%, q%

MousePointer = 11
On Error Resume Next
Kill "rtfs.ini"
On Error GoTo 0
o% = FreeFile
Open "rtfs.ini" For Output As #o%
tr$ = Dir(form1.s0dir() & "\apdemo.mdb.rtf\*.*")
While tr$ <> ""
  Print #o%, basename(tr$, "")
  p% = FreeFile
  Open form1.s0dir() & "\apdemo.mdb.rtf\" + tr For Input As #p%
  While Not EOF(p%)
    Line Input #p%, l$
    Print #o%, l$
  Wend
  Print #o%, "***EOF***AGENCYPROF***"
  Close #p%
  tr$ = Dir
Wend
Close #o%
MousePointer = 0

End Sub

Private Sub Command6_Click()
Dim r As Recordset, s As Recordset, t As Recordset
stopfl = 0
form1.currentconfmode = "ok, deleted"
Set r = form1.sqla.OpenRecordset( _
    "SELECT * FROM auftritt where datum>='" + trm(Text1.text) + "' order by datum", dbOpenDynaset, dbReadOnly)

Do
If List1.ListCount > 0 Then Call Command4_Click
While Not r.EOF And List1.ListCount < 30000 And stopfl = 0
  typ$ = r!auftrittstyp
  If typ$ <> "Neuer Auftritt" Then
    id$ = r!id
    If id$ <> "" Then
      Text1.text = r!datum
      DoEvents
      Set t = form1.sqla.OpenRecordset( _
        "SELECT * FROM usr_" & utabn(typ$) + " where id='" + id$ + "'", dbOpenDynaset, dbReadOnly)
      If t.EOF Then
        List2.Clear
        List3.Clear
        List2.AddItem "id"
        List3.AddItem id$
        Set s = form1.sqla.OpenRecordset( _
          "SELECT * FROM auftritthigru where auftrittsid='" + id$ + "' and auftrittstyp='" & typ$ & "' order by auftrittstyp, feldname", dbOpenDynaset, dbReadOnly)
        tt$ = ""
        While Not s.EOF
          If tt$ <> "" & s!auftrittstyp & s!feldname And s!feldname <> "zzzsysez" Then
            tt$ = "" & s!auftrittstyp & s!feldname
            If Not IsNull(s!felddaten) Then
              List2.AddItem s!feldname
              List3.AddItem s!felddaten
              DoEvents
            End If
          Else
'            Debug.Print "multiple or undefined background "; tt$; "-"; id$
          End If
          s.MoveNext
        Wend
        If List2.ListCount = 1 Then
          form1.sqlqry ("delete from auftritthigru where auftrittsid='" + id$ + "'")
          form1.sqlqry ("delete FROM usr_" & utabn(typ$) + " where id='" + id$ + "'")
          form1.sqlqry ("delete FROM auftritt where id='" + id$ + "'")
        Else
          cmd$ = "insert into usr_" & utabn(typ$) + " ("
          For i% = 0 To List2.ListCount - 1
            cmd$ = cmd$ + List2.List(i%)
            If i% < List2.ListCount - 1 Then
              cmd$ = cmd$ + ","
            End If
          Next i%
          cmd$ = cmd$ + ") values("
          For i% = 0 To List3.ListCount - 1
            li$ = List3.List(i%)
            If InStr(li$, Chr$(13) + Chr$(10)) > 0 Then
              If form1.higruzeilen(typ$, List2.List(i%)) < 2 Then
                li$ = strrepl(li$, Chr$(13) + Chr$(10), " ")
              End If
            End If
            cmd$ = cmd$ + "'" + li$ + "'"
            If i% < List3.ListCount - 1 Then
              cmd$ = cmd$ + ","
            End If
          Next i%
          cmd$ = cmd$ + ")"
          List1.AddItem cmd$
          List1.ListIndex = List1.ListCount - 1
          Call form1.sqlqry(cmd$)
          Call form1.check_tst(id$)
        End If
        DoEvents
      End If
    End If
  End If
  r.MoveNext
Wend
Loop Until r.EOF Or stopfl <> 0
List1.Clear
form1.currentconfmode = ""
End Sub

Private Sub Command7_Click()
Call verwalt_public.Command26_Click
End Sub

Private Sub Command8_Click()
stopfl = 1
End Sub

Private Sub Command9_Click()
MousePointer = 11: DoEvents
For i% = 0 To form1.sqla.TableDefs.Count - 1
  If Left$(form1.sqla.TableDefs(i%).name, 4) = "usr_" Then
    form1.sqlqry ("delete from " & form1.sqla.TableDefs(i%).name)
  End If
Next i%
MousePointer = 0
End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
Randomize
stopfl = 0
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
'dbpara$ = form1.getconnstr()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
'End If
End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub

Private Sub Timer1_Timer()
Label1.Caption = List1.ListCount
End Sub

Private Function higruget(id$, kid$, typ$, f$) As String
Dim c$, k$, r As ADODB.Recordset

higruget = ""
k$ = kid$: If k$ = "-1" Then k$ = ""
c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + k$ + "' and auftrittstyp='" + typ$ + "' and feldname='" + f$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  higruget = trm(r!felddaten)
Else
  higruget = "(null)"
End If
End Function

