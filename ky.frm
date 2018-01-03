VERSION 5.00
Begin VB.Form ky 
   Caption         =   "Jahresübersicht"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Beenden 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Picture         =   "ky.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Kalender schliessen"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label mleg 
      Caption         =   "legende"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "ky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mnams$(1 To 12), wdays$(7), tag0$
Dim lip%(366), lipv%, adrgrpselcache$, adrgrpnoselcache$, ypixperentry%
Dim c_typ$(366, 299), c_bez$(366, 299), c_stat(366, 299) As Long, c_id$(366, 299)
Dim c_col(366, 299) As Long, c_white, c_black, mscale, dscale, monat0, yahr0, day0

Sub rdsels()
Dim gw$, fsel$, bisi%, kid$, old As Variant, tpid$
Dim dv$, db$, selstr$, cmd$, nosel As Integer, shwpriv As Boolean
Dim r As ADODB.Recordset, rrr, c_stat0 As Long, optcol As Boolean, col As Long
Dim prvid$, offs%, i%, cbz$, gw1$, ent2 As Boolean, wasalles As String
Dim dkz As Boolean, noshow As Boolean, tpidokcache$, tpidnokcache$
Dim prz As Boolean, pvon, pbis, idat, d0, yline, x0, xoff, yoff, mnum

Dim d2infile As String, d2insub As String
d2infile = "k3": d2insub = "rdsels"
c_stat0 = RGB(255, 255, 255)
dv$ = trm(yahr0) + "-" + trm(monat0) + "-01"
db$ = datum2sql(CDate(dv$) + 366)
For offs% = 0 To 366: lip%(offs%) = -1: Next offs%
d0 = CDate(dv$): day0 = d0
old = d0
x0 = mleg(0).Width

prz = False
'If form1.getusersetting("Projektezeigen", "nein") = "ja" Then
If False Then
  prz = True
  cmd$ = "select * from tplan where (Hauptperson<>'Dekade') and (von>='" + dv$ + "' and von<='" + db$ + "') or (bis>='" + dv$ + "' and bis<='" + db$ + "') or (von<'" + dv$ + "' and bis>'" + db$ + "')"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
  While Not r.EOF
    On Error Resume Next
    pvon = Max(CDate(r!von), CDate(old))
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then pvon = CDate(old)
    On Error Resume Next
    pbis = MyMin(CDate(r!bis), CDate(old) + 366)
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      pbis = MyMin(CDate(r!von), CDate(old) + 366)
    End If
    For idat = pvon To pbis
      offs% = idat - old
      lip%(offs%) = lip%(offs%) + 1
      c_id$(offs%, lip%(offs%)) = r!id
      c_typ$(offs%, lip%(offs%)) = r!hauptperson
      c_stat(offs%, lip%(offs%)) = c_stat0
      c_col(offs%, lip%(offs%)) = form1.projektfarbe(trm(r!hauptperson))
      c_bez$(offs%, lip%(offs%)) = "Projekt: " + trm(r!id)
'      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry%)-(80, (lip%(offs%) + 1) * ypixperentry%), c_stat(offs%, lip%(offs%)), BF
'      p1(offs%).Line (80, (lip%(offs%) + 2) * ypixperentry%)-(0, (lip%(offs%) + 1) * ypixperentry%), c_col(offs%, lip%(offs%)), BF
'      p1(offs%).Print c_bez$(offs%, lip%(offs%))
      'Linie drüber
'      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry% - 1)-(0, (lip%(offs%) + 2) * ypixperentry% - 1), 0
    Next idat
    DoEvents
    r.MoveNext
  Wend
  End If
End If
On Error GoTo exrds
selstr$ = ""
selstr$ = selstr$ + "((datum>='" + dv$ + "' and datum<='" + db$ + "')) "
gw1$ = kc.getwho()
optcol = False
If kc.selct(2).ListCount = 0 And gw1$ = "" Then
  gw$ = kc.getwhere()
  wasalles = "id as aid,astatus,TourneeplanID,datum as adatum,auftritt.zeit as azeit, bezeichnung as abez,ort as aort, auftrittstyp as atyp "
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    wasalles = wasalles + ", optkalcolor as tf "
    optcol = True
  End If
  cmd$ = "SELECT " + wasalles + " from auftritt "
  If gw$ = "" Then
    gw$ = "where "
  Else
    gw$ = gw$ + " and "
  End If

  cmd$ = cmd$ + gw$ + selstr$
Else
  wasalles = "auftritt.id as aid,astatus,auftritt.TourneeplanID,auftritt.datum as adatum,auftritt.zeit as azeit,auftritt.bezeichnung as abez,auftritt.ort as aort, auftritthigru.auftrittstyp as atyp, auftritthigru.FeldName, auftritthigru.Felddaten "
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    wasalles = wasalles + ", auftritt.optkalcolor as tf "
    optcol = True
  End If
  cmd$ = "SELECT " + wasalles
  cmd$ = cmd$ + " FROM auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid "
  gw$ = kc.getwhere()
  If gw$ = "" Then
    cmd$ = cmd$ + " Where "
  Else
    cmd$ = cmd$ + gw$ + " and "
  End If
  nosel = 1
  For i% = 0 To kc.selct(2).ListCount - 1
    If kc.selct(2).Selected(i%) = True Then
      nosel = 0
      fsel$ = kc.selct(2).List(i%)
      i% = kc.selct(2).ListCount
    End If
  Next i%
  bisi% = 30
  If nosel = 1 Then
    kid$ = Trim("" & kc.selct(2).List(0))
    kid$ = "(instr(FeldDaten,'" + kid$ + "')>0) "
    For i% = 1 To kc.selct(2).ListCount - 1
      kid$ = kid$ + "or (instr(FeldDaten,'" + Trim("" & kc.selct(2).List(i%)) + "')>0) "
      bisi% = bisi% - 1
      If bisi% < 0 Then i% = kc.selct(2).ListCount - 1
    Next i%
  Else
    kid$ = fsel$
    kid$ = "(instr(FeldDaten],'" + kid$ + "')>0) "
    bisi% = kc.selct(2).ListCount - 1: If bisi% > 20 Then bisi% = 20
    For i% = 0 To bisi%
      If kc.selct(2).Selected(i%) = True And kc.selct(2).List(i%) <> fsel$ Then
        kid$ = kid$ + "or (instr(FeldDaten,'" + Trim("" & kc.selct(2).List(i%)) + "')>0) "
        bisi% = bisi% - 1
        If bisi% < 0 Then i% = kc.selct(2).ListCount - 1
      End If
    Next i%
  End If
  cmd$ = cmd$ + " ( " + kid$ + ") and  "
  cmd$ = cmd$ + selstr$
End If
cmd$ = cmd$ + " ORDER BY auftritt.datum,auftritt.zeit"
dkz = False
If form1.getusersetting("Dekadenzeigen", "nein") = "ja" Then dkz = True
If form1.getusersetting("Privateszeigen", "nein") = "ja" Then shwpriv = True
'daten selektieren
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  cmd$ = strrepl(cmd$, ",astatus", "")
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Sub
End If
prvid$ = "-x"
adrgrpselcache$ = "|"
adrgrpnoselcache$ = "|"
tpidokcache$ = ""
tpidnokcache$ = ""
While Not r.EOF
  ent2 = True
  noshow = False
  If Not dkz Then
    tpid$ = trm(r!TourneeplanID)
    If InStr(tpidnokcache$, tpid$) > 0 Then noshow = True
    If Not noshow And tpid$ <> "" And tpid$ <> "-1" And InStr(tpidokcache$, tpid$) = 0 Then
      If form1.projekttyp(tpid$) = "Dekade" Then
        noshow = True
        tpidnokcache$ = tpidnokcache$ + " " + tpid$
      Else
        tpidokcache$ = tpidokcache$ + " " + tpid$
      End If
    Else
      tpidokcache$ = tpidokcache$ + " " + tpid$
    End If
  End If
  If gw1$ <> "" Then ent2 = adrisinselectedgroup(trm(r!felddaten), gw1$)
  If r!aid <> prvid$ And ent2 Then
    prvid$ = r!aid
    If Not shwpriv And trm(r!atyp) = "Privat" Then noshow = True
    If Not noshow Then
      On Error Resume Next
      offs% = CDate(datfromsql(r!adatum)) - CDate(old)
      On Error GoTo 0
      lip%(offs%) = lip%(offs%) + 1
      c_id$(offs%, lip%(offs%)) = r!aid
      If Not IsNull(r!atyp) Then
        c_typ$(offs%, lip%(offs%)) = r!atyp
        If optcol Then
          c_col(offs%, lip%(offs%)) = Val(trm0(r!tf))
        Else
          c_col(offs%, lip%(offs%)) = -1
        End If
        If c_col(offs%, lip%(offs%)) <= 0 Then c_col(offs%, lip%(offs%)) = form1.get_eventcolor(r!atyp)
        On Error Resume Next
        c_stat(offs%, lip%(offs%)) = form1.get_eventstatuscolor(r!astatus)
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then c_stat(offs%, lip%(offs%)) = c_stat0
      End If
      cbz$ = trm(r!aort & " " & r!abez)
      If r!TourneeplanID <> -1 Then cbz$ = cbz$ & " " & r!TourneeplanID
      If trm(r!azeit) <> "" Then cbz$ = cbz$ & " " & r!azeit & " h"
      c_bez$(offs%, lip%(offs%)) = cbz$
    End If
  End If
  r.MoveNext
Wend
lipv% = 1
'Call buttonset
exrds:
On Error GoTo 0

End Sub

Private Sub Beenden_Click()
Unload Me

End Sub


Private Sub Form_Load()
Dim d

ypixperentry% = 18
If form1.getusersetting("datumsformat", "de") = "de" Then
  d = CDate("1." & (kc.Combo2.ListIndex + 1) & "." & (kc.Combo3.ListIndex + kc.yyyy0))
Else
  d = CDate("1/" & (kc.Combo2.ListIndex + 1) & "/" & (kc.Combo3.ListIndex + kc.yyyy0))
End If
monat0 = kc.Combo2.ListIndex + 1
yahr0 = kc.Combo3.ListIndex + kc.yyyy0
mscale = 100
dscale = 100
c_white = RGB(255, 255, 255)
c_black = RGB(0, 0, 0)
Me.ForeColor = RGB(0, 0, 255)
Me.Font.Bold = False
Me.Caption = transe("Jahresübersicht")
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
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Width = form1.mylastwidth(Me.name, 1)
Me.Height = form1.mylastheight(Me.name, 1)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Top = Me.Height / 3
  Me.Left = Me.Width / 3
End If
Call form1.formpos(Me)
Show

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d%, M%, x0

x0 = mleg(0).Width
'Label1(0).Caption = trm(Int(X)) + "/" + trm(Int(Y))
d% = Int((X - x0) / dscale) + 1
M% = Int(Y / mscale) + 1
Label1(0).Caption = trm(d%) + "." + trm(M%)
End Sub

Private Sub Form_Resize()
Dim i%, x0, mnum As Integer, ady

If Height < 4000 Then
  Height = 4000
End If
x0 = mleg(0).Width
ScaleWidth = 3200 + x0
ScaleHeight = 1200 + Beenden.Height
Cls
Font.Size = 10
ForeColor = RGB(22, 22, 22)
Call buttonset
For i% = 1 To 12
  mnum = (i% + monat0 - 2) Mod 12
  ady = 0: If mnum = 11 Then ady = -1
  Me.Line (x0, (i% * mscale))-(31 * dscale + x0, (i% * mscale) + ady), c_black, BF
  If i% > 1 Then mleg(i% - 1).Width = mleg(0).Width
  mleg(i% - 1).Left = 0: mleg(i% - 1).Top = (i% - 1) * mscale + 0.1 * mscale
  mleg(i% - 1).Caption = transe(mnams$(mnum + 1))
Next i%
For i% = 1 To 31
  Me.Line ((i% * dscale) + x0, 0)-((i% * dscale) + x0, 12 * mscale), c_black, BF
  Me.PSet (i% * dscale, 12 * mscale)
  Me.Print trm(i%)
Next i%
Call rdsels
Call rdrw
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)
exuld:
On Error GoTo 0

End Sub

Sub buttonset()
'up4.Left = ScaleWidth - up4.Width
Beenden.Top = ScaleHeight - Beenden.Height
Beenden.Left = 0
Label1(0).Left = ScaleWidth - Label1(0).Width
Label1(0).Top = ScaleHeight - Label1(0).Height
End Sub

Function adrisinselectedgroup(i$, selstr$) As Boolean
Dim r As ADODB.Recordset, cmd$, rrr

Dim d2infile As String, d2insub As String
d2infile = "k3": d2insub = "adrisinselectedgroup"
adrisinselectedgroup = False
If InStr(adrgrpselcache$, "|" & i$ & "|") > 0 Then
  adrisinselectedgroup = True
  GoTo exfu
End If
If InStr(adrgrpnoselcache$, "|" & i$ & "|") > 0 Then
  adrisinselectedgroup = False
  GoTo exfu
End If
cmd$ = "select grpid from adressgruppen where adressid='" & i$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  While Not r.EOF
    If InStr(selstr$, "|" & r!grpid & "|") > 0 Then
      adrisinselectedgroup = True
      adrgrpselcache$ = adrgrpselcache$ & i$ & "|"
      GoTo exfu
    End If
    r.MoveNext
  Wend
  adrisinselectedgroup = False
  adrgrpnoselcache$ = adrgrpnoselcache$ & i$ & "|"
  GoTo exfu
Else
  adrisinselectedgroup = False
  adrgrpnoselcache$ = adrgrpnoselcache$ & i$ & "|"
  GoTo exfu
End If
adrisinselectedgroup = True
exfu:
End Function

Sub rdrw()
Dim cmd$, xoff, yoff, yline, offs%, i%, d0, mnum, x0

d0 = day0
x0 = mleg(0).Width
For offs% = 0 To 366
   For i% = 0 To lip%(offs%)
      cmd$ = datum2sql(CDate(d0 + offs%))
      xoff = Val(Right$(cmd$, 2)) - 1
      yoff = Val(Mid$(cmd$, 6, 2)) - 1
      mnum = (yoff + monat0 + 1) Mod 12
      yline = mnum * mscale
      Me.Line (x0 + 1 + xoff * dscale, yline + 1 + i% * ypixperentry%)-(x0 - 1 + (xoff + 1) * mscale, yline + 1 + ypixperentry% + i% * ypixperentry%), c_col(offs%, i%), BF
      Me.PSet (x0 + 1 + xoff * dscale, yline + 1 + i% * ypixperentry%)
      Me.Print Left$(c_bez$(offs%, i%), 5)
'      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry%)-(80, (lip%(offs%) + 1) * ypixperentry%), c_stat(offs%, lip%(offs%)), BF
'      p1(offs%).Line (80, (lip%(offs%) + 2) * ypixperentry%)-(0, (lip%(offs%) + 1) * ypixperentry%), c_col(offs%, lip%(offs%)), BF
'      p1(offs%).Print c_bez$(offs%, lip%(offs%))
      DoEvents
  Next i%
Next offs%

End Sub
