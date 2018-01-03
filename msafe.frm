VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form msafe 
   Caption         =   "Mailsafe"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   5265
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command8 
      Caption         =   "gelöschte"
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   4800
      Width           =   975
   End
   Begin VB.ComboBox ftrg 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   6840
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4200
      Picture         =   "msafe.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   20
      ToolTipText     =   "Markierte Nachrichten neu an mich versenden"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4920
      Picture         =   "msafe.frx":0B66
      Style           =   1  'Grafisch
      TabIndex        =   19
      ToolTipText     =   "Öffnet die Ordneransicht"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "in Ordner schieben"
      Enabled         =   0   'False
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
      Left            =   5640
      TabIndex        =   18
      ToolTipText     =   "verschiebt markierte Mails"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox bdat 
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox vdat 
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox sdet 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mail-Details zeigen"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Antworten"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
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
      Left            =   840
      TabIndex        =   11
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Datum der Email setzen"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   9240
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reindex"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      ToolTipText     =   "Daten für Volltextsuche neu aufbauen"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox shmax 
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Text            =   "20"
      Top             =   225
      Width           =   495
   End
   Begin VB.CheckBox monly 
      BackColor       =   &H00C0C0C0&
      Caption         =   "nur meine"
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   360
      Value           =   1  'Aktiviert
      Width           =   1095
   End
   Begin VB.CheckBox vsuch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Volltextsuche"
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox suchw 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Suchworter ( keine Umlaute )"
      Top             =   225
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   480
      Picture         =   "msafe.frx":1190
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Formular schiessen"
      Top             =   4560
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2280
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      View            =   3
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "bis"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "von"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label getroffn 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Caption         =   "Treffer:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      ToolTipText     =   "In Kontakten suchen"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Max. Treffer:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "In Kontakten suchen"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suche:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "In Kontakten suchen"
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   5055
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "msafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currkid$
Dim mineonly%, pcode As Integer
Dim hdr_status$(1 To 2), tm_brk%
Dim monams$(1 To 12)
Dim merunning As Integer


Private Sub bdat_Change()
'd2infile = "msafe": d2insub = "bdat_Change"
Call timerreset

End Sub

Private Sub bdat_DblClick()
'd2infile = "msafe": d2insub = "bdat_DblClick"
  With frmCalendar
    .init bdat, bdat.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      bdat.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar

End Sub

Private Sub Command1_Click()
'd2infile = "msafe": d2insub = "Command1_Click"
Unload msafe

End Sub

Private Sub Command18_Click()
'd2infile = "msafe": d2insub = "Command18_Click"
Call form1.handbuchcall("14-Mailsafe.htm")

End Sub


Private Sub Command2_Click()
Dim rrr
Dim i As Integer, id$, fn$, sli$, o%, hd%, l$, w$, p%, sq$, rc$
Dim r As ADODB.Recordset, p0%, j%, z$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "Command2_Click"
Call gd1.SetFocus
For i = gd1.ListItems.Count To 1 Step -1
  If (gd1.ListItems(i).Selected = True) Then
    id$ = gd1.ListItems(i)
    p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
    id$ = Mid$(id$, p% + 4)
    If id$ <> "" Then
      MousePointer = 11: DoEvents
      sq$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, sq$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        fn$ = r!message
        If exist(fn$) = 1 Then
          sli$ = ""
          o% = FreeFile
          Open fn$ For Input As #o%
          hd% = 1
          While Not EOF(o%)
            Line Input #o%, l$: l$ = trm(l$)
            If l$ = "" And hd% = 1 Then
              hd% = 0
            Else
              If hd% = 0 And InStr(l$, " ") > 0 Then
                l$ = trm(l$): rc$ = ""
                For j% = 1 To Len(l$)
                  z$ = Mid$(l$, j%, 1)
                  If (z$ >= "0" And z$ <= "9") Or (z$ >= "a" And z$ <= "z") Or (z$ >= "A" And z$ <= "Z") Then
                    rc$ = rc$ + z$
                  Else
                    rc$ = rc$ + " "
                  End If
                Next j%
                l$ = rc$
                While Len(l$) > 0
                  w$ = mkalphanum(LCase(word1(l$)))
                  p% = Len(w$)
                  If p% > 0 Then
                    l$ = trm(Mid(l$, p% + 1))
                    If p% > 2 And p% < 30 Then
                      If InStr(sli$, w$) = 0 Then sli$ = trm(sli$ & " " & w$)
                    End If
                  Else
                    l$ = trm(Mid$(l$, 2))
                  End If
                Wend
              End If
            End If
          Wend
          Close #o%
          If sli$ <> "" Then
            w$ = "update mailsafe set volltext='" & sli$ & "' where id='" & id$ & "'"
            Call form1.sqlqry(w$)
          End If
        End If
      End If
      MousePointer = 0: DoEvents
    End If
  End If
  gd1.ListItems(i).Selected = False
  DoEvents
Next i
End Sub

Private Sub Command3_Click()
Dim r As ADODB.Recordset, id$, p%, cmd$, fn$, o%, from$, l$, X
Dim s As ADODB.Recordset, l1l$
Dim i As Integer, hd%, j%, ndate, c$, rrr

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "Command3_Click"
If gd1.ListItems.Count <= 0 Then Exit Sub

For i = gd1.ListItems.Count To 1 Step -1
  If (gd1.ListItems(i).Selected = True) Then
    id$ = gd1.ListItems(i)
    p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
    id$ = Mid$(id$, p% + 4)


MousePointer = 11: DoEvents
cmd$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fn$ = r!message
  If exist(fn$) = 1 Then
    o% = FreeFile
    Open fn$ For Input As #o%
    hd% = 1
    While Not EOF(o%) And hd% = 1
      Line Input #o%, l$
      If Len(l$) = 0 Then
        hd% = 1
      Else
        If InStr(LCase(l$), "date: ") = 1 Then

p% = InStr(LCase(l$), ",")
If p% > 0 Then l$ = trm(Mid$(l$, p% + 1))
p% = InStr(LCase(l$), " +")
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
p% = InStr(LCase(l$), " -")
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
p% = InStr(LCase(l$), " pdt")
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
p% = InStr(LCase(l$), " gmt")
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
p% = InStr(LCase(l$), " edt")
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
p% = InStr(LCase(l$), "date: ")
If p% = 1 Then l$ = trm(Mid(l$, p% + 6))
For j% = 1 To 12
  p% = InStr(LCase(l$), monams$(j%))
  If p% > 0 Then
    l$ = strrepl(LCase(l$), monams(j%), "." & trm(j%) & ".")
    Exit For
  End If
Next j%
On Error Resume Next
ndate = CDate(l$)
p% = InStr(ndate, " ")
If p% > 0 Then
  l1l$ = Mid$(ndate, p%)
  ndate = datum2sql(word1(trm(ndate))) & l1l$
End If
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  c$ = "update mailsafe set erstellt='" & ndate & "' where id='" & id$ & "'"
  Call form1.sqlqry(c$)
  c$ = "select id,erstellt,docname from dochist where docname='" & strrepl(strrepl(fn$, "\", "//"), "//", "\\") & "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If s.EOF Then
  End If
  While Not s.EOF
    c$ = "update dochist set erstellt='" & ndate & "' where id='" & s!id & "'"
    Call form1.sqlqry(c$)
    s.MoveNext
  Wend
Else
  Debug.Print "cannot convert " & l$
End If
          hd% = 0
        End If
      End If
    Wend
    Close #o%
  Else
    MousePointer = 0: DoEvents
'    MsgBox (fn$ + " kann nicht gefunden werden.")
  End If
  MousePointer = 0: DoEvents
End If


  End If
Next i
Call rlist1("")
End Sub

Private Sub Command4_Click()
Dim rrr
Dim id$, aid$, kid$, p%, r As ADODB.Recordset, r1 As ADODB.Recordset, em$, tgi$, o%
Dim cmd$, hd%, brk%, l$, dop%, tadr$, lcount%

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "Command4_Click"
Call gd1_Click
DoEvents
tadr$ = form1.Combo1.text
Load smtp
DoEvents
Call smtp.Command2_Click
DoEvents
emailadrselect.Text1.text = tadr$
DoEvents
Call emailadrselect.Timer1_Timer
DoEvents
If emailadrselect.List2.ListCount > 0 Then
  emailadrselect.List2.ListIndex = 0
  DoEvents
  Call emailadrselect.List2_DblClick
  DoEvents
Else
  smtp.txtSendTo.text = tadr$
End If
Unload emailadrselect
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
MousePointer = 11: DoEvents
id$ = Mid$(id$, p% + 4)
cmd$ = "select * from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  tgi$ = r!message
  o% = FreeFile
  Open tgi$ For Input As #o%
  lcount% = 0
  hd% = 1
  brk% = 0
  smtp.txtMessageSubject = "AW: " & r!Subject
  cmd$ = "select kontakt,adresse from dochist where memoinhalt='" + r!id + "'"
Set r1 = New ADODB.Recordset
r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r1.EOF Then
    aid$ = r1!adresse
    kid$ = r1!kontakt
  End If
  While Not EOF(o%) And brk% = 0 And lcount% < 99
    Line Input #o%, l$
    If trm(l$) = "" And hd% = 1 Then
      l$ = form1.get_kontaktname_by_id(kid$)
      If l$ = "" Then
        l$ = form1.getnamebyid(aid$)
      End If
      l$ = vbCrLf & l$ & " schrieb:"
      hd% = 0
      dop% = 1
    End If
    If hd% = 0 And InStr(LCase(l$), "content-type:") = 1 Then
      dop% = 0
      While l$ <> ""
        If InStr(LCase(l$), "text/plain") Then dop% = 1
        'SMTP.txtMessageText = SMTP.txtMessageText & vbCrLf & "|    " & l$
        Line Input #o%, l$
      Wend
      If dop% = 0 Then
        l$ = "..."
        'brk% = 1
      End If
    End If
    If hd% = 0 And dop% = 1 Then
      smtp.txtMessageText = smtp.txtMessageText & vbCrLf & "> " & l$
      lcount% = lcount% + 1
    End If
  Wend
  Close #o%
End If
MousePointer = 0


End Sub

Private Sub Command5_Click()
'd2infile = "msafe": d2insub = "Command5_Click"
Unload trvw: DoEvents
Load trvw
On Error Resume Next
Call trvw.SetFocus
On Error GoTo 0
Command5.Enabled = False
Call trvw.setmode("mail")

End Sub

Private Sub Command6_Click()
Dim i, c$, rtmp As ADODB.Recordset, fn$, id$, p%, nfn$, fno$, rrr, nfn0$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "Command6_Click"
nfn0$ = ftrg.text
If trm(nfn0$) = "" Then Exit Sub

nfn0$ = form1.mydir() + "\mail\" + nfn0$
On Error Resume Next
Call gd1.SetFocus
On Error GoTo 0
For i = gd1.ListItems.Count To 1 Step -1
      If (gd1.ListItems(i).Selected = True) Then
        id$ = gd1.ListItems(i)
        p% = InStr(id$, "(ID:")
        If p% <> 0 Then
          id$ = Mid$(id$, p% + 4)
          If id$ <> "" Then
            c$ = "SELECT * from mailsafe where id='" + id$ + "';"
            Set rtmp = New ADODB.Recordset
            rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not rtmp.EOF Then
              fn$ = trm(rtmp!message)
              fno$ = FileName(fn$)
              nfn$ = nfn0$ + "\" + fno$
              If Not nexist(fn$) Then
                If fn$ <> nfn$ Then
                  On Error Resume Next
                  FileCopy fn$, nfn$
                  rrr = Err
                  On Error GoTo 0
                  If rrr = 0 And (Not nexist(nfn$)) Then
                    gd1.ListItems(i).Selected = False
                    c$ = "update mailsafe set message='" + nfn$ + "' where id='" + id$ + "';"
                    Call form1.sqlqry(c$)
                    On Error Resume Next
                    Kill fn$
                    On Error GoTo 0
                    DoEvents
                  Else
                    MsgBox transe("Datei") + " " + nfn$ + ", " + vbCrLf + transe("konnte nicht kopiert werden")
                  End If
                Else
                  gd1.ListItems(i).Selected = False
                End If
              End If
            End If
          End If
        End If
      End If
Next i

End Sub

Private Sub Command7_Click()
Dim rrr
Dim i%, p%, id$, c$, rtmp As ADODB.Recordset, X, fn$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "Command7_Click"
On Error Resume Next
Call gd1.SetFocus
On Error Resume Next
MousePointer = 11: DoEvents
For i% = 1 To gd1.ListItems.Count
  If gd1.ListItems(i%).Selected Then
    id$ = gd1.ListItems(i).text
    p% = InStr(id$, "(ID:")
    If p% > 0 Then
      id$ = Mid$(id$, p% + 4)
      c$ = "select * from mailsafe where id='" + id$ + "'"
      Set rtmp = New ADODB.Recordset
      rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not rtmp.EOF Then
        fn$ = rtmp!message
        If Not nexist(fn) Then
          Call form1.mailresend(fn$)
        End If
        DoEvents
      End If
    End If
  End If
Next i%
MousePointer = 0
End Sub

Private Sub Command8_Click()
Dim rrr
Dim idx%, id$, sq$, i As Integer
Dim r As ADODB.Recordset, ask%, p%, fn$

    For i = gd1.ListItems.Count To 1 Step -1
      gd1.ListItems(i).Selected = False
    Next i
    For i = gd1.ListItems.Count To 1 Step -1
        
        id$ = gd1.ListItems(i)
        p% = InStr(id$, "(ID:")
        If p% = 0 Then
          If nexist(id$) Then
            gd1.ListItems(i).Selected = True
          End If
        End If
        id$ = Mid$(id$, p% + 4)
        If id$ <> "" Then
          sq$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, sq$, form1.adoc, dbOpenDynaset, dbReadOnly)
          If Not r.EOF Then
            fn$ = form1.dupcheck(r!message)
            If nexist(fn$) Then
              gd1.ListItems(i).Selected = True
            End If
          End If
        End If
    
    Next i
    On Error Resume Next
    gd1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim colHeader, klrv%

'd2infile = "msafe": d2insub = "Form_Load"
merunning = 0
monams$(1) = "jan"
monams$(2) = "feb"
monams$(3) = "mar"
monams$(4) = "apr"
monams$(5) = "may"
monams$(6) = "jun"
monams$(7) = "jul"
monams$(8) = "aug"
monams$(9) = "sep"
monams$(10) = "oct"
monams$(11) = "nov"
monams$(12) = "dec"


tm_brk% = 0
hdr_status$(1) = transe("ungelesen")
hdr_status$(2) = transe("gelesen")
Timer1.Enabled = False
pcode = -1
axsResizer1.SaveControlPositions
    gd1.View = lvwReport
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Absender"), 2000)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum"), 1800)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Betreff"), 4000)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Status"), 2)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Eigentümer"), 1200)
    Set colHeader = gd1.ColumnHeaders.add(, , transe("Datei"), 1000)
gd1.Font.Size = form1.myfontsize()

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
mineonly% = 1
If form1.getsystemsetting("mailviewer") = "ja" Or form1.getusersetting("mailviewer") = "ja" Then
  klrv% = Val(form1.mylastFormVar(Me.name, "sdet", "0"))
  If klrv% <> 0 Then klrv% = 1
  sdet.value = klrv%
Else
  sdet.Enabled = False
End If
klrv% = Val(form1.mylastFormVar(Me.name, "vsuch", "0"))
If klrv% <> 0 Then klrv% = 1
vsuch.value = klrv%
msafe.Caption = transe("Mailsafe")
sdet.Caption = transe("Mail-Details zeigen")
Command4.Caption = transe("&Antworten")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command3.Caption = transe("Datum der Email setzen")
Command2.Caption = transe("Reindex")
Command2.ToolTipText = transe("Daten für Volltextsuche neu aufbauen")
monly.Caption = transe("nur meine")
vsuch.Caption = transe("Volltextsuche")
suchw.ToolTipText = transe("Suchworter ( keine Umlaute )")
Command1.ToolTipText = transe("Formular schliessen")
Label4.Caption = transe("bis")
Label2.Caption = transe("von")
getroffn.Caption = transe("Treffer:")
getroffn.ToolTipText = transe("In Kontakten suchen")
Label1.Caption = transe("Max. Treffer:")
Label1.ToolTipText = transe("In Kontakten suchen")
Label3.Caption = transe("Suche:")
Label3.ToolTipText = transe("In Kontakten suchen")
Show
Call rlist1("")
klrv% = Val(form1.mylastFormVar("trvw", "mailimmertreeview", "0"))
If klrv% = 1 Then Call Command5_Click
End Sub

Private Sub Form_Resize()
'd2infile = "msafe": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "msafe": d2insub = "Form_Unload"
Unload trvw
DoEvents
merunning = 0
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Public Sub rlist1(pfad As String)
Dim r As ADODB.Recordset, fnd%, vsu$, vbd$, vd$, bd$, rrr
Dim nurmeine$, c$, lvitem, shm, connwrd$, krit$, s$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "rlist1"
If merunning = 1 Then Exit Sub
merunning = 1
msafe.MousePointer = 11: DoEvents
If pfad <> "" Then
  Call rlist1bypath(pfad)
  merunning = 0
  MousePointer = 0
  Exit Sub
End If
shm = Val(shmax.text)
If shm < 1 Then shm = 1
fnd% = 0
If shm < 0 Then
  shm = 20
  shmax.text = trm(shm)
End If
nurmeine$ = ""
connwrd$ = " where "
If mineonly% <> 0 Then
  nurmeine$ = "where (mailsafe.owner='" & form1.getuserid() & "') "
  connwrd$ = " and "
End If
gd1.ListItems.Clear
s$ = trm(suchw.text)
krit$ = ""
vbd$ = ""
vd$ = "": If trm(vdat.text) <> "" Then vd$ = datum2sql(vdat.text)
bd$ = "": If trm(bdat.text) <> "" Then bd$ = datum2sql(bdat.text)
If vd$ <> "" Or bd$ <> "" Then
  If vd$ <> "" Then vbd$ = "erstellt >='" & vd$ & " 00:00:00'"
  If bd$ <> "" Then
    If vbd$ <> "" Then vbd$ = vbd$ & " and "
    vbd$ = vbd$ & "erstellt <= '" & bd$ & " 23:59:59' "
  End If
End If
'normales suchen
  
  If s$ <> "" Then
    krit$ = connwrd$ & "(mailsafe.Subject like '%" + LCase(s$) + "%' or mailsafe.Subject like '" + LCase(s$) + "%' " + _
                     " or mailsafe.frm like '%" + LCase(s$) + "%' or mailsafe.frm like '" + LCase(s$) + "%') "
'    krit$ = connwrd$ & "instr(lcase(mailsafe.Subject),'" + LCase(s$) + "')>0 " + _
'                     " or instr(lcase(mailsafe.frm),'" + LCase(s$) + "')>0 "
    connwrd$ = " or "
  End If
  vsu$ = ""
  If vsuch.value <> 0 Then
    vsu$ = connwrd$ & "instr(lcase(mailsafe.volltext),'" + LCase(s$) + "')>0 "
  End If
  c$ = "SELECT mailsafe.id as msgid, mailsafe.frm as sender, " + _
                    "mailsafe.Subject as sbj, mailsafe.Header as hdr, " + _
                    "mailsafe.erstellt as crdtg, mailsafe.owner as own, " + _
                    "mailsafe.Message as msg " + _
                    "FROM mailsafe " + _
                    nurmeine$ & krit$ & vsu$
  If trm(vbd$) <> "" Then
  If trm(nurmeine$ & krit$ & vsu$) = "" Then
    c$ = c$ & " where " & vbd$
  Else
    c$ = c$ & " and (" & vbd$ & ") "
  End If
  End If
  c$ = c$ & " ORDER BY mailsafe.erstellt DESC;"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
While Not r.EOF And tm_brk% = 0 And fnd% < shm
  fnd% = fnd% + 1
  Set lvitem = gd1.ListItems.add(, , r!sender & Space$(80) & "(ID:" & r!msgid)
  lvitem.SubItems(1) = trm(r!crdtg)
  lvitem.SubItems(2) = r!sbj
  If Not IsNull(r!hdr) And r!hdr > 0 Then
    On Error Resume Next
    lvitem.SubItems(3) = hdr_status$(r!hdr)
    On Error GoTo 0
  End If
  lvitem.SubItems(4) = trm(r!own)
  lvitem.SubItems(5) = trm(r!msg)
  DoEvents
  r.MoveNext
Wend
getroffn.Caption = "Treffer: " & trm(fnd%)
merunning = 0
msafe.MousePointer = 0
End Sub

Public Sub rlist1bypath(pfad As String)
Dim rtmp As ADODB.Recordset, fnd%, pth$
Dim c$, lvitem, tr, ffn$, rrr

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "rlist1bypath"
gd1.ListItems.Clear
pth$ = pfad: If Right$(pth$, 1) <> "\" Then pth$ = pth$ + "\"
tr = Dir(pth$ + "\*.msg")
While tr <> ""
  ffn$ = pth$ + tr
  ffn$ = strrepl(ffn$, "\", "|")
  ffn$ = strrepl(ffn$, "|", "\\")
  c$ = "SELECT mailsafe.id AS msgid, mailsafe.frm AS sender, mailsafe.Subject AS sbj, mailsafe.Header AS hdr, mailsafe.erstellt AS crdtg, mailsafe.owner AS own, mailsafe.Message AS msg from mailsafe where message='" + ffn$ + "';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    Set lvitem = gd1.ListItems.add(, , rtmp!sender & Space$(80) & "(ID:" & rtmp!msgid)
    lvitem.SubItems(1) = trm(rtmp!crdtg)
    lvitem.SubItems(2) = rtmp!sbj
    If Not IsNull(rtmp!hdr) And rtmp!hdr > 0 Then lvitem.SubItems(3) = hdr_status$(rtmp!hdr)
    lvitem.SubItems(4) = trm(rtmp!own)
    DoEvents
  Else
    Set lvitem = gd1.ListItems.add(, , pth$ + tr)
    lvitem.SubItems(2) = "ohne Eintrag im Mailsafe"
    lvitem.SubItems(4) = tr
    DoEvents
  End If
  On Error Resume Next
  tr = Dir
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then tr = ""
Wend

End Sub



Private Sub ftrg_Click()
'd2infile = "msafe": d2insub = "ftrg_Click"
Command6.Enabled = True
End Sub

Private Sub gd1_Click()
Dim frm$, p%, rrr, i%, r As ADODB.Recordset, cmd$, id$, fn$, xx$, f0$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "gd1_Click"
On Error Resume Next
frm$ = gd1.SelectedItem
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

p% = 0
For i% = 1 To gd1.ListItems.Count
  If gd1.ListItems(i%).Selected = True Then
    p% = p% + 1
    If p% > 1 Then Exit For
  End If
Next i%
If InStr(frm$, "(ID:") > 0 Then frm$ = trm(Left(frm$, InStr(frm$, "(ID:") - 1))
p% = InStr(frm$, "<")
If p% > 0 Then
  frm$ = Mid$(frm$, p% + 1)
  If InStr(frm$, ">") > 1 Then frm$ = Left$(frm$, InStr(frm$, ">") - 1)
Else
  p% = InStr(frm$, "(")
  If p% > 0 Then
    frm$ = trm(Left$(frm$, p% - 1))
  End If
End If
form1.Combo1.text = frm$

If form1.getusersetting("mailviewer") = "ja" Then
If sdet.value = 1 Then
MousePointer = vbHourglass: DoEvents
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:")
If p% > 0 Then
  id$ = Mid$(id$, p% + 4)
  cmd$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    fn$ = form1.dupcheck(r!message)
    f0$ = form1.composeemlname(r!message): If f0$ <> "" Then fn$ = f0$
    If exist(fn$) = 1 Then
      xx$ = trm(form1.getmyeditor(FileExtension(fn$)))
      If xx$ <> "" Then
        Call form1.openthisdoc(fn$, "")
      Else
        Load mexplore
        On Error Resume Next
        'do not Call mexplore.SetFocus
        mexplore.fnam = fn$
        On Error GoTo 0
      End If
    End If
  End If
End If
End If
MousePointer = 0
End If
End Sub

Private Sub gd1_DblClick()
Dim rrr
Dim r As ADODB.Recordset, cl$, id$, p%, cmd$, fn$, fno$, mlcl$, o%, from$, l$, X, mlclf$, xx$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "gd1_DblClick"
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:")
If p% = 0 Then
  X = Shell("notepad.exe " + id$, 1)
  Exit Sub
End If
MousePointer = 11: DoEvents
id$ = Mid$(id$, p% + 4)
cmd$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  fn$ = form1.dupcheck(trm(r!message))
  If exist(fn$) = 1 Then
    cl$ = form1.getusersetting("mailclient")
    If InStr(LCase(cl$), "netscape") > 0 Or LCase(form1.getusersetting("Mozillaclient")) = "ja" Then cl$ = "NETSCAPE47"
    fno$ = FileName(fn$)
    If cl$ = "NETSCAPE47" Then
      mlcl$ = strrepl(form1.getusersetting("netscape47inbox"), """", "")
      If exist(mlcl$) = 0 Then
        MousePointer = 0: DoEvents
        Call form1.openthisdoc(fn$, "")
      Else
        o% = FreeFile
        Open fn$ For Input As #o%
        p% = FreeFile
        Open mlcl$ For Append As #p%
        Print #p%, "From " & from$
        While Not EOF(o%)
          Line Input #o%, l$
          If Left(LCase(l$), 9) <> "x-mozilla" Then Print #p%, l$
        Wend
        Close #o%
        Close #p%
        mlcl$ = form1.getusersetting("mailclient")
        MousePointer = 0: DoEvents
        If exist(word1(mlcl$)) > 0 Then
          X = Shell(mlcl$, 1)
        Else
          Call form1.openthisdoc(fn$, "")
        End If
      End If
      GoTo dnx
    End If
    'letzte Rettung
    Call form1.openthisdoc(fn$, "noconvert")
dnx:
  Else
    MousePointer = 0: DoEvents
    MsgBox (fn$ + " kann nicht gefunden werden.")
  End If
  MousePointer = 0: DoEvents
End If

End Sub

Private Sub gd1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rrr
Dim idx%, id$, sq$, i As Integer
Dim r As ADODB.Recordset, ask%, p%, fn$

Dim d2infile As String, d2insub As String
d2infile = "msafe": d2insub = "gd1_KeyDown"
'<strg>a
If KeyCode = 65 And pcode = 17 Then
  For i = gd1.ListItems.Count To 1 Step -1
    gd1.ListItems(i).Selected = True
  Next i
End If
If KeyCode = 8 Or KeyCode = 46 Then
  ask% = MsgBox(transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Historyeintrag löschen?"))
  If ask% = vbYes Then
    For i = gd1.ListItems.Count To 1 Step -1
      If (gd1.ListItems(i).Selected = True) Then
        id$ = gd1.ListItems(i)
        p% = InStr(id$, "(ID:")
        If p% = 0 Then
          On Error Resume Next
          Kill id$
          On Error GoTo 0
          Exit Sub
        End If
        id$ = Mid$(id$, p% + 4)
        If id$ <> "" Then
          sq$ = "select message from mailsafe where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, sq$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
          If Not r.EOF Then
            fn$ = form1.dupcheck(r!message)
            If exist(fn$) = 1 Then Kill fn$
          End If
          sq$ = "delete from dochist where docname='" + form1.dupcheck(fn$) + "'"
          Call form1.sqlqry(sq$)
          sq$ = "delete from mailsafe where id='" + id$ + "'"
          Call form1.sqlqry(sq$)
        End If
      End If
    Next i
    Call rlist1("")
  End If
End If
pcode = KeyCode

End Sub

Private Sub monly_Click()
'd2infile = "msafe": d2insub = "monly_Click"
mineonly% = monly.value
Call rlist1("")

End Sub

Private Sub sdet_Click()
'd2infile = "msafe": d2insub = "sdet_Click"
Call form1.setmylastFormVar(Me.name, "sdet", trm(sdet.value))

End Sub

Private Sub shmax_Change()
'd2infile = "msafe": d2insub = "shmax_Change"
Call timerreset

End Sub

Private Sub suchw_Change()
Dim r$
'd2infile = "msafe": d2insub = "suchw_Change"
r$ = strrepl(suchw.text, "[", "")
suchw.text = r$
Call timerreset
End Sub

Sub timerreset()
'd2infile = "msafe": d2insub = "timerreset"
Timer1.Enabled = False
tm_brk% = 1
Me.Caption = "Mailsafe"
DoEvents
Timer1.Interval = form1.getsuchvz()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
'd2infile = "msafe": d2insub = "Timer1_Timer"
Call form1.dbg2f("msafe Timer1 start")
Timer1.Enabled = False
tm_brk% = 0
DoEvents
Call rlist1("")
tm_brk% = 0
Call form1.dbg2f("msafe Timer1 exit")
End Sub

Private Sub Timer2_Timer()
Dim frm$, rrr

'd2infile = "msafe": d2insub = "Timer2_Timer"
Call form1.dbg2f("fdet Timer2 start")
On Error Resume Next
frm$ = gd1.SelectedItem
rrr = Err
On Error GoTo 0
If gd1.ListItems.Count > 0 And rrr = 0 Then
  If Command2.Enabled = False Then Command2.Enabled = True
Else
  If Command2.Enabled = True Then Command2.Enabled = False
End If
Call form1.dbg2f("fdet Timer2 exit")
End Sub

Private Sub vdat_Change()
'd2infile = "msafe": d2insub = "vdat_Change"
Call timerreset

End Sub

Private Sub vdat_DblClick()

'd2infile = "msafe": d2insub = "vdat_DblClick"
  With frmCalendar
    .init vdat, vdat.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      vdat.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar

End Sub

Private Sub vsuch_Click()
'd2infile = "msafe": d2insub = "vsuch_Click"
Call form1.setmylastFormVar(Me.name, "vsuch", trm(vsuch.value))
Call rlist1("")

End Sub

