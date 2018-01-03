VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "resizer.ocx"
Begin VB.Form verwalt_sbf 
   Caption         =   "Serienbriefe"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.ListBox List4 
      Height          =   1530
      IntegralHeight  =   0   'False
      Left            =   7080
      MultiSelect     =   2  'Erweitert
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   840
      Picture         =   "verwalt_sbf.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Formular schiessen"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7080
      TabIndex        =   13
      Text            =   "5"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3720
      Picture         =   "verwalt_sbf.frx":0702
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Adresse als Dokument"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PLZ, Ort, Strasse"
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   1530
      IntegralHeight  =   0   'False
      Left            =   7080
      MultiSelect     =   2  'Erweitert
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   2010
      IntegralHeight  =   0   'False
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Alle zeigen"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2040
      IntegralHeight  =   0   'False
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   7920
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   360
      Picture         =   "verwalt_sbf.frx":0F24
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   2760
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"verwalt_sbf.frx":1174
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Land beachten"
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dokumente öffnen"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   9360
      Picture         =   "verwalt_sbf.frx":11F6
      ToolTipText     =   "Markierte Dokumente drucken"
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   20
      Left            =   6120
      Picture         =   "verwalt_sbf.frx":1380
      ToolTipText     =   "Markiertes Dokument drucken"
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drucker:"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   300
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3135
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3135
      Left            =   6840
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "verwalt_sbf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prset As Boolean, wv$, l3cnocnt As Boolean

Private Function placedoc(mde$) As Boolean
Dim sid$, sida$, sidk$, c$, i%, p%, j%, rtmp As ADODB.Recordset, xx$, rrr
Dim wt%, wv$, dd$, wbf As Integer
Dim Result As Long, Required As Long, BufLen As Long
Dim Buffer() As Long, Entries As Long
Dim hPrinter As Long, l As Long, X As Long
Dim LiMem As Integer
Dim pname As String, aa As String, printer As printer
Dim d2infile As String, d2insub As String

d2infile = "verwalt_sbf": d2insub = "placedoc"
i% = List2.ListIndex
If i% < 0 Then Exit Function
j% = List1.ListIndex
If j% < 0 Then Exit Function
c$ = word1(List1.List(j%))
placedoc = True
BufLen = 4
ReDim Buffer(0)
wv$ = form1.getusersetting("wordviewer", "")
sid$ = List2.List(i%)
p% = InStr(sid$, "{")
sida$ = sid$: sidk$ = "-1"
If p% > 0 Then
    sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    sidk$ = form1.getkontaktidbyname(sida$, trm(Left(sid$, p% - 1)))
End If
dd$ = "delete from dochist where doctyp='serienbrief_" + c$ + "' and adresse='" + sida$ + "' and kontakt='" + sidk$ + "';"
c$ = "select docname from dochist where doctyp='serienbrief_" + c$ + "' and adresse='" + sida$ + "' and kontakt='" + sidk$ + "';"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Or mde$ = "delete" Then
  c$ = trm(rtmp!docname)
  If Not nexist(c$) Or mde$ = "delete" Then
    Select Case mde$
      Case "open": On Error Resume Next
                   RichTextBox1.LoadFile (c$)
                   rrr = Err
                   On Error GoTo 0
                   If rrr <> 0 Then
                     MsgBox "das Dokument kann nicht angezeigt werden." + vbCrLf + "Möglicherweise ist es bereits geöffnet.?"
                   End If
      Case "wwclose":
                   MousePointer = 11: DoEvents
                   On Error Resume Next
                   AppActivate FileName(c$)
                   rrr = Err
                   On Error GoTo 0
                   DoEvents
                   If rrr = 0 Then
                     form1.SendKys "%{F4}", 1
                     DoEvents
                   End If
                   MousePointer = 0
                   DoEvents
      Case "print":
                   wt% = Int(FileSize(c$) / 1000000)
                   wbf = 0
                   On Error Resume Next
                   wbf = Val(form1.getusersetting("serienbriefwartenvordruck", "1"))
                   wt% = Val(form1.getusersetting("serienbriefwartennachdruck", trm(wt%)))
                   rrr = Err
                   On Error GoTo 0
                   If rrr <> 0 Then wbf = 0
                   If wbf < 0 Then wbf = 0
                   If wt% > 5 Then wt% = 5
                   If wt% < 2 Then wt% = 2
                   MousePointer = 11: DoEvents
                   If wv$ <> "" And (Not nexist(wv$)) Then
                     X = Shell(wv$ + " " + c$, 1)
                     DoEvents
                     Call wait(wbf)
                     Call form1.SendKys("^p{Enter}", 1)
                   End If
                   Call wait(wt%)
                   MousePointer = 0
                   DoEvents
      Case "getname":
        Label2.Caption = c$
      Case "delete":
        On Error Resume Next
        Kill trm(rtmp!docname)
        On Error GoTo 0
        Call form1.sqlqry(dd$)
        List2.RemoveItem i%
      Case Else
    End Select
  Else
    placedoc = False
  End If
End If
End Function

Private Sub Check1_Click()
Call form1.setmylastFormVar(Me.name, "useland", trm(Check1.value))

End Sub

Private Sub Combo1_Click()
Dim X As Boolean
If Not prset Then Exit Sub
X = SetPrinter(Combo1.text)
Call rdrucker
End Sub

Private Sub Combo2_Click()
Call rlist1
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Combo2.text = "": Call rlist1
End Sub

Private Sub Command3_Click()
Dim o%, p%, nam$, betr$, land$, plz$, ort$, plzort$
Dim rtmp As ADODB.Recordset, kabt$, pfadr As Boolean, pferg$, adressname$
Dim i%, sid$, sida$, sidk$, stra$, knam$, kid$, tz$, adl As String

pfadr = False
pferg$ = form1.getusersetting("postfachergänzen", "")
List3.Clear

For i% = 0 To List2.ListCount - 1
  sid$ = List2.List(i%)
  p% = InStr(sid$, "{")
  sida$ = sid$: sidk$ = ""
  If p% > 0 Then
    sidk$ = trm(Left(sid$, p% - 1))
    sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rtmp.Open "SELECT * FROM adresse where id ='" + sida$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
  land$ = "": plz$ = "": ort$ = "": plzort = "": nam$ = ""
  If Not rtmp.EOF Then
    stra$ = trm(rtmp!strasse)
    If shwAdrDetail.Check3.value = 1 And trm(rtmp!postfach) <> "" And trm(rtmp!plzpostfach) <> "" Then
      stra$ = trm(rtmp!postfach)
      pfadr = True
      If pferg$ <> "" Then
        If InStr(LCase(stra$), pferg$) = 0 Then
          stra$ = pferg$ & " " & stra$
        End If
      End If
    End If
    If Not IsNull(rtmp!land) Then land$ = rtmp!land
    If LCase(land$) = LCase(form1.getusersetting("meinland")) Then land = ""
    If Not IsNull(rtmp!plz) Then plz$ = rtmp!plz
    If pfadr And trm(rtmp!plzpostfach) <> "" Then plz$ = trm(rtmp!plzpostfach)
    If Not IsNull(rtmp!ort) Then
      ort$ = rtmp!ort
    End If
  End If
  
  If sidk$ <> "" Then
    kid$ = form1.getkontaktidbyname(sida$, sidk$)
    If kid$ <> "-1" Then
      Set rtmp = New ADODB.Recordset
      rtmp.CursorLocation = adUseServer
      rtmp.Open "SELECT * FROM kontakt where id='" + kid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
      If Not rtmp.EOF Then
        If trm(rtmp!strasse) <> "" Then stra$ = rtmp!strasse
        If shwAdrDetail.Check3.value = 1 And trm(rtmp!postfach) <> "" And trm(rtmp!plzpostfach) <> "" Then
          stra$ = trm(rtmp!postfach)
          pfadr = True
          If pferg$ <> "" Then
            If InStr(LCase(stra$), pferg$) = 0 Then
              stra$ = pferg$ & " " & stra$
            End If
          End If
        End If
        If trm(rtmp!lkz) <> "" Then land$ = rtmp!lkz
        If LCase(land$) = LCase(form1.getusersetting("meinland")) Then land = ""
        If trm(rtmp!plz) <> "" Then plz$ = rtmp!plz
        If pfadr And trm(rtmp!plzpostfach) <> "" Then plz$ = trm(rtmp!plzpostfach)
        If trm(rtmp!ort) <> "" Then ort$ = rtmp!ort
      End If
    End If
  End If

  adl = plz$ + " " + ort$ + " " + stra$ + "|" + List2.List(i%)
  If Check1.value = 1 Then adl = land$ + " " + adl
  List3.AddItem adl
  DoEvents
Next i%
Label3.Caption = trm(List3.ListCount)
End Sub

Private Sub Command32_Click()
If Not nexist(trm(Label2.Caption)) Then Call form1.openthisdoc(Label2.Caption, "")
End Sub

Private Sub Command4_Click()
Unload form1

End Sub

Private Sub Image1_Click(Index As Integer)
Dim i%, j%, slct As Boolean, rrr, mxop%, k As Integer

If Index = 20 Then
  Call placedoc("print")
  Exit Sub
End If

If Index = 0 Then
  l3cnocnt = True
  j% = 0
  On Error Resume Next
  mxop% = Val(Text1.text)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then mxop% = 5
  If mxop% < 1 Then mxop% = 1
  slct = False
  For i% = 0 To List3.ListCount - 1
    If List3.Selected(i%) Then
      slct = True
      Exit For
    End If
  Next i%
  List4.Clear
  For i% = 0 To List3.ListCount - 1
    If Not slct Or (slct And List3.Selected(i%)) Then List4.AddItem List3.List(i%)
  Next i%
  List4.Visible = True
  For i% = 0 To List4.ListCount - 1
    List3.ListIndex = -1
    For k = 0 To List3.ListCount - 1
      If List3.List(k) = List4.List(i%) Then
        List3.ListIndex = k
        Exit For
      End If
    Next k
    If List3.ListIndex >= 0 Then
      DoEvents
      Call List3_Click
      DoEvents
      Call Image1_Click(20)
    End If
    If ((i% + 1) Mod mxop%) = 0 Then
      For j% = Max(0, i% - mxop%) To i%
        List4.ListIndex = j%
        List3.ListIndex = -1
        For k = 0 To List3.ListCount - 1
          If List3.List(k) = List4.List(j%) Then
            List3.ListIndex = k
            Exit For
          End If
        Next k
        If List3.ListIndex >= 0 Then
          DoEvents
          Call placedoc("wwclose")
          DoEvents
        End If
      Next j%
      DoEvents
    End If
  Next i%
  While j% < i%
    List4.ListIndex = j%
    List3.ListIndex = -1
    For k = 0 To List3.ListCount - 1
      If List3.List(k) = List4.List(j%) Then
        List3.ListIndex = k
        Exit For
      End If
    Next k
    If List3.ListIndex >= 0 Then
      DoEvents
      Call placedoc("wwclose")
      DoEvents
    End If
    j% = j% + 1
  Wend
  List4.Visible = False
  List4.Clear
  l3cnocnt = False
  Exit Sub
End If

End Sub

Private Sub Form_Load()
Dim k1lrv%

prset = True
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Me.Caption = transe("Serienbriefverwaltung")
Command1.ToolTipText = transe("Formular schliessen")
Label5.Caption = transe("Land beachten")
Label1.Caption = transe("Drucker:")
l3cnocnt = False
List4.Visible = False
k1lrv% = Val(form1.mylastFormVar(Me.name, "useland", "0"))
If k1lrv% <> 1 Then k1lrv% = 0
Check1.value = k1lrv%

Show
Call rdrucker
Call rcombo2
Call rlist1
Image1(0).Enabled = False
Image1(20).Enabled = False
wv$ = form1.getusersetting("wordviewer", "")
If wv$ <> "" Then
  If Not nexist(wv$) Then
    Image1(0).Enabled = True
    Image1(20).Enabled = True
  End If
End If
BackColor = form1.cleancolor()

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


Sub rdrucker()
Dim i As Integer
    Dim PrinterName As String
    
Combo1.Clear
    For i = 0 To Printers.Count - 1
        PrinterName = Printers(i).DeviceName
        'PrinterName = PrinterName & " (" & Printers(i).port & ")"
        Combo1.AddItem PrinterName
      
        If printer.DeviceName = Printers(i).DeviceName Then
            'Combo1.Text = PrinterName
            prset = False
            Combo1.ListIndex = i
            prset = True
        End If
    Next i
End Sub
Private Function SetPrinter(ByVal prnName As String) _
  As Boolean

  Dim Result As Boolean
  Dim X As Integer

  Result = False
  If Printers.Count > 0 Then
    For X = 0 To Printers.Count - 1
      If Printers(X).DeviceName = prnName Then
        Set printer = Printers(X)
        Result = True
        Exit For
      End If
    Next X
  End If
  SetPrinter = Result
End Function

Sub rlist1()
Dim c$, r As ADODB.Recordset, rrr
Dim pb$, pt$, cnt%, usr$, pbx$
Dim d2infile As String, d2insub As String

d2infile = "verwalt_sbf": d2insub = "rlist1"

List1.Clear
MousePointer = 11: DoEvents
pb$ = ""
usr$ = trm(Combo2.text)
If usr$ <> "" Then usr$ = "owner='" + usr$ + "' and "
c$ = "select * from dochist where " + usr$ + " instr(doctyp,'serienbrief_')=1 order by Betreff;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
cnt% = 0
While Not r.EOF
  pt$ = trm(r!doctyp)
  If pb$ <> pt$ Then
    If cnt% > 0 Then
      List1.AddItem pbx$ + " " + transe("Anz.") + " " + trm(cnt%)
      DoEvents
    End If
    cnt% = 0
    pb$ = pt$
    pbx$ = Mid$(pb$, 13)
  End If
  cnt% = cnt% + 1
  r.MoveNext
Wend
If cnt% > 1 Then List1.AddItem pbx$ + " " + transe("Anz.") + " " + trm(cnt%)
MousePointer = 0
End Sub

Sub rcombo2()
Dim rtmp As ADODB.Recordset

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly
Combo2.Clear
While Not rtmp.EOF
  Combo2.AddItem rtmp!id
  rtmp.MoveNext
Wend
Combo2.text = form1.getuserid()

End Sub



Private Sub Label5_Click()
If Check1.value = 0 Then
  Check1.value = 1
Else
  Check1.value = 0
End If
End Sub

Private Sub List1_Click()
Dim i%

List2.Clear: List3.Clear
i% = List1.ListIndex
If i% < 0 Then Exit Sub
Call rlist2(word1(List1.List(i%)))

End Sub

Sub rlist2(dtyp$)
Dim c$, r As ADODB.Recordset, kkid$, rrr

List2.Clear: List3.Clear
MousePointer = 11: DoEvents
c$ = "select * from dochist where doctyp='serienbrief_" + dtyp$ + "';"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
While Not r.EOF
  If trm(r!kontakt) <> "-1" Then
    kkid$ = form1.get_kontaktname_by_id(r!kontakt) & "{" & trm(r!adresse) + "}"
  Else
    kkid$ = trm(r!adresse)
  End If
  List2.AddItem kkid$
  r.MoveNext
Wend
MousePointer = 0
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X, i%, ask, dtyp$

If KeyCode = 8 Or KeyCode = 46 Then
  ask = MsgBox(transe("Dokumente und Einträge der Kontakthistorie löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Serienbriefe löschen?"))
  If ask <> vbYes Then Exit Sub
  i% = List1.ListIndex
  If i% < 0 Then Exit Sub
  dtyp$ = word1(List1.List(i%))
  While List2.ListCount > 0
    List2.ListIndex = 0
    DoEvents
    X = placedoc("delete")
  Wend
  List3.Clear
  Call form1.sqlqry("delete from dochist where doctyp='serienbrief_" + dtyp$ + "';")
  Call rlist1
End If
End Sub

Private Sub List2_Click()
If placedoc("getname") Then
  Call placedoc("open")
  Command32.Enabled = True
Else
  Command32.Enabled = False
End If
End Sub

Private Sub List2_DblClick()
Dim sid$, p%, sida$, sidk$, i%

i% = List2.ListIndex
If i% < 0 Then Exit Sub

  sid$ = List2.List(i%)
  p% = InStr(sid$, "{")
  sida$ = sid$: sidk$ = ""
  If p% > 0 Then
    sidk$ = trm(Left(sid$, p% - 1))
    sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If Len(sida$) > 0 Then
    Load shwAdrDetail
    Call shwAdrDetail.refreshadrdetail(sida$, sidk$)
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
  End If

End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X

If KeyCode = 8 Or KeyCode = 46 Then X = placedoc("delete")
End Sub

Private Sub List3_Click()
Dim sid$, p%, sida$, sidk$, i%, cnt As Integer

List2.ListIndex = -1
i% = List3.ListIndex
If i% < 0 Then Exit Sub

sid$ = List3.List(i%)
p% = InStr(sid$, "|")
sid$ = Mid$(sid$, p% + 1)
For i% = 0 To List2.ListCount - 1
    If List2.List(i%) = sid$ Then
      List2.ListIndex = i%
      Exit For
    End If
Next i%
If Not l3cnocnt Then
cnt = 0
For i% = 0 To List3.ListCount - 1
  If List3.Selected(i%) Then cnt = cnt + 1
Next i%
If cnt = 0 Then
  Label3.Caption = trm(List3.ListCount)
Else
  Label3.Caption = trm(cnt)
End If
End If
End Sub

Private Sub List3_DblClick()
Dim sid$, p%, sida$, sidk$, i%

i% = List3.ListIndex
If i% < 0 Then Exit Sub
Call List2_DblClick

End Sub
