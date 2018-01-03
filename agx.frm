VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form agx 
   Caption         =   "Datenimport aus Agencyprof"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   ScaleHeight     =   5085
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "alle auswählen"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "gewählte importieren"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "agx.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Formular schiessen"
      Top             =   4680
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   1875
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   240
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "agx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fndi%, posv As Double, posb As Double, nosho As Boolean, vcfmode As Boolean

Private Sub Command1_Click()
d2infile = "agx": d2insub = "Command1_Click"
Unload Me
End Sub

Private Function vcfstarttest(lne$) As Boolean
Dim i%, z$

vcfstarttest = False
i% = 1
Do
  z$ = Mid$(lne$, i%, 1)
  If z$ < "A" Or z$ > "Z" Then
    If z$ = ":" Or z$ = ";" Then vcfstarttest = True
    Exit Function
  End If
 i% = i% + 1
Loop Until i% > Len(lne$)

End Function

Public Sub Command2_Click()
Dim r As ADODB.Recordset, iskontakt As Boolean, cmd$
Dim o%, l$, sq$, fn$, i%, j%, strzus$, hid$
Dim vcfcrlf As String, nvcfcrlf As Integer, telfld As String
Dim vcfn(99) As String, vcfa(99) As String, vcfw(99) As String
Dim vcfptr As Integer, vcmem As String, vcorg As String, vcnam As String
Dim vcid As String, rrr, kid$, fnk$

vcfcrlf = form1.getusersetting("vcfcrlf", "=0D=0A")
nvcfcrlf = Len(vcfcrlf)
d2infile = "agx": d2insub = "Command2_Click"
i% = 0
While i% < List1.ListCount
If List1.Selected(i%) Then

fn$ = form1.s0dir() & "\" & List1.List(i%)
fnk$ = fn$
If exist(fn$) <> 0 Then

MousePointer = 11
DoEvents
o% = FreeFile
Open fn$ For Input As #o%
List2.Clear
If Not vcfmode Then
  While Not EOF(o%)
    sq$ = ""
    Do
      On Error Resume Next
      Line Input #o%, l$
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        If rrr = 62 Then GoTo errrx
        MsgBox "Fehler Nr." & rrr & " beim Import" & vbCrLf & Error$(rrr)
        End
      End If
      If Left(l$, 2) <> "--" Then
        If Len(sq$) > 0 Then sq$ = sq$ & vbCrLf
        sq$ = sq$ + l$
      End If
    Loop Until Right$(trm(sq$), 1) = ";"
    If trm(sq$) <> ";" Then
      If InStr(LCase$(sq$), "insert into ") = 1 Then
        form1.err_dupok% = 1
      End If
      If List2.ListCount > 1000 Then List2.Clear
      List2.AddItem l$
      List2.ListIndex = List2.ListCount - 1
      DoEvents
      Call form1.sqlqry(sq$)
      err_dupok% = 0
    End If
  Wend
Else
' vcf
'  Unload shwAdrDetail: DoEvents
'  Load shwAdrDetail: DoEvents
  vcfptr = -1: vcmem = ""
  While Not EOF(o%)
    sq$ = ""
    Do
        On Error Resume Next
        Line Input #o%, l$
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then
          If rrr = 62 Then GoTo errrx
          MsgBox "Fehler Nr." & rrr & " beim Import" & vbCrLf & Error$(rrr)
          End
        End If
        sq$ = sq$ + l$
        If Right$(sq$, 1) = "=" Then sq$ = Left(sq$, Len(sq$) - 1)
    Loop Until EOF(o%) Or Right$(l$, 1) <> "="
    If Left$(sq$, 1) <> " " And InStr(sq$, "X-") <> 1 And sq$ <> "" Then
      sq$ = strrepl(sq$, vcfcrlf, vbCrLf)
      vcfptr = vcfptr + 1
      vcfn(vcfptr) = cut_d1(sq$, ":")
      vcfa(vcfptr) = cut_d2bis(vcfn(vcfptr), ";")
      vcfn(vcfptr) = cut_d1(vcfn(vcfptr), ";")
      vcfw(vcfptr) = cut_d2bis(sq$, ":")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=3D", "=")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=E4", "ä")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=F6", "ö")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=FC", "ü")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=DF", "ß")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=C4", "Ä")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=D6", "Ö")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "=DC", "Ü")
      vcfw(vcfptr) = strrepl(vcfw(vcfptr), "'", "´")
      If LCase(vcfn(vcfptr)) = "fn" Then vcnam = vcfw(vcfptr)
      If LCase(vcfn(vcfptr)) = "org" Then vcorg = vcfw(vcfptr)
    End If
  Wend
  vcid = vcorg
  If vcid = "" Then vcid = vcnam
  iskontakt = False
  vcid = strrepl(vcid, "(", " ")
  cmd$ = "select * from adresse where id='" & vcid$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
  If r.EOF Then
    If vcorg <> "" Then vcnam = vcorg + vbCrLf + vcnam
    cmd$ = "insert into adresse (id,name) values('" + vcid + "','" + vcnam + "')"
    Call form1.sqlqry(cmd$)
    kid$ = "-1"
  Else
    iskontakt = True
    kid$ = form1.newid("kontakt", "id", 10)
    cmd$ = "insert into kontakt (id) values('" + kid$ + "')"
    Call form1.sqlqry(cmd$)
    cmd$ = "update kontakt set vid='" + vcid + "' where id ='" + kid$ + "'"
    Call form1.sqlqry(cmd$)
    cmd$ = "update kontakt set vid='" + vcnam + "' where id ='" + kid$ + "'"
    Call form1.sqlqry(cmd$)
  End If
  Call form1.neuart(vcid$, kid$, "Person", "")
  For j% = 0 To vcfptr
    Select Case LCase(vcfn(j%))
      Case "org":
      Case "fn":
      Case "begin":
      Case "label":
      Case "title":
      Case "version":
      Case "rev":
      Case "end":
      Case "tel":
        telfld = "Tel"
        If InStr(LCase(vcfa(j%)), "cell") > 0 Then telfld = "handy"
        If InStr(LCase(vcfa(j%)), "fax") > 0 Then telfld = "fax"
        l$ = "adresse": If iskontakt Then l$ = "kontakt"
        If Not iskontakt Then
          sq$ = vcid
        Else
          sq$ = kid$
        End If
        cmd$ = "update " + l$ + " set " + telfld + "='" + vcfw(j%) + "' where id='"
        cmd$ = cmd$ + sq$ + "'"
        Call form1.sqlqry(cmd$)
      Case "email":
        l$ = "adresse": If iskontakt Then l$ = "kontakt"
        If Not iskontakt Then
          sq$ = vcid
        Else
          sq$ = kid$
        End If
        cmd$ = "update " + l$ + " set email='" + vcfw(j%) + "' where id='"
        cmd$ = cmd$ + sq$ + "'"
        Call form1.sqlqry(cmd$)
      Case "url":
        l$ = "adresse": If iskontakt Then l$ = "kontakt"
        If Not iskontakt Then
          sq$ = vcid
        Else
          sq$ = kid$
        End If
        cmd$ = "update " + l$ + " set url='" + vcfw(j%) + "' where id='"
        cmd$ = cmd$ + sq$ + "'"
        Call form1.sqlqry(cmd$)
      Case "n":
      Case "note":
        If Not iskontakt Then
          cmd$ = "update adresse set Hinweise='" + vcfw(j%) + "' where id='" + vcid$ + "'"
          Call form1.sqlqry(cmd$)
        Else
          hid$ = vcid$: If kid$ <> "-1" Then hid$ = hid$ + kid$
          Call form1.higruinsert(hid$, "Person", "Hinweise", vcfw(j%))
        End If
      Case "adr":
        If InStr(LCase(vcfa(j%)), "work") > 0 Then
        
        l$ = "adresse": If iskontakt Then l$ = "kontakt"
        If Not iskontakt Then
          sq$ = vcid
        Else
          sq$ = kid$
        End If
        fn$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        If fn$ <> "" Then
          cmd$ = "update " + l$ + " set Postfach='" + fn$ + "' where id='"
          cmd$ = cmd$ + sq$ + "'"
          Call form1.sqlqry(cmd$)
        End If
        strzus$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        fn$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        If strzus$ <> "" Then fn$ = fn$ + vbCrLf + strzus$
        If fn$ <> "" Then
          cmd$ = "update " + l$ + " set strasse='" + fn$ + "' where id='"
          cmd$ = cmd$ + sq$ + "'"
          Call form1.sqlqry(cmd$)
        End If
        fn$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        strzus$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        If strzus$ <> "" Then fn$ = fn$ + " " + strzus$
        If fn$ <> "" Then
          cmd$ = "update " + l$ + " set ort='" + fn$ + "' where id='"
          cmd$ = cmd$ + sq$ + "'"
          Call form1.sqlqry(cmd$)
        End If
        fn$ = trm(cut_d1(vcfw(j%), ";")): vcfw(j%) = cut_d2bis(vcfw(j%), ";")
        If fn$ <> "" Then
          cmd$ = "update " + l$ + " set plz='" + fn$ + "' where id='"
          cmd$ = cmd$ + sq$ + "'"
          Call form1.sqlqry(cmd$)
        End If
        fn$ = vcfw(j%)
        If fn$ <> "" Then
          cmd$ = "update " + l$ + " set "
          If iskontakt Then
            cmd$ = cmd$ + "lkz"
          Else
            cmd$ = cmd$ + "Land"
          End If
          cmd$ = cmd$ + "='" + fn$ + "' where id='" + sq$ + "'"
          Call form1.sqlqry(cmd$)
        End If
        
        End If
      Case Else:
        Debug.Print vcfn(j%); ":-:"; vcfa(j%); ":-:"; vcfw(j%)
        Debug.Print
    End Select
  Next j%
  form1.Combo1.text = strrepl(vcnam, vbCrLf, " ")
  DoEvents
End If

errrx:
Close #o%
MousePointer = 0

On Error Resume Next
Kill fnk$
On Error GoTo 0
List1.RemoveItem i%
i% = i% - 1: If i% < -1 Then i% = -1

End If
End If
i% = i% + 1
Wend


End Sub

Public Sub Command3_Click()
Dim i%
d2infile = "agx": d2insub = "Command3_Click"
nosho = True
  For i% = List1.ListCount - 1 To 0 Step -1
    List1.Selected(i%) = True
  Next i%
nosho = False
End Sub

Private Sub Form_Load()
d2infile = "agx": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
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
vcfmode = False
If form1.outmylanguage(form1.sqlmess.Caption) = "SQL" Then
  agx.Caption = transe("Datenimport aus Agencyprof")
Else
  agx.Caption = transe("Datenimport aus vCards")
  vcfmode = True
End If
Command3.Caption = transe("alle auswählen")
Command2.Caption = transe("gewählte importieren")
Command1.ToolTipText = transe("Formular schliessen")
nosho = False
Show
Call rlist1
BackColor = form1.cleancolor()
If List1.ListCount < 0 Then Exit Sub

For i% = List1.ListCount - 1 To 0 Step -1
  If InStr(List1.List(i%), "autoimport_") = 1 Then
    List1.Selected(i%) = True: DoEvents
  End If
Next i%
DoEvents
Call Command2_Click

End Sub

Private Sub Form_Resize()
d2infile = "agx": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "agx": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Sub rlist1()
Dim ext As String

d2infile = "agx": d2insub = "rlist1"
List1.Clear
List2.Clear
ext = "sql": If vcfmode Then ext = "vcf"
tr = Dir(form1.s0dir() + "\*." + ext)
While tr <> ""
  List1.AddItem tr
  tr = Dir
Wend

End Sub

Private Sub List1_Click()
Dim o%, i%, k As Integer

d2infile = "agx": d2insub = "List1_Click"
List2.Clear
i% = List1.ListIndex
If i% < 0 Or nosho Then Exit Sub
  
  MousePointer = 11
  DoEvents
  k = 0
  o% = FreeFile
  Open form1.s0dir() & "\" & List1.List(i%) For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    List2.AddItem l$
    k = k + 1
    If k > 1000 Then
      k = 0
      List2.ListIndex = List2.ListCount - 1
      DoEvents
    End If
  Wend
  Close #o%
  If List2.ListCount > 0 Then List2.ListIndex = 0
  MousePointer = 0
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim fn$, i%

d2infile = "agx": d2insub = "List1_KeyDown"
'<strg>a
If KeyCode = 65 And pcode = 17 Then Call Command3_Click

If KeyCode = 8 Or KeyCode = 46 Then
  For i% = List1.ListCount - 1 To 0 Step -1
    If List1.Selected(i%) = True Then
      fn$ = form1.s0dir() & "\" & List1.List(i%)
      On Error Resume Next
      Kill fn$
      On Error GoTo 0
      List1.RemoveItem i%
      DoEvents
    End If
  Next i%
  Call rlist1
End If

End Sub
