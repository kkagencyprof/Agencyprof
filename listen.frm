VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "resizer.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form listen 
   Caption         =   "Listenverwaltung"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1680
      Picture         =   "listen.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "als CSV-Datei speichern"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton svdat 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00000000&
      Picture         =   "listen.frx":04B3
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Speichern"
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton mveup 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6480
      TabIndex        =   9
      ToolTipText     =   "Markierten Kontakt nach unten verschieben"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton mvedwn 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5880
      TabIndex        =   8
      ToolTipText     =   "Markierten Kontakt nach oben verschieben"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1080
      Picture         =   "listen.frx":0B25
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Spalten bereinigen"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command25 
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
      Left            =   600
      Picture         =   "listen.frx":168B
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Neue Liste"
      Top             =   4680
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   7080
      Picture         =   "listen.frx":1A1D
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Liste der aufgerufenen Projekte löschen. (Löscht NICHT das Projekt)"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "listen.frx":2CF3
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   4680
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
      _Version        =   393216
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "listen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currx As Integer, curry As Integer

Private Sub combo1_Change()
Dim i%
Dim r As ADODB.Recordset, c$, d2infile As String, d2insub As String
Dim l$, ccol As Integer

d2infile = "listen": d2insub = "Combo1_Change"
fg2.Clear
fg2.Cols = 2
fg2.Rows = 2
l$ = trm(Combo1.text)
c$ = "select * from opt_listen where liste='" + l$ + "' and id='99lid:" + l$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  c$ = "select * from opt_listen where instr(id,'98lid:" + l$ + "|')=1 order by ipos"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    c$ = cut_d2bis(r!id, "|")
    ccol = Val(c$)
    fg2.Cols = Max(fg2.Cols, ccol + 2)
    fg2.ColWidth(ccol) = 1000
    fg2.TextMatrix(0, ccol) = r!vid
    r.MoveNext
  Wend
  c$ = "select * from opt_listen where instr(id,'97lid:" + l$ + "|')=1 order by ipos"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    c$ = cut_d2bis(r!id, "|")
    c$ = cut_d1(c$, "|")
    ccol = Val(c$)
    fg2.Rows = Max(r!iPos + 2, fg2.Rows)
    fg2.Cols = Max(ccol + 2, fg2.Cols)
    fg2.TextMatrix(r!iPos, ccol) = r!vid
    DoEvents
    r.MoveNext
  Wend
End If
Me.BackColor = form1.cleancolor
svdat.Visible = False
mveup.Visible = False
mvedwn.Visible = False

End Sub

Private Sub Combo1_Click()
Dim i%

d2infile = "listen": d2insub = "Combo1_Click"
i% = Combo1.ListIndex
If i% < 0 Then Exit Sub
Call combo1_Change
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim X%, Y%, i%, spl%, von As Integer, bis As Integer
Dim mx%, my%

von = X%
bis = X%
If X% = 0 Then
  von = 1
  bis = fg2.Cols - 1
End If
For spl% = von To bis
  For i% = 1 To fg2.Rows - 1
    If fg2.TextMatrix(i%, spl%) = "" Then
      For j% = i% + 1 To fg2.Rows - 1
        If fg2.TextMatrix(j%, spl%) <> "" Then
          fg2.TextMatrix(i%, spl%) = fg2.TextMatrix(j%, spl%)
          fg2.TextMatrix(j%, spl%) = ""
          Exit For
        End If
      Next j%
    End If
  Next i%
Next spl%
Call svcontent
End Sub

Sub svcontent()
Dim tbl$, X As Integer, Y As Integer

tbl$ = trm(Combo1.text)
If tbl$ = "" Then Exit Sub

c$ = "delete from opt_listen where instr(id,'97lid:" + tbl$ + "|')=1"
Call form1.sqlqry(c$)
For X = 1 To fg2.Cols - 1
  For Y = 1 To fg2.Rows - 1
    If fg2.TextMatrix(Y, X) <> "" Then
      c$ = "insert into opt_listen (id,liste,vid,ipos) values('97lid:" + trm(tbl$) + "|" + trm(X) + "|" + trm(Y) + "','" + tbl$ + "','" + fg2.TextMatrix(Y, X) + "'," + trm(Y) + ")"
      Call form1.sqlqry(c$)
    End If
  Next Y
Next X
Call combo1_Change

End Sub
Private Sub Command25_Click()
Dim betr$, c$, i%

betr$ = trm(InputBox(transe("Name der neuen Liste:"), transe("Neue Liste erstellen"), betr$, 100, 100))
If Len(betr$) = 0 Then Exit Sub
c$ = "insert into opt_listen (id,liste) values('99lid:" + betr$ + "','" + betr$ + "')"
Call form1.sqlqry(c$)
Call listeninit

'beware: exit sub ahead

For i% = 0 To Combo1.ListCount - 1
  If Combo1.List(i%) = betr$ Then
    Combo1.ListIndex = i%
    Exit Sub
  End If
Next i%
End Sub

Private Sub Command3_Click()
Dim xld$, o%, fn$, X%, Y%, c$

xld$ = form1.getusersetting("exceldelimiter", ",")
fn$ = form1.myuniquedocname("", "csv")
If trm(fn$) <> "" Then
  o% = FreeFile
  Open fn$ For Output As #o%
  MousePointer = 11: DoEvents
  For Y% = 0 To fg2.Rows - 1
    c$ = ""
    For X% = 1 To fg2.Cols - 1
      If Len(c$) > 0 Then c$ = c$ + xld$
      c$ = c$ + """" + fg2.TextMatrix(Y%, X%) + """"
    Next X%
    Print #o%, c$
  Next Y%
  Close #o%
  MousePointer = 0: DoEvents
  X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)
End If
End Sub

Private Sub delme_Click()
Dim tbl$

tbl$ = trm(Combo1.text)
If tbl$ = "" Then Exit Sub
c$ = "delete from opt_listen where liste='" + tbl$ + "'"
Call form1.sqlqry(c$)
Call listeninit

End Sub

Private Sub fg2_DblClick()
Dim s0$, neukwert As String, neuwert As String, tbl$

tbl$ = trm(Combo1.text)
If tbl$ = "" Then Exit Sub
Load adrselect
s0$ = fg2.TextMatrix(curry, currx)
If Len(s0$) = 0 Then s0$ = ""
Call adrselect.sel_init(s0$, transe("Person"))
Call adrselect.SetFocus
Do
  DoEvents
Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
If adrselect.sel_brk() = 0 Then
  neukwert = adrselect.get_kontsel()
  neuwert = adrselect.sel_getselected(): neuawert = neuwert
  If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
  c$ = "delete from opt_listen where id='97lid:" + trm(tbl$) + "|" + trm(currx) + "|" + trm(curry) + "'"
  Call form1.sqlqry(c$)
  c$ = "insert into opt_listen (id,liste,vid,ipos) values('97lid:" + trm(tbl$) + "|" + trm(currx) + "|" + trm(curry) + "','" + tbl$ + "','" + neuwert + "'," + trm(curry) + ")"
  Call form1.sqlqry(c$)
  fg2.TextMatrix(curry, currx) = neuwert
End If
Unload adrselect
End Sub

Private Sub fg2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c$, tbl$

If KeyCode = 8 Or KeyCode = 46 Then
  tbl$ = trm(Combo1.text)
  If tbl$ = "" Then Exit Sub
  fg2.TextMatrix(curry, currx) = ""
  c$ = "delete from opt_listen where id='97lid:" + trm(tbl$) + "|" + trm(currx) + "|" + trm(curry) + "'"
  Call form1.sqlqry(c$)
End If
End Sub

Private Sub fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If currx < 0 Or curry < 0 Or currx >= fg2.Cols Or curry >= fg2.Rows Then
  Exit Sub
End If
fg2.col = currx
fg2.Row = curry
fg2.CellBackColor = RGB(0, 0, 0)

End Sub

Private Sub Form_Load()

axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.BackColor = form1.cleancolor
Call listeninit
Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuldl
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuldl:
On Error GoTo 0

End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Sub listeninit()
Dim r As ADODB.Recordset, c$, d2infile As String, d2insub As String

d2infile = "listen": d2insub = "listeninit"
fg2.Clear
fg2.Cols = 2
fg2.Rows = 2
Combo1.Clear
c$ = "select * from opt_listen where instr(id,'99lid:')=1"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  Combo1.AddItem trm(r!liste)
  r.MoveNext
Wend

End Sub

Private Sub fg2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xm%, ym%, rrr, tx%, z$, p%, sida$, sid$, sidk$

mveup.Visible = False
mvedwn.Visible = False
xm% = fg2.MouseCol
ym% = fg2.MouseRow
currx = xm%: curry = ym%
fg2.col = xm%
fg2.Row = ym%
fg2.CellBackColor = RGB(192, 192, 192)
If ym% <> 0 Then
  Text1.Enabled = False
  Text1.Visible = False
  Label1.Visible = True
  Label1.Caption = fg2.TextMatrix(curry, currx)
  sid$ = Label1.Caption
  sida$ = sid$: sidk$ = ""
  p% = InStr(sid$, "{")
  If p% > 0 Then
    sidk$ = trm(Left(sid$, p% - 1))
    sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If Len(sida$) > 0 Then
    mveup.Visible = True
    mvedwn.Visible = True
    Load shwAdrDetail
    Call shwAdrDetail.refreshadrdetail(sida$, sidk$)
  End If
Else
  If xm% > 0 Then
    Text1.Enabled = True
    Text1.Visible = True
    Label1.Visible = False
    Text1.text = fg2.TextMatrix(curry, currx)
    On Error Resume Next
    Call Text1.SetFocus
    On Error GoTo 0
  End If
End If
If xm% = fg2.Cols - 1 Then
  fg2.Cols = fg2.Cols + 1
  fg2.ColWidth(fg2.Cols - 1) = fg2.ColWidth(fg2.Cols - 1) / 2
End If
If ym% = fg2.Rows - 1 Then
  fg2.Rows = fg2.Rows + 1
  For tx% = 0 To fg2.Cols - 1
    fg2.TextMatrix(fg2.Rows - 1, tx%) = fg2.TextMatrix(fg2.Rows - 2, tx%)
    fg2.TextMatrix(fg2.Rows - 2, tx%) = ""
  Next tx%
End If

End Sub

Private Sub mvedwn_Click()
Dim a$

If curry < 2 Then Exit Sub

a$ = fg2.TextMatrix(curry - 1, currx)
fg2.TextMatrix(curry - 1, currx) = fg2.TextMatrix(curry, currx)
fg2.TextMatrix(curry, currx) = a$
fg2.col = currx
fg2.Row = curry
fg2.CellBackColor = RGB(0, 0, 0)
fg2.Row = curry - 1
curry = curry - 1
fg2.Row = curry
fg2.CellBackColor = RGB(192, 192, 192)
svdat.Visible = True
Me.BackColor = form1.dirtycolor()

End Sub

Private Sub mveup_Click()
Dim a$

fg2.Rows = Max(curry + 3, fg2.Rows)
a$ = fg2.TextMatrix(curry + 1, currx)
fg2.TextMatrix(curry + 1, currx) = fg2.TextMatrix(curry, currx)
fg2.TextMatrix(curry, currx) = a$
fg2.col = currx
fg2.Row = curry
fg2.CellBackColor = RGB(0, 0, 0)
fg2.Row = curry + 1
curry = curry + 1
fg2.Row = curry
fg2.CellBackColor = RGB(192, 192, 192)
svdat.Visible = True
Me.BackColor = form1.dirtycolor()

End Sub

Private Sub svdat_Click()

Call svcontent
svdat.Visible = False
Me.BackColor = form1.cleancolor()

End Sub

Private Sub Text1_Change()
Dim c$, tbl$

tbl$ = Combo1.text
If tbl$ = "" Then Exit Sub

If curry > 0 Then Exit Sub
c$ = "delete from opt_listen where id='98lid:" + tbl$ + "|" + trm(currx) + "'"
Call form1.sqlqry(c$)
c$ = "insert into opt_listen (vid,id,liste,ipos) values('" + trm(Text1.text) + "','98lid:" + tbl$ + "|" + trm(currx%) + "','" + trm(tbl$) + "'," + trm(currx) + ")"
Call form1.sqlqry(c$)
If currx > 0 Or curry > 0 Then fg2.TextMatrix(curry, currx) = Text1.text
End Sub
