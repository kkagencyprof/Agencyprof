VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form bplan 
   Caption         =   "Bühnenplan"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox dtg 
      Height          =   285
      Left            =   7080
      TabIndex        =   35
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
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
      Left            =   600
      TabIndex        =   34
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox stbreit 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   32
      Text            =   "200"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox stbreit 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   31
      Text            =   "500"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7200
      TabIndex        =   30
      Text            =   "1"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Streicher"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   3600
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1560
      Picture         =   "bplan.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   28
      ToolTipText     =   "löschen"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox fbreit 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   27
      Text            =   "250"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox fbreit 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   26
      Text            =   "180"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox sbreit 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   25
      Text            =   "100"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox sbreit 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   24
      Text            =   "100"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   23
      Text            =   "100"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   22
      Text            =   "100"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox pbreit 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   21
      Text            =   "100"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox pbreit 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   19
      Text            =   "100"
      Top             =   2760
      Width           =   495
   End
   Begin VB.ComboBox pgid 
      Height          =   315
      Left            =   7080
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox hid 
      Height          =   315
      Left            =   7080
      TabIndex        =   16
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7200
      TabIndex        =   14
      Text            =   "1"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "neu zeichnen"
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bläser"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "gr.Trommel"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Flügel"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "bplan.frx":12D6
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Schlagwerk"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grafik speichern "
      Height          =   255
      Left            =   6240
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Speichert das Bild als Graphik im Medienverzeichnis der Halle"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   960
      MaskColor       =   &H00000000&
      Picture         =   "bplan.frx":1526
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Speichern"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Text            =   "10"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Text            =   "7"
      Top             =   5160
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4740
      ScaleWidth      =   5940
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
      Height          =   255
      Left            =   6240
      TabIndex        =   36
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Programm:"
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Anzahl, Breite,   Tiefe"
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle:"
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Breite:"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiefe:"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   5160
      Width           =   495
   End
End
Attribute VB_Name = "bplan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fndi%
Dim posv As Double, posb As Double
Dim chm%, nupd%, co%, px, py, pco%, msx, msy
Dim obj%(199), cob(1), obptr%, obx(199), oby(199), obb(1, 199), pcob(1), obs(1, 199) As Double
Dim obd$(199)
Dim cobd$, t1w%, osv As Double, osb As Double, drwfast%

Private Sub bbreit_Change(Index As Integer)
'd2infile = "bplan": d2insub = "bbreit_Change"
For j% = 0 To 1: cob(j%) = Val(bbreit(j%).text): Next j%
End Sub

Public Sub Check1_Click(Index As Integer)
'd2infile = "bplan": d2insub = "Check1_Click"
If Index < 0 Then Exit Sub
If nupd% = 0 Then
  nupd% = 1
  aw% = Check1(Index).value
  cobd$ = Check1(Index).Caption
  co% = -1: If aw% = 1 Then co% = Index
  For i% = 0 To chm%
    If i% <> Index Then Check1(i%) = 0
  Next i%
  Select Case co%
    Case 0: For j% = 0 To 1: cob(j%) = Val(sbreit(j%).text): Next j%
    Case 1: cob(0) = Val(fbreit(1).text)
            cob(1) = Val(fbreit(0).text)
    Case 2: For j% = 0 To 1: cob(j%) = Val(pbreit(j%).text): Next j%
    Case 3: For j% = 0 To 1: cob(j%) = Val(bbreit(j%).text): Next j%
    Case 4: For j% = 0 To 1: cob(j%) = Val(stbreit(j%).text): Next j%
    Case Else:
  End Select
End If
nupd% = 0
End Sub

Private Sub Command1_Click()
'd2infile = "bplan": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command19_Click()
'd2infile = "bplan": d2insub = "Command19_Click"
Call form1.handbuchcall("09-Buehnenplaene.htm")

End Sub

Public Sub Command2_Click()
'd2infile = "bplan": d2insub = "Command2_Click"
drwfast% = 0
p1.Cls
DoEvents
List1.ListIndex = -1
For i% = 0 To 199
  If obj%(i%) >= 0 Then
    Call drw(obj%(i%), obx(i%), oby(i%), RGB(0, 0, 0), obb(0, i%), obb(1, i%), obd$(i%), obs(0, i%), obs(1, i%))
    DoEvents
  End If
Next i%
For i% = 0 To chm%
  If Check1(i%).value >= 0 Then
    Check1(i%).value = 0
    Call Check1_Click(i%)
    Exit For
  End If
Next i%
drwfast% = 1
Call drwlegend
End Sub

Private Sub Command4_Click()
Dim h$, p$

'd2infile = "bplan": d2insub = "Command4_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" And p$ = "" Then Exit Sub
p1.Cls
c$ = "delete from bplan where adressid='" & h$ & "' and prgid='" & p$ & "'"
Call form1.sqlqry(c$)
List1.ListIndex = -1
For i% = 0 To 199
  If obj%(i%) >= 0 Then
    Call drw(obj%(i%), obx(i%), oby(i%), RGB(0, 0, 0), obb(0, i%), obb(1, i%), obd$(i%), obs(0, i%), obs(1, i%))
    DoEvents
    c$ = "insert into bplan (id,adressid,prgid,objtyp,objx,objy,objnr,mx,my,obx,oby,obdesc,segv,segb) values('" + _
        form1.newid("bplan", "id", 35) & "','" + _
        h$ & "','" + _
        p$ & "'," + _
        trm(obj%(i%)) & "," + _
        d2db(obx(i%)) & "," + _
        d2db(oby(i%)) & "," + _
        trm(i%) & "," + _
        d2db(p1.ScaleHeight) & "," + _
        d2db(p1.ScaleWidth) & "," + _
        d2db(obb(0, i%)) & "," + _
        d2db(obb(1, i%)) & ",'" + _
        obd$(i%) & "'," + _
        d2db(obs(0, i%)) & "," + _
        d2db(obs(1, i%)) & ")"
    Call form1.sqlqry(c$)
  End If
Next i%

Call Command2_Click
End Sub

Private Sub Command6_Click()

'd2infile = "bplan": d2insub = "Command6_Click"
h$ = trm(hid.text)
If h$ = "" Then
  Command6.Enabled = False
  Exit Sub
End If
trgd$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(h$)
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir trgd$
On Error GoTo 0
fn$ = form1.myuniquebmpnameinpath(trgd$)
If fn$ <> "" Then
  SavePicture p1.Image, fn$
  X = Shell("explorer.exe " + form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(h$), vbNormalFocus)
End If
End Sub

Public Sub delme_Click()
'd2infile = "bplan": d2insub = "delme_Click"
List1.Clear
p1.Cls
px = -1
py = -1
co% = -1
fndi% = -1
obptr% = -1
For i% = 0 To 199: obj%(i%) = -1: obd$(i%) = "": Next i%
Text1.text = 800
Text2.text = 1000
h$ = trm(hid.text): If h$ <> "" Then Command6.Enabled = True
End Sub

Private Sub Form_Load()
'd2infile = "bplan": d2insub = "Form_Load"
fndi% = -1
Command6.Enabled = False
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
drwfast% = 1
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
'dbpara$ = form1.getconnstr()
'If dbpara$ <> "msaccessmdb" Then
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
'Else
'  Set sqla = wrkJet.OpenDatabase(form1.getdbname(), False, False)
'End If
px = -1
py = -1
co% = -1
obptr% = -1
For i% = 0 To 199: obj%(i%) = -1: obd$(i%) = "": Next i%
Text1.text = 800
Text2.text = 1000
p1.BackColor = RGB(255, 255, 255)
p1.AutoRedraw = True
p1.Cls
nupd% = 0
BackColor = form1.cleancolor()
chm% = 4
For i% = 0 To chm%: Check1(i%).BackColor = form1.cleancolor(): Next i%
Label3.Caption = ""

Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "bplan": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub


Private Sub hid_Change()
'd2infile = "bplan": d2insub = "hid_Change"
Call ldmap

End Sub

Private Sub hid_Click()
'd2infile = "bplan": d2insub = "hid_Click"
DoEvents
Call ldmap
End Sub

Private Sub hid_DropDown()
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "bplan": d2insub = "hid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT adresse.ID, adresstyp.typ FROM adresstyp INNER JOIN adresse ON adresstyp.vid = adresse.ID WHERE (((adresstyp.typ)='Halle')) OR (((adresstyp.typ)='Theater')) ORDER BY adresse.ID;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
hid.Clear
While Not r.EOF
  hid.AddItem r!id
  r.MoveNext
Wend
MousePointer = 0
End Sub


Private Sub List1_Click()

'd2infile = "bplan": d2insub = "List1_Click"
i% = List1.ListIndex
If i% < 0 Then Exit Sub

p% = Val(List1.List(i%))
co% = obj%(p%)
Select Case co%
  Case 0: sbreit(0).text = obb(0, p): Call sbreit_Change(0)
  Case 2: pbreit(0).text = obb(0, p): Call pbreit_Change(0)
  Case 3: For j% = 0 To 1: bbreit(j%).text = obb(j%, p): Next j%
          Text3.text = "1"
  Case 4: For j% = 0 To 1: stbreit(j%).text = obb(j%, p): Next j%
          Text3.text = "1"
  Case Else:
End Select
Check1(co%).value = 1
DoEvents

cob(0) = obb(0, p%)
cob(1) = obb(1, p%)
px = obx(p%)
py = oby(p%)
Call drw(co%, obx(p%), oby(p%), RGB(255, 255, 255), cob(0), cob(1), obd$(p%), obs(0, i%), obs(1, i%))
t1w% = 1
Timer1.Interval = 100
Timer1.Enabled = True
While t1w% = 1
  DoEvents
Wend
Call drw(co%, obx(p%), oby(p%), RGB(0, 0, 0), cob(0), cob(1), obd$(p%), obs(0, i%), obs(1, i%))
t1w% = 1
Timer1.Enabled = True
While t1w% = 1
  DoEvents
Wend
Call drw(co%, obx(p%), oby(p%), RGB(255, 255, 255), cob(0), cob(1), obd$(p%), obs(0, p%), obs(1, p%))
t1w% = 1
Timer1.Enabled = True
While t1w% = 1
  DoEvents
Wend
Call drw(co%, obx(p%), oby(p%), RGB(0, 0, 0), cob(0), cob(1), obd$(p%), obs(0, p%), obs(1, p%))
End Sub

Private Sub List1_DblClick()

'd2infile = "bplan": d2insub = "List1_DblClick"
i% = List1.ListIndex
If i% < 0 Then Exit Sub
p% = Val(List1.List(i%))
Call drw(co%, px, py, RGB(0, 0, 0), cob(0), cob(1), obd$(p%), obs(0, p%), obs(1, p%))
e$ = InputBox(transe("Neue Beschreibung:"), transe("Neue Beschreibung eingeben"), "")
If e$ <> "" Then
  obd$(p%) = trm(e$)
  List1.RemoveItem p%
  List1.AddItem Format$(p%, "0#") & "," & obd$(p%) & "," & Check1(co%).Caption
End If
Call Command2_Click
End Sub
Public Sub List1DClick(i%, e$)

'd2infile = "bplan": d2insub = "List1DClick"
p% = Val(List1.List(i%))
Call drw(co%, px, py, RGB(0, 0, 0), cob(0), cob(1), obd$(p%), obs(0, p%), obs(1, p%))
If e$ <> "" Then
  obd$(p%) = trm(e$)
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'd2infile = "bplan": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  i% = List1.ListIndex
  If i% < 0 Then Exit Sub
  p% = Val(List1.List(i%))
  obj%(p%) = -1
  co% = -1
  List1.RemoveItem i%
  Call Command2_Click
End If
End Sub

Public Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ostmp As Double
'd2infile = "bplan": d2insub = "p1_MouseDown"
If Button = 1 Then
  drwfast% = 0
  pco% = -1
  n% = 1
  If co% < 0 Then
    If fndi% >= 0 Then
      For i% = 0 To List1.ListCount - 1
        If Val(List1.List(i%)) = fndi% Then
          List1.ListIndex = i%
          Exit For
        End If
      Next i%
    End If
    Exit Sub
  End If
  If co% = 3 Then
    n% = Val(Text3.text)
  End If
  If co% = 4 Then
    n% = Val(Text4.text)
    If n% = 0 Then n% = 1
    ostmp = (3.14 / n%)
  End If
  p% = 0
  While n% > 0
    p% = p% + 1
    osv = 0: osb = 0
    Select Case co%
      Case 3: xoff% = (p% - 1) * (cob(0) + 20)
      Case 4: xoff% = 0: osv = (p% - 1) * ostmp: osb = osv + ostmp
            osv = osv + 0.1 * ostmp: osb = osb - 0.1 * ostmp
      Case Else: xoff% = 0
    End Select
    Call drw(co%, X + xoff%, Y, RGB(0, 0, 0), cob(0), cob(1), "", osv, osb)
    i% = List1.ListIndex
    If i% < 0 Then
      obptr% = -1
      Do
        obptr% = obptr% + 1
      Loop Until obj%(obptr%) = -1
    Else
      obptr% = Val(List1.List(i%))
    End If
    Call drw(co%, X + xoff%, Y, RGB(0, 0, 0), cob(0), cob(1), "", osv, osb)
    obj%(obptr%) = co%
    obx(obptr%) = X + xoff%
    oby(obptr%) = Y
    obs(0, obptr%) = osv: obs(1, obptr%) = osb
    Select Case co%
      Case 0: For i% = 0 To 1: obb(i%, obptr%) = cob(i%): Next i%
      Case 1: For i% = 0 To 1: obb(i%, obptr%) = cob(i%): Next i%
      Case 2: For i% = 0 To 1: obb(i%, obptr%) = cob(i%): Next i%
      Case 3: For i% = 0 To 1: obb(i%, obptr%) = cob(i%): Next i%
      Case 4: For i% = 0 To 1: obb(i%, obptr%) = cob(i%): Next i%
      Case Else: xoff% = 0
    End Select
    n% = n% - 1
  Wend
  Call Command2_Click
  DoEvents
  Call Check1_Click(co%)
  DoEvents
  drwfast% = 1
End If
Call rlist1
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cmx As Double, cmy As Double, M As Double, ostmp As Double

'd2infile = "bplan": d2insub = "p1_MouseMove"
cmy = p1.ScaleHeight * Y / p1.ScaleHeight
cmx = p1.ScaleWidth * X / p1.ScaleWidth
Label3.Caption = "x=" & trm(Int(cmx) / 100) & " m / y=" & trm(Int(cmy) / 100)
msx = X: msy = Y
n% = 1
If co% = 3 Then
  n% = Val(Text3.text)
End If
While n% > 0
  Select Case co%
    Case 3: xoff% = (n - 1) * (cob(0) + 20)
    Case 4: xoff% = 0: osv = (n - 1) * (ostmp + 0.1 * ostmp): osb = osv + 0.8 * ostmp
    Case Else: xoff% = 0
  End Select
'  Call drw(co%, x + xoff%, y, RGB(0, 0, 0))
  Call drw(pco%, px + xoff%, py, RGB(255, 255, 255), pcob(0), pcob(1), "", 0, 3.14)
  Call drw(co%, X + xoff%, Y, RGB(0, 0, 0), cob(0), cob(1), "", 0, 3.14)
  n% = n% - 1
Wend
pco% = co%
posv = osv
posb = osb
pcob(0) = cob(0)
pcob(1) = cob(1)
px = X
py = Y
If co% < 0 Then
  mi% = -1
  md = 10000000#
  For i% = 0 To 199
    If obj%(i%) >= 0 And obs(1, i%) = 0 Then
      dx = X - obx(i%): dx = dx * dx
      dy = Y - oby(i%): dy = dy * dy
      If md > dx + dy Then
        md = dx + dy
        mi% = i%
      End If
    End If
  Next i%
  If mi% >= 0 Then

fndi% = mi%
Call drw(obj%(mi%), obx(mi%), oby(mi%), RGB(255, 255, 255), obb(0, mi%), obb(1, mi%), "", obs(0, mi%), obs(1, mi%))
t1w% = 1
Timer1.Interval = 100
Timer1.Enabled = True
While t1w% = 1
  DoEvents
Wend
Call drw(obj%(mi%), obx(mi%), oby(mi%), RGB(0, 0, 0), obb(0, mi%), obb(1, mi%), "", obs(0, mi%), obs(1, mi%))
  co% = -1
  End If
End If
End Sub
Private Sub cmdPrint_Click()
Dim picture_aspect As Single
Dim printer_aspect As Single
Dim wid As Single
Dim hgt As Single
Dim X As Single
Dim Y As Single

'd2infile = "bplan": d2insub = "cmdPrint_Click"
MousePointer = 11
DoEvents
                     ' Set the PictureBox's ScaleMode to pixels to
                     ' make things interesting.
                     p1.ScaleMode = vbPixels

                     ' Compare the picture's and Printer's
                     ' aspect ratios.
                     picture_aspect = p1.ScaleHeight / _
                         p1.ScaleWidth
                     printer_aspect = printer.ScaleHeight / _
                         printer.ScaleWidth
                     If pic_aspect > printer_aspect Then
                         ' The picture is too tall and thin.
                         ' Print it as tall as possible.
                         hgt = printer.ScaleHeight
                         wid = hgt / picture_aspect
                     Else
                         ' The picture is too short and wide.
                         ' Print it as wide as possible.
                         wid = printer.ScaleWidth
                         hgt = wid * picture_aspect
                     End If

                     ' See where we need to place the picture
                     ' to center it.
                     X = printer.ScaleLeft + (printer.ScaleWidth - wid) / 2
                     Y = printer.ScaleTop + (printer.ScaleHeight - hgt) / 2

                     ' Print the picture.
                     printer.PaintPicture p1.Image, X, Y, wid, hgt
For i% = 0 To 199
  If obj%(i%) >= 0 Then
    Call prndrw(obj%(i%), obx(i%), oby(i%), RGB(0, 0, 0), obb(0, i%), obb(1, i%), obd$(i%))
  End If
Next i%

                     ' Draw the box.
                     printer.Line (X, Y)-Step(wid, hgt), , B

                     ' Finish printing.
                     printer.EndDoc

MousePointer = 0
End Sub

Private Sub pbreit_Change(Index As Integer)
'd2infile = "bplan": d2insub = "pbreit_Change"
If Index = 0 Then
  pbreit(1).text = pbreit(0).text
  If co% = 2 Then
    cob(0) = Val(pbreit(0).text)
    cob(1) = cob(0)
  End If
End If
End Sub

Private Sub pgid_Change()
'd2infile = "bplan": d2insub = "pgid_Change"
Call ldmap
End Sub

Private Sub pgid_Click()
'd2infile = "bplan": d2insub = "pgid_Click"
DoEvents
Call ldmap

End Sub

Private Sub pgid_DropDown()
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "bplan": d2insub = "pgid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT programmid FROM programm order by programmid;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
pgid.Clear
While Not r.EOF
  pgid.AddItem r!programmid
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Sub sbreit_Change(Index As Integer)
'd2infile = "bplan": d2insub = "sbreit_Change"
If Index = 0 Then
  sbreit(1).text = sbreit(0).text
  If co% = 2 Then
    cob(0) = Val(sbreit(0).text)
    cob(1) = cob(0)
  End If
End If

End Sub


Private Sub Text1_Change()
'd2infile = "bplan": d2insub = "Text1_Change"
p1.ScaleHeight = Val(Text1.text)
End Sub

Private Sub Text2_Change()
'd2infile = "bplan": d2insub = "Text2_Change"
p1.ScaleWidth = Val(Text2.text)
End Sub

Public Sub ldmap()
Dim h$, p$, r As ADODB.Recordset, c$
Dim d2infile As String, d2insub As String
d2infile = "bplan": d2insub = "ldmap"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" And p$ = "" Then Exit Sub
drwfast% = 0
c$ = "select * from bplan where adressid='" & h$ & "' and prgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Call delme_Click
  Text1.text = r!mx
  Text2.text = r!my
  DoEvents
  While Not r.EOF
    i% = r!objnr
    obj%(i%) = r!objtyp
    obx(i%) = r!objx
    oby(i%) = r!objy
    obb(0, i%) = r!obx
    obb(1, i%) = r!oby
    obs(0, i%) = 0: If Not IsNull(r!segv) Then obs(0, i%) = r!segv
    obs(1, i%) = 0: If Not IsNull(r!segb) Then obs(1, i%) = r!segb
    obd$(i%) = trm(r!obdesc)
    Call drw(obj%(i%), obx(i%), oby(i%), RGB(0, 0, 0), obb(0, i%), obb(1, i%), obd$(i%), obs(0, i%), obs(1, i%))
    List1.AddItem Format$(i%, "0#") & "," & obd$(i%) & "," & Check1(obj%(i)).Caption
    DoEvents
    r.MoveNext
  Wend
End If
drwfast% = 1
Call drwlegend
If h$ <> "" Then Command6.Enabled = True
End Sub

Private Sub Timer1_Timer()
'd2infile = "bplan": d2insub = "Timer1_Timer"
t1w% = 0
Timer1.Enabled = False
End Sub

Sub drw(o%, X, Y, col As Long, obb, obt, dsc$, segv As Double, segb As Double)
Dim i As Integer, xo%

'd2infile = "bplan": d2insub = "drw"
Select Case o%
  Case 0: p1.Circle (X, Y), obb / 2, col
          p1.Line (X, Y)-(X - obb / 3, Y - obt / 3), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              p1.Print dsc$
            Else
              p1.Print Space$(Len(dsc$))
            End If
          End If
  Case 1: sdx = 20
          p1.Line (X - obb / 2, Y - obt / 2)-(X + obb / 2, Y - obt / 2), col
          p1.Line (X + obb / 2, Y - obt / 2)-(X + obb / 2, Y + 100 - obt / 2), col
          p1.Line (X - obb / 2, Y - obt / 2)-(X - obb / 2, Y + obt / 2), col
          p1.Line (X - obb / 2, Y + obt / 2)-(X + 50 - obb / 2, Y + obt / 2), col
          p1.Line (X + 50 - obb / 2, Y + obt / 2)-(X + obb / 2, Y + 100 - obt / 2), col
          p1.Line (X - obb / 2, Y - obt / 2)-(X - obb / 2, Y - obt / 2), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              p1.Print dsc$
            Else
              p1.Print Space$(Len(dsc$))
            End If
          End If
  Case 2: p1.Circle (X, Y), obb / 2, col
          p1.Line (X, Y)-(X - obb / 3, Y - obt / 3), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              p1.Print dsc$
            Else
              p1.Print Space$(Len(dsc$))
            End If
          End If
          p1.Line (X - obb / 2, Y - obb / 2)-(X - obb / 2, Y + obb / 2), col
          p1.Line (X + obb / 2, Y - obb / 2)-(X + obb / 2, Y + obb / 2), col
          p1.Line (X - 20 + obb / 2, Y + obb / 2)-(X + 20 + obb / 2, Y + obb / 2), col
          p1.Line (X - 20 - obb / 2, Y + obb / 2)-(X + 20 - obb / 2, Y + obb / 2), col
  Case 3: p1.Line (X - obb / 2, Y - obt / 2)-(X + obb / 2, Y + obt / 2), col, B
          If dsc$ <> "" Then
            p1.Line (X + obb / 2, Y + obt / 2)-(X - obb / 2, Y - obt / 2), RGB(255, 255, 255)
            If col = 0 Then
              p1.Print dsc$
            Else
              p1.Print Space$(Len(dsc$))
            End If
          End If
  Case 4:
          If drwfast% = 1 Then
            p1.Circle (X, Y), imax(80, obb / 2 - obt / 2), col, segv, segb
            p1.Circle (X, Y), imin(900, obb / 2 + obt / 2), col, segv, segb
          Else
            p1x = X + imax(80, obb / 2 - obt / 2) * Cos(segb)
            p1y = Y - imax(80, obb / 2 - obt / 2) * Sin(segb)
            p2x = X + imax(80, obb / 2 - obt / 2) * Cos(segv)
            p2y = Y - imax(80, obb / 2 - obt / 2) * Sin(segv)
            p3x = X + imin(900, obb / 2 + obt / 2) * Cos(segb)
            p3y = Y - imin(900, obb / 2 + obt / 2) * Sin(segb)
            p4x = X + imin(900, obb / 2 + obt / 2) * Cos(segv)
            p4y = Y - imin(900, obb / 2 + obt / 2) * Sin(segv)

            p1.Line (p1x, p1y)-(p2x, p2y)
            p1.Line (p3x, p3y)-(p4x, p4y)
            p1.Line (p1x, p1y)-(p3x, p3y)
            p1.Line (p2x, p2y)-(p4x, p4y)
            p1.Line ((p1x + p3x) / 2, (p1y + p4y) / 2 - 8)-((p1x + p3x) / 2, (p1y + p4y) / 2 - 8), RGB(255, 255, 255)
            If dsc$ <> "" Then
              If col = 0 Then
                p1.Print dsc$
              Else
                p1.Print Space$(Len(dsc$))
              End If
            End If
          End If
  Case Else:
End Select

End Sub
Sub rlist1()
'd2infile = "bplan": d2insub = "rlist1"
List1.Clear
For i% = 0 To 199
  If obj%(i%) >= 0 Then
    List1.AddItem Format$(i%, "0#") & "," & obd$(i%) & "," & Check1(obj%(i%)).Caption
  End If
Next i%
End Sub

Sub prndrw(o%, X, Y, col As Long, obb, obt, dsc$)
Dim i As Integer, xo%

'd2infile = "bplan": d2insub = "prndrw"
Select Case o%
  Case 0: printer.Circle (X, Y), obb / 2, col
          printer.Line (X, Y)-(X - obb / 3, Y - obt / 3), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              printer.Print dsc$
            End If
          End If
  Case 1: sdx = 20
          printer.Line (X - obb / 2, Y - obt / 2)-(X + obb / 2, Y - obt / 2), col
          printer.Line (X + obb / 2, Y - obt / 2)-(X + obb / 2, Y + 100 - obt / 2), col
          printer.Line (X - obb / 2, Y - obt / 2)-(X - obb / 2, Y + obt / 2), col
          printer.Line (X - obb / 2, Y + obt / 2)-(X + 50 - obb / 2, Y + obt / 2), col
          printer.Line (X + 50 - obb / 2, Y + obt / 2)-(X + obb / 2, Y + 100 - obt / 2), col
          printer.Line (X - obb / 2, Y - obt / 2)-(X - obb / 2, Y - obt / 2), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              printer.Print dsc$
            End If
          End If
  Case 2: printer.Circle (X, Y), obb / 2, col
          printer.Line (X, Y)-(X - obb / 3, Y - obt / 3), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              printer.Print dsc$
            End If
          End If
          printer.Line (X - obb / 2, Y - obb / 2)-(X - obb / 2, Y + obb / 2), col
          printer.Line (X + obb / 2, Y - obb / 2)-(X + obb / 2, Y + obb / 2), col
          printer.Line (X - 20 + obb / 2, Y + obb / 2)-(X + 20 + obb / 2, Y + obb / 2), col
          printer.Line (X - 20 - obb / 2, Y + obb / 2)-(X + 20 - obb / 2, Y + obb / 2), col
  Case 3: printer.Line (X - obb / 2, Y - obt / 2)-(X + obb / 2, Y + obt / 2), col, B
          If dsc$ <> "" Then
            printer.Line (X + obb / 2, Y + obt / 2)-(X - obb / 2, Y - obt / 2), RGB(255, 255, 255)
            If col = 0 Then
              printer.Print dsc$
            End If
          End If
  Case 4: printer.Circle (X, Y), obb / 2, col
          printer.Line (X, Y)-(X - obb / 3, Y - obt / 3), RGB(255, 255, 255)
          If dsc$ <> "" Then
            If col = 0 Then
              printer.Print dsc$
            End If
          End If
  Case Else:
End Select

End Sub

Sub drwlegend()
'd2infile = "bplan": d2insub = "drwlegend"
p1.Line (0, 0)-(p1.ScaleWidth, 0)
p1.Line (10, 10)-(30, 0)
p1.Line (10, 10)-(30, 20)
p1.Line (p1.ScaleWidth - 10, 10)-(p1.ScaleWidth - 30, 0)
p1.Line (p1.ScaleWidth - 10, 10)-(p1.ScaleWidth - 30, 20)
p1.Line (p1.ScaleWidth / 2, 15)-(p1.ScaleWidth / 2, 15)
p1.Print trm(Int(p1.ScaleWidth)) & "cm"

p1.Line (10, 10)-(10, p1.ScaleHeight - 10)
p1.Line (0, 0)-(0, p1.ScaleHeight)
p1.Line (10, 10)-(20, 30)
p1.Line (10, 10)-(0, 30)
p1.Line (10, p1.ScaleHeight - 10)-(20, p1.ScaleHeight - 30)
p1.Line (10, p1.ScaleHeight - 10)-(0, p1.ScaleHeight - 30)
p1.Line (15, p1.ScaleHeight / 2)-(15, p1.ScaleHeight / 2)
p1.Print trm(Int(p1.ScaleHeight)) & "cm"

For X = 0 To p1.ScaleWidth Step 50
  dh = 10: If (X Mod 100) = 0 Then dh = 30
  p1.Line (X, p1.ScaleHeight)-(X, p1.ScaleHeight - dh)
  If dh = 30 Then
    p1.Line (X - 50, p1.ScaleHeight - dh * 2)-(X - 50, p1.ScaleHeight - dh * 2)
    p1.Print X
  End If
Next X
For Y = 0 To p1.ScaleHeight Step 50
  dh = 10: If (Y Mod 100) = 0 Then dh = 30
  p1.Line (p1.ScaleWidth, Y)-(p1.ScaleWidth - dh, Y)
  If dh = 30 Then
    p1.Line (p1.ScaleWidth - 2 * dh, Y)-(p1.ScaleWidth - 2 * dh, Y)
    p1.Print Y
  End If
Next Y
End Sub

Public Function getobd(i%)
'd2infile = "bplan": d2insub = "getobd"
On Error Resume Next
getobd = obd$(i%)
On Error GoTo 0
End Function
