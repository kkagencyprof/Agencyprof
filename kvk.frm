VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form kvk 
   Caption         =   "Kartenverkauf"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check2 
      Caption         =   "&Rechnung drucken"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ListBox ralist 
      Height          =   1425
      IntegralHeight  =   0   'False
      Left            =   5280
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox aboplatzids 
      Height          =   2790
      IntegralHeight  =   0   'False
      Left            =   7200
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton epb 
      Caption         =   "Endbetrag: "
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abbruch"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.ListBox gd1ids 
      Height          =   4005
      IntegralHeight  =   0   'False
      Left            =   8520
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox psel 
      Height          =   645
      Index           =   1
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox psel 
      Height          =   645
      Index           =   0
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox selstr 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2055
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
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   4560
      Width           =   255
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   240
      Top             =   4080
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6800
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kontakt"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Vorgang"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   16
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label epn 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Netto Endbetrag:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label epm 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "--,--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Summe MwSt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label vg 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Von"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu menu_t1 
      Caption         =   "Änderungen"
      Visible         =   0   'False
      Begin VB.Menu menu_chg_mwst 
         Caption         =   "MwSt ändern"
      End
      Begin VB.Menu menu_chg_pb 
         Caption         =   "Bruttopreis ändern"
      End
      Begin VB.Menu menu_chg_proz 
         Caption         =   "Prozentualer Nachlass"
      End
   End
End
Attribute VB_Name = "kvk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currkid$, pcode

Private Sub Check2_Click()
'd2infile = "kvk": d2insub = "Check2_Click"
Call form1.setmylastFormVar(Me.name, "printrech", trm(Check2.value))
End Sub

Private Sub combo1_Change(Index As Integer)

'd2infile = "kvk": d2insub = "combo1_Change"
If Index = 0 Then fld$ = "vonid="
If Index = 1 Then fld$ = "kontakt="

psel(Index).text = fld$ & "'" & trm(Combo1(Index).text) & "'"

End Sub

Private Sub Command1_Click()
Dim c$

'd2infile = "kvk": d2insub = "Command1_Click"
MousePointer = 11: DoEvents
If InStr(Combo1(0).text, "Barverkauf " & form1.getuserid() & " ") = 1 Then
  c$ = "delete from hbpstatus where adrid='" & Combo1(0).text & "'"
  Call form1.sqlqry(c$)
  Call splan.Command2_Click
  Call splan.beglist_Change
End If
MousePointer = 0: DoEvents
splan.Text6.text = ""
Unload Me

End Sub

Private Sub Command18_Click()
'd2infile = "kvk": d2insub = "Command18_Click"
Call form1.handbuchcall("index.html")
End Sub

Private Sub epb_Click()
Dim c$, i%, hbpid$, kid$, aboid$, r As ADODB.Recordset, dtg$, r1 As ADODB.Recordset
Dim o%, n1%, r2 As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "kvk": d2insub = "epb_Click"
kid$ = trm(Combo1(1).text): If kid$ = "" Then kid$ = "-1"
dtg$ = Left(vg.Caption, 16)
For i% = 0 To gd1ids.ListCount - 1
  hbpid$ = gd1ids.List(i%)
  If hbpid$ <> "NULL" Then
    c$ = "select * from hbpstatus where (((hbpstatus.hbpid)='" & hbpid$ & "') AND ((hbpstatus.dtg)='" & datum2sql(word1(dtg$)) & " " & word2(dtg$) & "') AND ((hbpstatus.pstatus)='Bestellung') AND ((hbpstatus.adrid)='" & Combo1(0).text & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    If Not r.EOF Then
      zstat$ = ""
      c$ = "SELECT hbpstatus_1.pstatus2 as ps2, hbpstatus_1.aboplatzid as aid " + _
         "FROM (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) INNER JOIN hbpstatus AS hbpstatus_1 ON hbplist.id = hbpstatus_1.hbpid " + _
         "WHERE (((hbpstatus.hbpid)='" & hbpid$ & "') AND ((hbpstatus_1.pstatus)='Abo') AND ((hbpstatus.dtg)='" & datum2sql(word1(dtg$)) & " " & word2(dtg$) & "') AND ((hbpstatus.pstatus)='bestellung') AND ((hbpstatus.adrid)='" & Combo1(0).text & "'));"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r1.EOF Then zstat$ = r1!ps2 & "/" & trm(r1!aid)
      c$ = "update hbpstatus set pstatus='Verkauft' where (((hbpstatus.hbpid)='" & hbpid$ & "') AND ((hbpstatus.dtg)='" & datum2sql(word1(dtg$)) & " " & word2(dtg$) & "') AND ((hbpstatus.pstatus)='Bestellung') AND ((hbpstatus.adrid)='" & Combo1(0).text & "'));"
      Call form1.sqlqry(c$)
      c$ = "insert into kassenbuch (id,thema,dtg,vorgang,zahlstatus,vonid,kontaktname,anzahl,Bezeichnung,epreisnetto,mwst) values('" + _
      form1.newid("kassenbuch", "id", 30) & "','Kartenverkauf','" + _
      datum2sql(Date) & " " & Time & "','" + _
      Right(vg.Caption, 240) & "','" + _
      zstat$ & "','" + _
      Combo1(0).text & "','" + _
      kid$ & "',1,'" + _
      gd1.ListItems(i% + 1).SubItems(1) & "'," + _
      d2db(gd1.ListItems(i% + 1).SubItems(2)) & "," + _
      d2db(gd1.ListItems(i% + 1).SubItems(4)) & ")"
      Call form1.sqlqry(c$)

      c$ = "SELECT hbpstatus.pstatus, hbpstatus.pstatus2, hbpstatus.aboplatzid, " + _
                   "hbpstatus_1.pstatus, hbpstatus_1.pstatus2 as aid, hbpstatus_1.aboplatzid as abpid, " + _
                   "hbpstatus.dtg, hbpstatus.id as tdelid " + _
           "FROM hbpstatus INNER JOIN hbpstatus AS hbpstatus_1 ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbpstatus.hbpid = hbpstatus_1.hbpid) " + _
           "WHERE (((hbpstatus.pstatus)='Bestellung') AND ((hbpstatus_1.pstatus)='Abo') " + _
           "AND ((hbpstatus.hbpid)='" & r!hbpid & "'));"
Set r1 = New ADODB.Recordset
r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r1.EOF Then
        While Not r1.EOF
          c$ = "SELECT hbpstatus.aboplatzid, hbpstatus.pstatus2, hbpstatus.pstatus, hbpstatus.hbpid, hbpstatus.dtg, hbpstatus_1.pstatus, hbpstatus_1.id as delme " + _
               "FROM (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) INNER JOIN hbpstatus AS hbpstatus_1 ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbplist.id = hbpstatus_1.hbpid) " + _
               "WHERE (((hbpstatus.aboplatzid)=" & r1!abpid & ") AND " + _
                     "((hbpstatus.pstatus2)='" & r1!aid & "') AND " + _
                     "((hbpstatus_1.pstatus)='Bestellung'));"
Set r2 = New ADODB.Recordset
r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
          While Not r2.EOF
            c$ = "delete from hbpstatus where id='" & r2!delme & "';"
            Call form1.sqlqry(c$)
            r2.MoveNext
          Wend
          r1.MoveNext
        Wend
      Else
        c$ = "delete from hbpstatus where (((hbpstatus.hbpid)='" & hbpid$ & "') " + _
             "AND ((hbpstatus.dtg)='" & datum2sql(word1(dtg$)) & " " & word2(dtg$) & "') " + _
             "AND ((hbpstatus.pstatus)='Bestellung') AND ((hbpstatus.adrid)='" & Combo1(0).text & "'));"
        Call form1.sqlqry(c$)
      End If
      c$ = "select * from hbpstatus where (((hbpstatus.hbpid)='" & hbpid$ & "') AND ((hbpstatus.pstatus)='Abo'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        c$ = "update hbpstatus set pstatus='Abo verkauft' where (((hbpstatus.pstatus2)='" & r!pstatus2 & "') AND ((hbpstatus.aboplatzid)=" & r!aboplatzid & "));"
        Call form1.sqlqry(c$)
      End If
    End If
  End If
Next i%
If Check2.value <> 0 Then
  Load kbuch
  On Error Resume Next
  Call kbuch.SetFocus
  On Error GoTo 0
  DoEvents
  For n1% = kbuch.gd1.ListItems.Count To 1 Step -1: kbuch.gd1.ListItems(n1%).Selected = False: Next n1%
  kbuch.gd1.ListItems(kbuch.gd1.ListItems.Count).Selected = True
  Call kbuch.gd1_Click
  DoEvents
  Call kbuch.Command3_Click
  DoEvents
  Unload kbuch
End If
splan.Text6.text = ""
Call splan.Command2_Click
Call splan.beglist_Change
Unload Me

End Sub

Private Sub Form_Load()
'd2infile = "kvk": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
gd1.View = lvwReport

Set colHeader = gd1.ColumnHeaders.add(, , "Anzahl", 700)
Set colHeader = gd1.ColumnHeaders.add(, , "Bezeichnung", 4900)
Set colHeader = gd1.ColumnHeaders.add(, , "Netto", 1200)
Set colHeader = gd1.ColumnHeaders.add(, , "% MwSt", 800)
Set colHeader = gd1.ColumnHeaders.add(, , "Betrag MwSt", 1100)
Set colHeader = gd1.ColumnHeaders.add(, , "Brutto", 1200)
Check2.BackColor = form1.dirtycolor()
Me.BackColor = form1.dirtycolor()
klrv% = Val(form1.mylastFormVar(Me.name, "printrech", "0"))
If klrv% <> 0 Then klrv% = 1
Check2.value = klrv%
kvk.Caption = transe("Kartenverkauf")
Check2.Caption = transe("&Rechnung drucken")
epb.Caption = transe("Endbetrag: ")
Command1.Caption = transe("Abbruch")
Command18.Caption = transe("?")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Label1(2).Caption = transe("Kontakt")
Label1(1).Caption = transe("Vorgang")
Label8(2).Caption = transe("Netto Endbetrag:")
Label8(0).Caption = transe("Summe MwSt")
Label1(0).Caption = transe("Von")
menu_t1.Caption = transe("Änderungen")
menu_chg_mwst.Caption = transe("MwSt ändern")
menu_chg_pb.Caption = transe("Bruttopreis ändern")
menu_chg_proz.Caption = transe("Prozentualer Nachlass")
Show

End Sub
Public Sub gd1_clear()
'd2infile = "kvk": d2insub = "gd1_clear"
gd1.ListItems.Clear
gd1ids.Clear
aboplatzids.Clear
End Sub
Private Sub Form_Resize()
'd2infile = "kvk": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "kvk": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub

Private Sub gd1_Click()

'd2infile = "kvk": d2insub = "gd1_Click"
If gd1.ListItems.Count <= 0 Then Exit Sub
id$ = gd1.SelectedItem
p% = InStr(id$, "(ID:"): If p% = 0 Then Exit Sub
id$ = Mid$(id$, p% + 4)

End Sub

Private Sub gd1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, j%, rrr, n$, V$
Dim lvitem

'd2infile = "kvk": d2insub = "gd1_KeyDown"
'<strg>a
If KeyCode = 65 And pcode = 17 Then
  For i = gd1.ListItems.Count To 1 Step -1
    'Set lvitem = gd1.ListItems(i)
    gd1.ListItems(i).Selected = True
  Next i
End If
pcode = KeyCode

End Sub

Private Sub gd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'd2infile = "kvk": d2insub = "gd1_MouseDown"
If Button = 2 Then
  PopupMenu menu_t1
  Exit Sub
End If

End Sub

Private Sub psel_Change(Index As Integer)

'd2infile = "kvk": d2insub = "psel_Change"
selstr.text = trm("select * from kassenbuch where " & psel(0).text) & " "
If trm(psel(1).text) <> "" Then
  selstr.text = selstr.text & " and " & psel(1).text
End If

End Sub
