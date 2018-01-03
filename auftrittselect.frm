VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form auftrittselect 
   Caption         =   "Auftritte und Termine"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "XML"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   3420
      Width           =   615
   End
   Begin VB.TextBox vts 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   13
      ToolTipText     =   "Termine nach diesem Text durchsuchen"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command26 
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
      Height          =   375
      Left            =   1920
      Picture         =   "auftrittselect.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Neuen Termin erstellen"
      Top             =   2700
      Width           =   375
   End
   Begin VB.TextBox eintermin 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   10
      ToolTipText     =   "erstellt einen Termin zum eingetragenen Datum (leer=heute)"
      Top             =   2820
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "?"
      Enabled         =   0   'False
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
      Left            =   840
      TabIndex        =   9
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3420
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   0
      Picture         =   "auftrittselect.frx":0392
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Formular schiessen"
      Top             =   3420
      Width           =   855
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Timer Timer2 
      Left            =   960
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   2280
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&alle zeigen"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "30"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "30"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3420
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   3840
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Volltextsuche"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "Sucht nach Text in Terminen"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NeuerTerm."
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      ToolTipText     =   "Doppelclick=heute"
      Top             =   2820
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "max.Tg.zur."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "max.Tg.vor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "auftrittselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selstr$, abreq%
Dim toffsvon As Long, toffsbis As Long

Private Sub Command1_Click()

'd2infile = "auftrittselect": d2insub = "Command1_Click"
abreq% = 1
DoEvents
Hide
Unload auftrittselect
Unload auftritt

End Sub


Private Sub Command2_Click()

'd2infile = "auftrittselect": d2insub = "Command2_Click"
Call rlist1

End Sub


Private Sub Command26_Click()
Dim nid$, tpid$, d0

Dim d2infile As String, d2insub As String
d2infile = "auftrittselect": d2insub = "Command26_Click"
tpid$ = "-1"
MousePointer = 11: DoEvents
d0 = CDate(Date)
nid$ = form1.newid("auftritt", "id", 20)
If eintermin.text <> "" Then d0 = CDate(eintermin.text)
nid$ = form1.newid("auftritt", "id", 20)
form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
               nid$ & "','" + tpid$ + _
               "','Neuer Auftritt','" + tpid$ & "','" + _
               datum2sql(CDate(d0)) & "')")
Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(nid$, 0)
MousePointer = 0

End Sub

Private Sub Command3_Click()
Dim id$, i%, o%, X, indent%, r As ADODB.Recordset, cmd$, rrr
Dim ra As ADODB.Recordset, kn$
Dim rp As ADODB.Recordset, xml$

MousePointer = 11
xml$ = form1.mydatadir() + "\auftritt.xml"
o% = FreeFile
Open xml$ For Output As #o%
indent% = 0
Print #o%, "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
Print #o%, "<agencyprof versionxml=""1.0"">"
indent% = indent% + 2
For i% = 0 To List1.ListCount - 1
  List1.ListIndex = i%: DoEvents
  id$ = List1.List(i%)
  id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
  Print #o%, Space$(indent%) + "<auftritt>"
  indent% = indent% + 2
  Print #o%, Space$(indent%) + "<auftrittid>" + trm(id$) + "</auftrittid>"
  cmd$ = "SELECT * FROM auftritt where id= '" & id$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "auftrittselect", "command3")
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  If Not r.EOF Then
    Print #o%, Space$(indent%) + "<auftrittdatum>" + trm(r!datum) + "</auftrittdatum>"
    Print #o%, Space$(indent%) + "<auftrittbeginn>" + trm(r!zeit) + "</auftrittbeginn>"
    Print #o%, Space$(indent%) + "<auftrittort>" + me2utf8(trm(r!ort)) + "</auftrittort>"
    cmd$ = "SELECT feldname, felddaten FROM auftritthigru where auftrittsid= '" & id$ & "'"
    Set ra = New ADODB.Recordset
    ra.CursorLocation = adUseServer
    On Error Resume Next
    rrr = form1.adoopen(ra, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "auftrittselect", "command3")
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then Exit Sub
    While Not ra.EOF
      If ra!feldname <> "Programm" Then
        Print #o%, Space$(indent%) + "<auftritt" + me2utf8(trm(ra!feldname)) + ">" + me2utf8(trm(ra!felddaten)) + "</auftritt" + trm(ra!feldname) + ">"
      Else
        cmd$ = "SELECT WerkID,Position FROM programmliste where ProgrammID= '" & trm(ra!felddaten) & "' order by Position"
        Set rp = New ADODB.Recordset
        rp.CursorLocation = adUseServer
        rrr = form1.adoopen(rp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "auftrittselect", "command3")
        If rrr = 0 Then
          Print #o%, Space$(indent%) + "<rowset name=""Programm"" key=""Position"" columns=""Position,Komponist,Werk"">"
          indent% = indent% + 2
          While Not rp.EOF
            Print #o%, Space$(indent%) + "<row Position=""" + trm(rp!Position) + """ ";
            kn$ = me2utf8(form1.getkompnamebywerkid(trm(rp!werkid)))
            Print #o%, "Komponist=""" + kn$ + """ ";
            Print #o%, "Werk=""" + me2utf8(form1.getwerknamebyid(trm(rp!werkid))) + """ />"
            rp.MoveNext
          Wend
          indent% = indent% - 2
          Print #o%, Space$(indent%) + "</rowset>"
        End If
      End If
      ra.MoveNext
    Wend
  End If
  indent% = indent% - 2
  Print #o%, Space$(indent%) + "</auftritt>"
Next i%
Print #o%, "</agencyprof>"
Close #o%
MousePointer = 0
X = Shell("notepad.exe " + xml$, 1)
End Sub

Private Sub Command6_Click()
Dim i%

'd2infile = "auftrittselect": d2insub = "Command6_Click"
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = True Then List2.Selected(i%) = False
Next i%
vts.text = ""
Call rlist1

End Sub


Private Sub eintermin_DblClick()
Dim p$
  p$ = eintermin.text
  With frmCalendar
    .init eintermin, eintermin.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      eintermin.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar

End Sub

Private Sub Form_Load()
Dim s%

'd2infile = "auftrittselect": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
abreq% = 0
toffsvon = Val(form1.mylastFormVar(Me.name, "toffsvon", "30"))
toffsbis = Val(form1.mylastFormVar(Me.name, "toffsbis", "30"))
Text1.text = trm(toffsvon)
Text2.text = trm(toffsbis)
s% = form1.myfontsize()
List1.Font.Size = s%
auftrittselect.Caption = form1.inmylanguage("Auftritte und Termine")
Command18.Caption = form1.inmylanguage("?")
Command1.ToolTipText = form1.inmylanguage("Formular schliessen")
Command6.Caption = form1.inmylanguage("&alle zeigen")
Command2.Caption = form1.inmylanguage("&GO")
Label2.Caption = form1.inmylanguage("max.Tg.zur.")
Label1.Caption = form1.inmylanguage("max.Tg.vor")
Show
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)



selstr$ = ""
Call rlist2
Call rlist1

End Sub

Sub rlist2()
Dim rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "auftrittselect": d2insub = "rlist2"
List2.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM auftrittstypen order by sortierung", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While Not rtmp.EOF
  List2.AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend
List2.ListIndex = -1

End Sub

Sub rlist1()
Dim r As ADODB.Recordset, seli%
Dim r1 As ADODB.Recordset
Dim dv$, db$, msg$, rrr, rrr2, vtxts$, c$, shw As Boolean

Dim d2infile As String, d2insub As String
d2infile = "auftrittselect": d2insub = "rlist1"
MousePointer = 11: DoEvents
vtxts$ = LCase(trm(vts.text))
If vtxts$ <> "" Then vtxts$ = strrepl(vtxts$, "'", "´")
On Error GoTo exr1
List1.Clear
selstr$ = getsel()
dv$ = datum2sql(Date - toffsvon)
db$ = datum2sql(Date + toffsbis)
If selstr$ = "" Then
  selstr$ = "where "
Else
  selstr$ = selstr$ + " and "
End If
selstr$ = selstr$ + "(datum>='" + dv$ + "' and datum<='" + db$ + "')"
Call form1.dbg2f(selstr$)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM auftritt " + selstr$ + " order by datum", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

seli% = -1
While Not r.EOF And abreq% = 0
  shw = True
  If vtxts$ <> "" Then
    shw = False
    If InStr(LCase(r!bezeichnung), vtxts$) > 0 Or InStr(LCase(r!ort), vtxts$) > 0 Then
      shw = True
    Else
      c$ = "select * from auftritthigru where auftrittsid='" + r!id + "' and instr(lcase(felddaten),'" + vtxts$ + "')>0"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
      rrr2 = form1.adoopen(r1, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr2 = 0 Then
        If Not r1.EOF Then shw = True
      End If
    End If
  End If
  If shw Then
    msg$ = form1.dayofweek(r!datum) + ", " & r!datum & " " & form1.get_atabkz(r!auftrittstyp)
    If Not IsNull(r!ort) Then msg$ = msg$ & " " & r!ort & " "
    msg$ = msg$ & "(" & r!bezeichnung & ")"
    msg$ = msg$ & Space$(80) + "(AID:" & r!id
    List1.AddItem msg$
    If r!datum >= datum2sql(Date) And seli% = -1 Then seli% = List1.ListCount - 1
  End If
  r.MoveNext
  DoEvents
Wend
If abreq% = 1 Then abreq% = 0
r.Close
If seli% >= 0 Then List1.ListIndex = seli%
MousePointer = 0
Exit Sub
exr1:
abreq% = 1
DoEvents
On Error GoTo 0

End Sub

Private Sub Form_Resize()
'd2infile = "auftrittselect": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "auftrittselect": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub List1_DblClick()
Dim id$, i%, j%

'd2infile = "auftrittselect": d2insub = "List1_DblClick"
MousePointer = 11
id$ = List1.List(List1.ListIndex)
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)

If form1.immerkalender() = "ja" Then
  Load kc
  For i% = 0 To List2.ListCount - 1
    If List2.Selected(i%) = True Then
      For j% = 0 To kc.selct(0).ListCount - 1
        If List2.List(i%) = kc.selct(0).List(j%) Then
          kc.selct(0).Selected(j%) = True
          j% = kc.selct(0).ListCount
        End If
      Next j%
    End If
  Next i%
End If

Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(id$, 0)
MousePointer = 0
End Sub
Private Function getsel() As String
Dim s$, i%

'd2infile = "auftrittselect": d2insub = "getsel"
s$ = ""
For i% = 0 To List2.ListCount - 1
  If List2.Selected(i%) = True Then
    If Len(s$) = 0 Then
      s$ = "where ((auftrittstyp='" + transo(List2.List(i%)) + "') "
    Else
      s$ = s$ + "or (auftrittstyp='" + transo(List2.List(i%)) + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"
getsel = s$

End Function

Private Sub Text1_Change()
'd2infile = "auftrittselect": d2insub = "Text1_Change"
toffsvon = Val(Text1.text)
Call form1.setmylastFormVar(Me.name, "toffsvon", Text1.text)
End Sub

Private Sub Text1_DblClick()
'd2infile = "auftrittselect": d2insub = "Text1_DblClick"
  With frmCalendar
    .init Text1, Text1.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar
  Text1.text = CDate(datum2sql("" & Date)) - CDate(datum2sql(Text1.text))

End Sub

Private Sub Text2_Change()
'd2infile = "auftrittselect": d2insub = "Text2_Change"
toffsbis = Val(Text2.text)
Call form1.setmylastFormVar(Me.name, "toffsbis", Text2.text)
End Sub

Private Sub Text2_DblClick()
'd2infile = "auftrittselect": d2insub = "Text2_DblClick"
  With frmCalendar
    .init Text2, Text2.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text2.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar

  Text2.text = CDate(datum2sql(Text2.text)) - CDate(datum2sql("" & Date))
End Sub


