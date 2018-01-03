VERSION 5.00
Begin VB.Form higruselect 
   Caption         =   "Query"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   ScaleHeight     =   3585
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   6480
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   16
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Start"
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
      Left            =   3840
      TabIndex        =   15
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00000000&
      Picture         =   "higruselect.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   14
      ToolTipText     =   "Speichern"
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      MaskColor       =   &H00000000&
      Picture         =   "higruselect.frx":0672
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Speichern"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   120
      Width           =   4455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "higruselect.frx":0CE4
      Left            =   4800
      List            =   "higruselect.frx":0CEE
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "higruselect.frx":0CF7
      Left            =   240
      List            =   "higruselect.frx":0D07
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Text            =   "%"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
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
      Height          =   1740
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "higruselect.frx":0D19
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Auf Wiedersehen!"
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "In Kontakten suchen"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Field"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "In Kontakten suchen"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label getroffn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "In Kontakten suchen"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "higruselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c3nochg As Boolean

Private Sub Combo2_DropDown()
Dim cmd$, s As ADODB.Recordset, rrr, t$

t$ = transo(trm(Combo1.text))
Combo2.Clear
cmd$ = "SELECT FeldName From auftrittsfelder where typ='" + t$ + "' order by position"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
While Not s.EOF
  Combo2.AddItem trm(s!feldname)
  s.MoveNext
Wend
End If
End Sub

Private Sub Combo3_Change()
If c3nochg Then Exit Sub
If Len(Combo3.text) > 1 Then
  If List1.ListCount > 0 Then List1.AddItem Combo3.text
  Combo3.text = ""
End If
End Sub

Private Sub Combo3_Click()
  Call Combo3_Change
End Sub

Private Sub Combo5_Change()
Dim f$

f$ = form1.mkfn(trm(Combo5.text))
If f$ <> "" Then Command4.Enabled = True

End Sub

Private Sub Combo5_Click()
Dim bfn$, f$, o%, i%

f$ = trm(Combo5.text)
List1.Clear
bfn$ = form1.mydatadir() + "\" + f$
o% = FreeFile
Open bfn$ For Input As #o%
While Not EOF(o%)
  Line Input #o%, f$
  If trm(f$) <> "" Then List1.AddItem f$
Wend
Close #o%
Call clrinfields
Call mkcmd

End Sub

Private Sub Combo5_DropDown()
Dim fn$

fn$ = form1.mydatadir()
Combo5.Clear
If Right$(fn$, 1) <> "\" Then fn$ = fn$ + "\"
fn$ = Dir(fn$ + "*.hsl")
While fn$ <> ""
  If Left$(fn$, 1) <> "." Then Combo5.AddItem fn$
  fn$ = Dir()
Wend
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim add$, adl$, i%, j%

i% = List1.ListIndex
add$ = trm(Combo3.text + " ")
If List1.ListCount > 1 And form1.isfieldmissing("auftritthigru", "opt_kid") Then
  Call MsgBox("A field is missing to combine filters (?and,or?) or a table/column is missing." + vbCrLf + "Please contact support: missing opt_kid in auftritthigru")
  Exit Sub
End If
If List1.ListCount > 1 And add$ = "" Then add$ = "and "
adl$ = Combo1.text + ":" + add$ + ":" + Combo2.text + ":" + Text1.text + ":" + trm(Combo4.text)
If i% >= 0 Then
  List2.Clear
  For j% = 0 To List1.ListCount - 1
    If j% <> i% Then
      List2.AddItem List1.List(j%)
    Else
      List2.AddItem adl$
    End If
  Next j%
  List1.Clear
  For j% = 0 To List2.ListCount - 1
    List1.AddItem List2.List(j%)
  Next j%
  List2.Clear
Else
  List1.AddItem adl$
End If
Call clrinfields
Call mkcmd

End Sub

Private Sub clrinfields()
Combo1.text = ""
Combo2.text = ""
Combo3.text = ""
Combo4.text = ""
Text1.text = "%"

End Sub

Private Sub Command3_Click()
Dim mx%, rrr, c$, na$
Dim r As ADODB.Recordset

Load adrselect
DoEvents
adrselect.Timer1.Enabled = False
DoEvents
On Error Resume Next
Call adrselect.SetFocus
On Error GoTo 0
adrselect.List1.Clear
adrselect.List2.Clear
mx% = Val("0" + trm(adrselect.Text3.text))
c$ = Text2.text + " order by adresse.id"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r.EOF
    adrselect.List1.AddItem trm(r!id)
    r.MoveNext
  Wend
Else
  adrselect.List1.AddItem "Error in SQL-request"
End If

c$ = Text3.text + " order by auftritthigru0.auftrittsid"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r.EOF
    na$ = trm(r!name)
'    Debug.Print trm(r!auftrittsid); " - "; trm(r!vid)
    If InStr(trm(r!auftrittsid), trm(r!vid)) = 1 Then
'      adrselect.List2.AddItem trm(r!id)
      adrselect.List2.AddItem form1.crlffake(na$) + Space$(160) + " (VID:" + r!vid + ") " + "ID:" + r!id
    End If
    r.MoveNext
  Wend
Else
  adrselect.List2.AddItem "Error in SQL-request"
End If
End Sub

Private Sub Command4_Click()
Dim bfn$, f$, o%, i%

f$ = form1.mkfn(cut_d1(trm(Combo5.text), "."))
bfn$ = form1.mydatadir() + "\" + f$ + ".hsl"
o% = FreeFile
Open bfn$ For Output As #o%
For i% = 0 To List1.ListCount - 1
  Print #o%, List1.List(i%)
Next i%
Close #o%

End Sub

Private Sub Form_Load()
Dim cmd$, s As ADODB.Recordset, rrr

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Show
c3nochg = False
cmd$ = "SELECT id From adresstypen order by id"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
While Not s.EOF
  If Left(trm(s!id), 4) <> "rel:" Then
    Combo1.AddItem transe(s!id)
  End If
  s.MoveNext
Wend
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld1
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld1:
On Error GoTo 0

End Sub

Sub mkcmd()
Dim ck$, c$, i%, lcnt%, k%, wh$, t$, r$, FD$, kz$, ad$, f$

Text2.text = "": wh$ = ""
If List1.ListCount < 1 Then Exit Sub
c$ = "select adresse.id from ": lcnt% = 0
ck$ = "select kontakt.id,kontakt.name,kontakt.vid,auftritthigru0.auftrittsid from "
For i% = 0 To List1.ListCount - 1
  If Len(List1.List(i%)) > 4 Then lcnt% = lcnt% + 1
Next i%
k% = lcnt% - 1
For i% = 1 To k%: c$ = c$ + "(": ck$ = ck$ + "(": Next i%
For i% = 0 To List1.ListCount - 1
  t$ = transo(cut_d1(List1.List(i%), ":")): r$ = cut_d2bis(List1.List(i%), ":")
  If r$ = "" Then
    wh$ = wh$ + t$
  Else
    ad$ = cut_d1(r$, ":"): r$ = cut_d2bis(r$, ":")
    f$ = transo(cut_d1(r$, ":")): r$ = cut_d2bis(r$, ":")
    FD$ = cut_d1(r$, ":"): r$ = cut_d2bis(r$, ":")
    kz$ = cut_d1(r$, ":")
    If wh$ <> "" Then wh$ = wh$ + ad$ + " "
    wh$ = wh$ + " (auftritthigru" + trm(i%) + ".auftrittstyp='" + t$ + "' "
    wh$ = wh$ + "and auftritthigru" + trm(i%) + ".FeldName='" + f$ + "' "
    wh$ = wh$ + "and auftritthigru" + trm(i%) + ".FeldDaten like '" + FD$ + "')" + trm(" " + kz$) + " "
    If k% = lcnt% - 1 Then
      c$ = c$ + "adresse inner join "
      ck$ = ck$ + "kontakt inner join "
    End If
    c$ = c$ + "auftritthigru AS auftritthigru" + trm(i%) + " ON adresse.ID = auftritthigru" + trm(i%) + ".auftrittsid"
    If form1.isfieldmissing("auftritthigru", "opt_kid") Then
      ck$ = ck$ + "auftritthigru AS auftritthigru" + trm(i%) + " ON instr(auftritthigru" + trm(i%) + ".auftrittsid,kontakt.ID)>0 "
    Else
      ck$ = ck$ + "auftritthigru AS auftritthigru" + trm(i%) + " ON kontakt.ID = auftritthigru" + trm(i%) + ".opt_kid "
    End If
    If k% > 0 Then
      c$ = c$ + ") inner join "
      ck$ = ck$ + ") inner join "
      k% = k% - 1
    End If
  End If
  c$ = c$ + " ": ck$ = ck$ + " "
Next i%
c$ = c$ + "where " + wh$
ck$ = ck$ + "where " + wh$
Text2.text = c$
Text3.text = ck$
End Sub

Private Sub List1_Click()
Dim r$, i%

i% = List1.ListIndex
If i% < 0 Then Exit Sub
r$ = List1.List(i%)

Combo1.text = transe(cut_d1(r$, ":")): r$ = cut_d2bis(r$, ":")
DoEvents
c3nochg = True
Combo3.text = transe(cut_d1(r$, ":")): r$ = cut_d2bis(r$, ":")
c3nochg = False
Combo2.text = transe(cut_d1(r$, ":")): r$ = cut_d2bis(r$, ":")
Text1.text = cut_d1(r$, ":")
Combo4.text = transe(cut_d2bis(r$, ":"))

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%

i% = List1.ListIndex
If i% < 0 Then Exit Sub
If KeyCode = 46 Or KeyCode = 8 Then
  List1.RemoveItem i%
  Call mkcmd
End If
End Sub

