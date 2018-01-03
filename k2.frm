VERSION 5.00
Begin VB.Form k2 
   Caption         =   "Kalender"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton invrt 
      Caption         =   "inv"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   19
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton selall 
      Caption         =   "alle"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   18
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton invrt 
      Caption         =   "inv"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton selall 
      Caption         =   "alle"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton invrt 
      Caption         =   "inv"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton selall 
      Caption         =   "alle"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "red-raw"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "|"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>>"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<<"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "k2.frx":0000
      Left            =   1800
      List            =   "k2.frx":0010
      TabIndex        =   4
      Text            =   "1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ListBox selct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   2
      Left            =   3000
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox selct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   1
      Left            =   3000
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox selct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   0
      Left            =   3000
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wochen/Zeile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "k2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wrkJet As Workspace
Dim sqla As Database
Dim mnams$(1 To 12)
Dim wdays$(7), selctday
Dim cw%, ch%, dx%, break%
Dim dpl%, fdow%, wkspl%, seldat%, tag0$
Dim xl%(28), yl%(10), olday, yyyy0
Dim lip%
Dim l_typ$(1999), l_bez$(1999), l_X%(1999), l_Y%(1999), l_yy%(1999)
Dim dystart%, rezno%, maxusedy%, dyst%(1999), l_col(1999) As Long

Function getwhere() As String
Dim s$

s$ = ""
For i% = 0 To selct(0).ListCount - 1
  If selct(0).Selected(i%) = True Then
    If Len(s$) = 0 Then
      s$ = "where (([auftrittstyp]='" + selct(0).List(i%) + "') "
    Else
      s$ = s$ + "or ([auftrittstyp]='" + selct(0).List(i%) + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"

getwhere = s$

End Function

Private Sub Combo1_Click()
wkspl% = Combo1.ListIndex + 1
lip% = -1
Call Form_Resize
Call gotodate(tag0$)
End Sub

Private Sub Combo2_Click()

If Combo2.ListIndex < 0 Or Combo3.ListIndex < 0 Then Exit Sub
Cls
nd$ = "1." & (Combo2.ListIndex + 1) & "." & Combo3.List(Combo3.ListIndex)
Call gotodate(nd$)

End Sub

Private Sub Combo3_Click()

If Combo2.ListIndex < 0 Or Combo3.ListIndex < 0 Then Exit Sub

nd$ = "1." & (Combo2.ListIndex + 1) & "." & Combo3.List(Combo3.ListIndex)
Cls
Call gotodate(nd$)


End Sub

Private Sub Command1_Click(Index As Integer)

If Index = 0 Then
  Unload k2
  Exit Sub
End If

End Sub

Private Sub Command2_Click()
i% = Combo3.ListIndex
i% = i% - 1
If i% < 0 Then Exit Sub
Combo3.ListIndex = i%
'Cls
'Call gotodate(nd$)

End Sub

Private Sub Command3_Click()
i% = Combo2.ListIndex
i% = i% - 1
If i% < 0 Then
  i% = 11
  Call Command2_Click
End If
Combo2.ListIndex = i%

End Sub

Private Sub Command4_Click()
i% = Combo2.ListIndex
i% = i% + 1
If i% > Combo2.ListCount - 1 Then
  i% = 0
  Call Command5_Click
End If
Combo2.ListIndex = i%

End Sub

Private Sub Command5_Click()
i% = Combo3.ListIndex
i% = i% + 1
If i% > Combo3.ListCount - 1 Then Exit Sub
Combo3.ListIndex = i%

End Sub

Private Sub Command6_Click()
Combo2.ListIndex = Month(Date) - 1
Combo3.ListIndex = Year(Date) - yyyy0

End Sub

Private Sub Command7_Click()
If Combo1.ListIndex < 0 Then
  Combo1.ListIndex = 0
Else
  Call Combo1_Click
End If

End Sub

Private Sub Form_Load()
Dim t$
wkspl% = 1
seldat% = 0
lip% = -1
rezno% = 1
maxusedy% = 0
dystart% = 300
Me.Top = Form1.mylasttop(Me.Name)
Me.Left = Form1.mylastleft(Me.Name)
Me.Width = Form1.mylastwidth(Me.Name)
Me.Height = Form1.mylastheight(Me.Name)
rezno% = 0
break% = 0
t$ = Form1.myfirstdayofweek()

  AutoRedraw = True
  mnams$(1) = "Januar"
  mnams$(2) = "Februar"
  mnams$(3) = "März"
  mnams$(4) = "April"
  mnams$(5) = "Mai"
  mnams$(6) = "Juni"
  mnams$(7) = "Juli"
  mnams$(8) = "August"
  mnams$(9) = "September"
  mnams$(10) = "Oktober"
  mnams$(11) = "November"
  mnams$(12) = "Dezember"

  wdays$(0) = "Mo"
  wdays$(1) = "Di"
  wdays$(2) = "Mi"
  wdays$(3) = "Do"
  wdays$(4) = "Fr"
  wdays$(5) = "Sa"
  wdays$(6) = "So"
For i% = 0 To 6
  If Left$(t$, 2) = wdays$(i%) Then
    fdow% = i%
    i% = 6
  End If
Next i%

Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
dbpara$ = Form1.getconnstr()
If dbpara$ <> "msaccessmdb" Then
  Set sqla = wrkJet.OpenDatabase(Form1.getdbname(), dbDriverNoPrompt, False, dbpara$)
Else
  Set sqla = wrkJet.OpenDatabase(Form1.getdbname(), False, False)
End If

Show
Call rlist1
Combo2.Clear
Combo2.text = ""
For i% = 1 To 12
  Combo2.AddItem mnams$(i%)
Next i%
Combo3.Clear
Combo3.text = ""
yyyy0 = 1940
For i% = yyyy0 To 2070
  Combo3.AddItem i%
Next i%


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If seldat% = 1 Then
  Form1.setdateselected (selctday)
  Unload k2
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim px%, py%, i%

i% = 0
While xl%(i%) > 0 And xl%(i%) < X: i% = i% + 1: Wend
px% = i%
i% = 0
While yl%(i%) > 0 And yl%(i%) < Y: i% = i% + 1: Wend
py% = i%
selctday = CDate(olday + px% - 1 + wkspl% * 7 * (py% - 1))


End Sub

Private Sub Form_Resize()
Dim col As Long

Cls
If rezno% = 1 Then Exit Sub
Call Form1.setdateselected("")
For i% = 0 To 28: xl%(i%) = 0: Next i%
For i% = 0 To 10: yl%(i%) = 0: Next i%
dpl% = 7 * wkspl%
Cls
If seldat% = 1 Then
  Width = 6330
  Height = 3000
End If
If Width < 6330 Then Width = 6330
If Height < 3000 Then Height = 3000

cw% = Width - 240 - Command1(0).Width
Command1(0).Left = cw% + 40
ct% = Height - 480 - Command1(0).Height
Command1(0).Top = ct%
Combo1.Top = ct%
Combo2.Top = ct%
Combo3.Top = ct%
Combo1.Left = cw% - 40 - Combo1.Width
Label1.Top = ct% + 40
Label1.Left = Combo1.Left - 40 - Label1.Width
Combo3.Left = Label1.Left - 40 - Combo3.Width
Combo2.Left = Combo3.Left - 40 - Combo2.Width

dx% = cw% / (dpl% + 1)
dy% = 0
For n% = 0 To 2
  selct(n%).Width = Command1(0).Width
  selct(n%).Left = cw% + 40
  selct(n%).Top = 120 + n% * (ct% / 3)
  selct(n%).Height = (ct% / 3) * 0.9
  selall(n%).Top = selct(n%).Height + selct(n%).Top + 20
  invrt(n%).Top = selct(n%).Height + selct(n%).Top + 20
  selall(n%).Left = cw% + 40
  invrt(n%).Left = cw% + 80 + selall(n%).Width
Next n%
For n% = 1 To dpl%
  idx% = (fdow% + n% - 1) Mod 7
  If idx% < 6 Then
    col = RGB(222, 222, 222)
  Else
    col = RGB(255, 60, 0)
  End If
  Line ((n + 1) * dx%, 260)-(n * dx%, 120), col, BF
  Print wdays$(idx%)
Next n%

If seldat% = 1 Then
  Call seldatkorrektur
  Exit Sub
Else
  If lip% >= 0 Then Call myredraw
End If
dy% = 300

End Sub

Private Sub Form_Unload(Cancel As Integer)
break% = 1
DoEvents
If seldat% = 0 Then
  Call Form1.setmylasttop(Me.Name, Me.Top)
  Call Form1.setmylastleft(Me.Name, Me.Left)
  Call Form1.setmylastwidth(Me.Name, Me.Width)
  Call Form1.setmylastheight(Me.Name, Me.Height)
End If
Hide

End Sub
Sub rlist1()
Dim rtmp As Recordset

selct(0).Clear

Set rtmp = sqla.OpenRecordset( _
  "SELECT * FROM auftrittstypen order by sortierung", dbOpenDynaset, dbReadOnly)

While Not rtmp.EOF
  selct(0).AddItem rtmp!id
  rtmp.MoveNext
Wend

selct(1).Clear

Set rtmp = sqla.OpenRecordset( _
    "SELECT * FROM adressgruppenindex", dbOpenDynaset, dbReadOnly)

While Not rtmp.EOF
  selct(1).AddItem rtmp!id
  rtmp.MoveNext
Wend

End Sub

Sub selectdate(wann0$)

Cls
tag0$ = wann0$
seldat% = 1
Width = 6330
Height = 4000
Caption = "Wählen Sie ein Datum"
Y = Val(Left(wann0$, 4))
m = Val(Mid$(wann0$, 6, 2))
d = Val(Right$(wann0$, 2))
tdy% = Weekday(CDate(wann0$), vbMonday) - 1
tdy% = (tdy% + wdayno(Form1.myfirstdayofweek())) Mod 7
Combo2.ListIndex = m - 1
Combo3.ListIndex = Y - yyyy0
Call seldatkorrektur

End Sub

Function wdayno(d$) As Integer

For i% = 0 To 6
  If Left(LCase(d$), 2) = LCase(wdays$(i%)) Then
    wdayno = i%
    Exit Function
  End If
Next i%
wdayno = i%

End Function
Sub seldatkorrektur()

For i% = 0 To 2
  selct(i).Enabled = False
Next i%
Command1(0).Enabled = False
Combo1.Enabled = False

If Combo2.ListIndex >= 0 And Combo3.ListIndex >= 0 Then
  fday = CDate("1." & Combo2.ListIndex + 1 & "." & Combo3.ListIndex + yyyy0)
  td0% = Weekday(fday, vbMonday) - 1
  td0% = (td0% + wdayno(Form1.myfirstdayofweek())) Mod 7
  'sday = CDate(fday - td0% - 7)
  sday = CDate(fday - td0%)
  olday = sday
  dy% = (Combo2.Top - 300) / 6
  thsdy = CDate(sday)
  For n = 0 To 41
    If LCase$(Left$(mnams$(Month(thsdy)), 3)) = LCase$(Left$(Combo2.text, 3)) Then
      If tag0$ <> Form1.datum2sql(thsdy) Then
        col = RGB(222, 222, 222)
      Else
        col = RGB(0, 192, 0)
      End If
    Else
      col = RGB(64, 64, 64)
    End If
    py% = Int(n / 7) + 1
    px% = (n Mod 7) + 1
    If xl%(px% - 1) = 0 Then xl%(px% - 1) = px% * dx% + 40
    If yl%(py% - 1) = 0 Then yl%(py% - 1) = dy% * py% - 20
    Line ((px% + 1) * dx% - 40, 210 + dy% * py%)-(px% * dx% + 40, dy% * py% - 20), col, BF
    Print Left(thsdy, 2);
    thsdy = CDate(thsdy + 1)
  Next n
End If
End Sub


Sub myredraw()
dystart% = 0
For i% = 0 To lip%
  Call plotme(i%)
Next i%

End Sub

Private Sub selct_Click(Index As Integer)
Call Command7_Click
End Sub
Public Sub setbreak()
break% = 1
End Sub

Sub gotodate(tag$)
Dim r As Recordset



tag0$ = Form1.datum2sql(tag$)
If seldat% = 0 Then
  Font.Bold = False
  Font.Name = "Small Fonts"
  Font.Size = 7
End If
dystart% = 300
maxusedy% = 0
lip% = -1
Y = Val(Left(tag0$, 4))
m = Val(Mid$(tag0$, 6, 2))
d = Val(Right$(tag0$, 2))
tdy% = Weekday(CDate(tag0$), vbMonday) - 1
tdy% = (tdy% + wdayno(Form1.myfirstdayofweek())) Mod 7
Combo2.ListIndex = m - 1
Combo3.ListIndex = Y - yyyy0
sday = CDate(tag$) - tdy%
olday = sday
dy% = (Combo2.Top - 300) / 6

thsdy = CDate(sday)
Do

dv$ = Form1.datum2sql(thsdy)
If Combo1.ListIndex >= 0 Then
  b_y% = Combo1.ListIndex + 1
Else
  b_y% = 1
End If
db$ = Form1.datum2sql(thsdy + b_y% * 7 - 1)

selstr$ = ""
selstr$ = selstr$ + "(([datum]>='" + dv$ + "' and [datum]<='" + db$ + "')) "
'Debug.Print selstr$
If selct(2).ListCount = 0 Then
  gw$ = getwhere()
  cmd$ = "select id as aid,datum as adatum, bezeichnung as abez, auftrittstyp as atyp from auftritt "
  If gw$ = "" Then
    gw$ = "where "
  Else
    gw$ = gw$ + " and "
  End If
  
  cmd$ = cmd$ + gw$ + selstr$
Else
  cmd$ = "SELECT auftritt.id as aid,auftritt.datum as adatum,auftritt.bezeichnung as abez, auftritthigru.auftrittstyp as atyp, auftritthigru.FeldName "
  cmd$ = cmd$ + "FROM auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid "
  cmd$ = cmd$ + "Where "
  nosel = 1
  For i% = 0 To selct(2).ListCount - 1
    If selct(2).Selected(i%) = True Then
      nosel = 0
      fsel$ = selct(2).List(i%)
      i% = selct(2).ListCount
    End If
  Next i%
  bisi% = 30
  If nosel = 1 Then
    kid$ = selct(2).List(0)
    'kid$ = "((InStr([FeldDaten], '" + kid$ + "')) <> '0') "
    kid$ = "([FeldDaten] like '" + kid$ + "*') "
    For i% = 1 To selct(2).ListCount - 1
'      kid$ = kid$ + "or ((InStr([FeldDaten], '" + selct(2).List(i%) + "')) <> '0') "
      kid$ = kid$ + "or ([FeldDaten] like '" + selct(2).List(i%) + "*') "
      bisi% = bisi% - 1
      If bisi% < 0 Then i% = selct(2).ListCount - 1
    Next i%
  Else
    kid$ = fsel$
    'kid$ = "((InStr([FeldDaten], '" + kid$ + "')) <> '0') "
    kid$ = "([FeldDaten] like '" + kid$ + "*') "
    bisi% = selct(2).ListCount - 1: If bisi% > 20 Then bisi% = 20
    For i% = 0 To bisi%
      If selct(2).Selected(i%) = True And selct(2).List(i%) <> fsel$ Then
        kid$ = kid$ + "or ([FeldDaten] like '" + selct(2).List(i%) + "*') "
        bisi% = bisi% - 1
        If bisi% < 0 Then i% = selct(2).ListCount - 1
      End If
    Next i%
  End If
'  Debug.Print kid$
  cmd$ = cmd$ + " ( " + kid$ + ") and  "
  cmd$ = cmd$ + selstr$
End If
cmd$ = cmd$ + " ORDER BY auftritt.datum,auftritt.id"
Debug.Print cmd$

'tage malen
d0 = CDate(dv): xoff% = 0
While d0 <= CDate(db)
    lip% = lip% + 1
    l_col(lip%) = RGB(255, 255, 255)
    If CDate(d0) = Date Then l_col(lip%) = RGB(0, 255, 0)
    l_typ$(lip%) = "SYSTEM_DATE"
    l_bez$(lip%) = CDate(d0)
    l_X%(lip%) = xoff%
    l_Y%(lip%) = 0
    l_yy%(lip%) = 0
    dyst%(lip%) = dystart%
    Call plotme(lip%)
  d0 = CDate(d0 + 1)
  xoff% = xoff% + 1
  DoEvents
  If break% <> 0 Then Exit Sub
Wend
'daten selektieren
Set r = sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
prvid$ = "-10"
ltg = ""
b_y% = 0: b_ymax% = 0
yoff% = 0
While Not r.EOF And break% = 0
'  Debug.Print r!atyp, r!adatum, CDate(r!adatum) - thsdy, r!abez, r!aid
  
  If ltg <> r!adatum Then
    xoff% = CDate(r!adatum) - thsdy
    ltg = r!adatum
    If b_ymax% < b_y% Then b_ymax% = b_y%
    b_y% = 1
  Else
    b_y% = b_y% + 1
  End If
  If r!aid <> prvid$ Then        ' bitte jeder nur ein kreuz
    prvid$ = r!aid
    lip% = lip% + 1
    If Not IsNull(r!atyp) Then
      l_col(lip%) = Form1.get_eventcolor(r!atyp)
    Else
      l_col(lip%) = RGB(255, 0, 0)
    End If
    l_typ$(lip%) = "" & r!atyp
    l_bez$(lip%) = r!abez
    l_X%(lip%) = xoff%
    l_Y%(lip%) = yoff%
    l_yy%(lip%) = b_y%
    dyst%(lip%) = dystart%
    Call plotme(lip%)
  End If
  DoEvents
  r.MoveNext
Wend
r.Close
'rahmen malen
If break% <> 0 Then Exit Sub
d0 = CDate(dv): xoff% = 0
While d0 <= CDate(db)
    lip% = lip% + 1
    l_col(lip%) = RGB(0, 0, 0)
    If CDate(d0) = Date Then l_col(lip%) = RGB(0, 255, 0)
    l_typ$(lip%) = "SYSTEM_FRAME"
    l_X%(lip%) = xoff%
    l_Y%(lip%) = b_ymax%
    l_yy%(lip%) = 0
    dyst%(lip%) = dystart%
    Call plotme(lip%)
  d0 = CDate(d0 + 1)
  xoff% = xoff% + 1
  DoEvents
Wend


thsdy = CDate(db$) + 1
dystart% = maxusedy% + 160

Loop Until dystart% > Int(0.8 * Combo2.Top)
MousePointer = 0
Call Form_Resize

End Sub

Sub plotme(li%)
Dim yh%

yh% = 280
typ$ = l_typ$(li%)
bez$ = l_bez$(li%)
xoff% = l_X%(li%)
yoff% = l_Y%(li%)
yy% = l_yy%(li%)
col = l_col(li%)
dyadd% = dyst%(li)

'Debug.Print typ$, bez$, xoff%, yoff%, yy%
yyoff% = yy% * yh%

'Debug.Print , , (xoff% + 1) * dx%, dystart%, yyoff%

curry% = dyadd% + yh% + yyoff%
If typ$ = "SYSTEM_FRAME" Then
'  col = RGB(0, 0, 0)
  Line ((xoff% + 2) * dx% - 60, dyadd% + yh% * (yoff% + 1))-((xoff% + 1) * dx%, dyadd%), col, B
Else
  Line ((xoff% + 2) * dx% - 60, curry% - 60)-((xoff% + 1) * dx%, dyadd% + yyoff%), col, BF
  Print bez$
End If
If maxusedy% < curry% Then
  maxusedy% = curry%
End If


End Sub
