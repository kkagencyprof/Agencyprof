VERSION 5.00
Begin VB.Form kurse 
   Caption         =   "Kurse"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   3285
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command9 
      Caption         =   "&Test"
      Height          =   375
      Left            =   960
      TabIndex        =   20
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   2160
      TabIndex        =   18
      Top             =   600
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton exme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "kurse.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Kalender schliessen"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton allclose 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   480
      Picture         =   "kurse.frx":0250
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Schliesst alle Formulare"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton refr 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      Picture         =   "kurse.frx":087A
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   2175
      Left            =   5280
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   0
      ToolTipText     =   "huhu"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "von"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   16
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "bis"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   3615
   End
End
Attribute VB_Name = "kurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p1mm As Boolean
Dim mdays(1 To 12)

Private Sub allclose_Click()

Call form1.Form_DblClick

End Sub

Private Sub Command1_Click()
Dim M%, Y%, l$, p%, t$

t$ = Text1(0).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
M% = M% - 1
If M% < 1 Then
  M% = 12
  Y% = Y% - 1
End If
Text1(0).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click
End Sub

Private Sub Command2_Click()
Dim M%, Y%, l$, p%, t$


t$ = Text1(0).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
Y% = Y% - 1
Text1(0).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command3_Click()
Dim M%, Y%, l$, p%, t$

t$ = Text1(1).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
M% = M% + 1
If M% > 12 Then
  M% = 1
  Y% = Y% + 1
End If
Text1(1).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command4_Click()
Dim M%, Y%, l$, p%, t$


t$ = Text1(1).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
Y% = Y% + 1
Text1(1).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command5_Click()
Dim M%, Y%, l$, p%, t$

t$ = Text1(0).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
Y% = Y% + 1
Text1(0).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click


End Sub

Private Sub Command6_Click()
Dim M%, Y%, l$, p%, t$

t$ = Text1(0).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
M% = M% + 1
If M% > 12 Then
  M% = 1
  Y% = Y% + 1
End If
Text1(0).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command7_Click()
Dim M%, Y%, l$, p%, t$

t$ = Text1(1).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
M% = M% - 1
If M% < 1 Then
  M% = 12
  Y% = Y% - 1
End If
Text1(1).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command8_Click()
Dim M%, Y%, t$, l$, p%

t$ = Text1(1).text: p% = InStr(t$, "."): l$ = Left$(t$, p%): t$ = Mid$(t$, p% + 1)
M% = Val(t$)
Y% = Val(Mid$(t$, InStr(t$, ".") + 1))
Y% = Y% - 1
Text1(1).text = l$ & trm(M%) & "." & trm(Y%)
Call refr_Click

End Sub

Private Sub Command9_Click()
Dim i%, brk As Boolean

i% = List1.ListIndex
Do
  i% = i% + 1
  If i% < List1.ListCount Then
    List1.ListIndex = i%
    DoEvents
    If List2.ListCount < 2 Then brk = True
  Else
    brk = True
  End If
Loop Until brk
List1.SetFocus
End Sub

Private Sub exme_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim M%, d%, Y%

mdays(1) = 31
mdays(2) = 28
mdays(3) = 31
mdays(4) = 30
mdays(5) = 31
mdays(6) = 30
mdays(7) = 31
mdays(8) = 31
mdays(9) = 30
mdays(10) = 31
mdays(11) = 30
mdays(12) = 31
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Width = form1.mylastwidth(Me.name, 1)
Me.Height = form1.mylastheight(Me.name, 1)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Top = Me.Height / 3
  Me.Left = Me.Width / 3
End If
Call form1.formpos(Me)
Y% = apyear(Now())
M% = apmonth(Now()) - 3
If M% < 1 Then
  M% = M% + 12
  Y% = Y% - 1
End If
d% = apday(Now())
Text1(0).text = trm(d%) & "." & trm(M%) & "." & trm(Y%)
Text1(1).text = trm(d%) & "." & apmonth(Now()) & "." & apyear(Now())
p1.AutoRedraw = True
Show
Call rlist1

End Sub

Private Sub Form_Resize()

If Width < 7000 Then Width = 7000
If Height < 2000 Then Height = 2000
exme.Top = Me.Height - exme.Height - 400
Command9.Top = exme.Top
allclose.Top = exme.Top
Label2.Top = exme.Top
p1.Width = Me.Width - 160 - p1.Left
p1.Height = exme.Top - p1.Top
List1.Height = exme.Top - List1.Top
List2.Height = exme.Top - List2.Top
p1.ScaleHeight = 1000
p1.ScaleWidth = 1000

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

Private Sub List1_Click()
Dim i%

List2.Clear
Label2.Caption = ""
i% = List1.ListIndex
If i% < 0 Then Exit Sub

Label2.Caption = List1.List(i%)
Label3.Caption = "kurse"

Call rlist2
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, sq$, r As Recordset
Dim c$, tbl$, sl$, dtg$, ans%, tbl1$, tbl2$

If KeyCode = 8 Or KeyCode = 46 Then
  tbl$ = Label3.Caption
  sl$ = "wid": tbl2$ = "kurse"
  tbl1$ = "waehrung"
  ans% = vbYes
  If List2.ListCount > 1 Then
    ans% = MsgBox(transe("Wirklich ") & trm(List2.ListCount) & transe(" Sätze löschen?"), vbYesNo + vbCritical + vbDefaultButton2, List2.ListCount & " Sätze löschen?")
  End If
  If ans% = vbYes Then
    c$ = "delete from " & tbl2$ & " where id='" & Label2.Caption & "'"
    Call form1.sqlqry(c$)
    c$ = "delete from " & tbl1$ & " where " & sl$ & "='" & Label2.Caption & "'"
    Call form1.sqlqry(c$)
    List2.Clear
    ans% = List1.ListIndex
    List1.RemoveItem ans%
    Label2.Caption = ""
    Label3.Caption = ""
    DoEvents
    List1.ListIndex = imin(ans%, List1.ListCount - 1)
  End If
End If
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, sq$, r As Recordset
Dim c$, tbl$, sl$, dtg$, ans%, tbl1$, tbl2$

idx% = List2.ListIndex
If idx% < 0 Then Exit Sub
If KeyCode = 8 Or KeyCode = 46 Then
  tbl$ = Label3.Caption
  sl$ = "wid": tbl2$ = "kurse"
  tbl1$ = "waehrung"
  id$ = strrepl(List2.List(idx%), ":", "")
  id$ = word1(id$)
  id$ = id$ & Label2.Caption
  ans% = MsgBox(id$ & ": " + transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, List2.ListCount & transe(" löschen?"))
  If ans% = vbYes Then
    c$ = "delete from " & tbl$ & " where id='" & id$ & "'"
    Call form1.sqlqry(c$)
    List2.RemoveItem idx%
    Call rlist2
  End If
End If
End Sub

Private Sub refr_Click()

Call rlist2

End Sub

Sub rlist1()
Dim c$, r As Recordset

List1.Clear
List2.Clear
Label2.Caption = ""
c$ = "select * from waehrung order by id"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  List1.AddItem r!id
  r.MoveNext
Wend
End Sub

Sub rlist2()
Dim c$, r As Recordset, tbl$, sl$, dtg$, i%, cl$, rrr
Dim sdt$, edt$, whr$, anzmd%, dmd%, rku As Double
Dim minx As Double, maxx As Double, miny As Double, maxy As Double
Dim scx, scy, x0 As Double, xe As Double, px0, py0, px1, py1

sdt$ = datum2sql(Text1(0).text)
On Error Resume Next
edt$ = datum2sql(Text1(1).text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
If IsDate(sdt$) = False Then Exit Sub
If IsDate(edt$) = False Then Exit Sub
anzmd% = (apyear(edt$) - apyear(sdt$)) * 12
anzmd% = (anzmd% + apmonth(edt$) - apmonth(sdt$)) + 1
If anzmd% <= 0 Then Exit Sub
whr$ = " and ( " + _
     "(id>='" & sdt$ & "')  and (id<='" & edt$ & "') " + _
     ") "

tbl$ = Label3.Caption
sl$ = "wid"
If tbl$ = "" Then Exit Sub
List2.Clear
p1.Cls
minx = 9999999: miny = 9999999: maxx = 0: maxy = 0
c$ = "select id," & sl$ & " as name,kurs from " & tbl$ & " where " & sl$ & "='" & Label2.Caption & "'" & whr$ & "order by id"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If sl$ = "wid" Then
    dtg$ = Left(r!id, 10)
  Else
    dtg$ = Left(r!id, 16)
  End If
  On Error Resume Next
  rku = var2dbl(r!kurs)
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    List2.AddItem dtg$ & ": " & r!kurs
    If rku < miny Then miny = rku
    If rku > maxy Then maxy = rku
  End If
  r.MoveNext
Wend
miny = 0.95 * miny
maxy = 1.05 * maxy
rku = maxy - miny
If rku < 0.0001 Then rku = 0.0001
scy = p1.ScaleHeight / rku
c$ = List2.List(0)
If c$ = "" Then Exit Sub

x0 = CDate(Text1(0).text)

xe = CDate(Text1(1).text)
px0 = 0
c$ = List2.List(0)
c$ = Mid$(c$, InStr(c$, ":") + 2)
py0 = (var2dbl(strrepl(c$, ";", ".")) - miny) * scy
rku = (xe - x0)
If rku = 0 Then rku = p1.ScaleWidth
scx = p1.ScaleWidth / rku
c$ = List2.List(0)
c$ = Left$(c$, InStr(c$, ":") - 1)
px0 = (CDate(c$) - x0) * scx
For i% = 1 To List2.ListCount - 1
  cl$ = List2.List(i%)
  c$ = cl$
  If c$ = "" Then Exit Sub
  c$ = Left$(c$, InStr(c$, ":") - 1)
  px1 = (CDate(c$) - x0) * scx
  c$ = cl$
  c$ = Mid$(c$, InStr(c$, ":") + 2)
  py1 = (var2dbl(strrepl(c$, ";", ".")) - miny) * scy
  On Error Resume Next
  p1.Line (px0, p1.ScaleHeight - py0)-(px1, p1.ScaleHeight - py1)
  On Error GoTo 0
  px0 = px1
  py0 = py1
Next i%

End Sub
