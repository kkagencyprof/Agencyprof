VERSION 5.00
Begin VB.Form k3 
   Caption         =   "Kalender"
   ClientHeight    =   3585
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton zoomout 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   54
      ToolTipText     =   "zoom out"
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton zoomin 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      ToolTipText     =   "zoom in"
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton up12 
      Caption         =   "<<"
      Height          =   255
      Left            =   1560
      TabIndex        =   52
      ToolTipText     =   "1 Jahr zurück"
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton dwn12 
      Caption         =   ">>"
      Height          =   255
      Left            =   3720
      TabIndex        =   51
      ToolTipText     =   "1 Jahr vor"
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton up4 
      Caption         =   "<"
      Height          =   255
      Left            =   2280
      TabIndex        =   50
      ToolTipText     =   "4 Wochen zurück"
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton dwn4 
      Caption         =   ">"
      Height          =   255
      Left            =   3000
      TabIndex        =   49
      ToolTipText     =   "4 Wochen vor"
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   41
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   41
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   40
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   40
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   39
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   39
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   38
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   38
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   37
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   37
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   36
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   35
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   34
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   33
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   32
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   31
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   30
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   29
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   28
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   28
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   27
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   26
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   25
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   24
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   23
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   22
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   21
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   20
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   19
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   18
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   17
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   16
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   15
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   14
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   13
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   12
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   11
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   10
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   9
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   8
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   7
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   6
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   48
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   47
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   46
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   45
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   44
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   43
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu neu 
      Caption         =   "Neues Ereignis"
      Visible         =   0   'False
      Begin VB.Menu pred 
         Caption         =   "öffne Projekt"
      End
      Begin VB.Menu trmed 
         Caption         =   "öffne Termin"
      End
      Begin VB.Menu daycal 
         Caption         =   "Tageskalender"
      End
      Begin VB.Menu tmrstrt 
         Caption         =   "Timer starten"
      End
      Begin VB.Menu ruler 
         Caption         =   "-----------------"
      End
      Begin VB.Menu project 
         Caption         =   "Neues Projekt"
      End
      Begin VB.Menu termin 
         Caption         =   "Neuer Termin"
      End
      Begin VB.Menu todo 
         Caption         =   "Neues Todo"
      End
   End
End
Attribute VB_Name = "k3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mnams$(1 To 12), wdays$(7), wkspl%, tag0$, fdow%, olday$
Dim blk As Long, prvi%, break%, mmblock As Boolean, prvtt$
Dim lip%(41), lipv%, nsel%, zoom As Integer
Dim c_typ$(41, 299), c_bez$(41, 299), c_stat(41, 299) As Long, c_id$(41, 299)
Dim c_col(41, 299) As Long, adrgrpselcache$, adrgrpnoselcache$
Dim nogoto%, ypixperentry%, k3date As String, nalert As String, nalertcap As String

Function adrisinselectedgroup(i$, selstr$) As Boolean
Dim r As ADODB.Recordset, cmd$, rrr

Dim d2infile As String, d2insub As String
d2infile = "k3": d2insub = "adrisinselectedgroup"
adrisinselectedgroup = False
If InStr(adrgrpselcache$, "|" & i$ & "|") > 0 Then
  adrisinselectedgroup = True
  GoTo exfu
End If
If InStr(adrgrpnoselcache$, "|" & i$ & "|") > 0 Then
  adrisinselectedgroup = False
  GoTo exfu
End If
cmd$ = "select grpid from adressgruppen where adressid='" & i$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  While Not r.EOF
    If InStr(selstr$, "|" & r!grpid & "|") > 0 Then
      adrisinselectedgroup = True
      adrgrpselcache$ = adrgrpselcache$ & i$ & "|"
      GoTo exfu
    End If
    r.MoveNext
  Wend
  adrisinselectedgroup = False
  adrgrpnoselcache$ = adrgrpnoselcache$ & i$ & "|"
  GoTo exfu
Else
  adrisinselectedgroup = False
  adrgrpnoselcache$ = adrgrpnoselcache$ & i$ & "|"
  GoTo exfu
End If
adrisinselectedgroup = True
exfu:
End Function

Public Sub setnogoto(w%)
'd2infile = "k3": d2insub = "setnogoto"
nogoto% = w%
End Sub
Public Sub settag0(d$)
Dim i%

'd2infile = "k3": d2insub = "settag0"
'For i% = 0 To 41
'  lip%(i%) = -1
'  p1(i%).Font.Name = "Small Fonts"
'  p1(i%).Font.Size = 7
'Next i%
lipv% = 0
tag0$ = d$
Call Form_Resize

End Sub

Private Sub daycal_Click()

'd2infile = "k3": d2insub = "daycal_Click"
Load dayvw
On Error Resume Next
Call dayvw.SetFocus
On Error GoTo 0
dayvw.Text1.text = k3date
End Sub

Private Sub dwn12_Click()
Call kc.Command6_Click
End Sub

Private Sub dwn4_Click()
Call kc.Command5_Click
End Sub

Private Sub Form_Load()
Dim i%, zc$, rrr

'd2infile = "k3": d2insub = "Form_Load"
mmblock = False
k3date = ""
For i% = 0 To 41: lip%(i%) = -1: Next i%
lipv% = 0
nalert = ""
nalertcap = ""
nogoto% = 1
blk = RGB(0, 0, 0)
ypixperentry% = 16
prvi% = -1
break% = 0
nsel% = 1

zc$ = form1.getusersetting("zkalzoom", "1")
On Error Resume Next
zoom = Int(zc$)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then zoom = 1
If Not form1.weckerpresent Then tmrstrt.Enabled = False
neu.Caption = transe("Neues Ereignis")
pred.Caption = transe("Öffne Projekt")
project.Caption = transe("Neues Projekt")
trmed.Caption = transe("Öffne Termin")
termin.Caption = transe("Neuer Termin")
daycal.Caption = transe("Tageskalender")
todo.Caption = transe("Neues Todo")
mnams$(1) = form1.inmylanguage("Januar")
mnams$(2) = form1.inmylanguage("Februar")
mnams$(3) = form1.inmylanguage("März")
mnams$(4) = form1.inmylanguage("April")
mnams$(5) = form1.inmylanguage("Mai")
mnams$(6) = form1.inmylanguage("Juni")
mnams$(7) = form1.inmylanguage("Juli")
mnams$(8) = form1.inmylanguage("August")
mnams$(9) = form1.inmylanguage("September")
mnams$(10) = form1.inmylanguage("Oktober")
mnams$(11) = form1.inmylanguage("November")
mnams$(12) = form1.inmylanguage("Dezember")

wdays$(0) = form1.inmylanguage("Mo")
wdays$(1) = form1.inmylanguage("Di")
wdays$(2) = form1.inmylanguage("Mi")
wdays$(3) = form1.inmylanguage("Do")
wdays$(4) = form1.inmylanguage("Fr")
wdays$(5) = form1.inmylanguage("Sa")
wdays$(6) = form1.inmylanguage("So")

Me.Caption = form1.inmylanguage("Kalender")
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Width = form1.mylastwidth(Me.name, 1)
Me.Height = form1.mylastheight(Me.name, 1)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Width = Int(1.3 * Me.Width)
  Me.Top = Me.Height / 3
  Me.Left = Me.Width / 3
End If
Call form1.formpos(Me)
form1.kalopen = True
Load ttform
Show
nogoto% = 0
End Sub

Private Sub Form_Resize()

'd2infile = "k3": d2insub = "Form_Resize"
If Height < 4000 Then
  Height = 4000
End If
ScaleWidth = 700
ScaleHeight = 600 + Label1(0).Height
Cls
Font.Size = 10
ForeColor = RGB(22, 22, 22)
Call form1.dbg2f("k3 calling gotoday:" + tag0$)
If tag0$ <> "" And nogoto% = 0 Then
  Call gotoday(tag0$)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "k3": d2insub = "Form_Unload"
form1.kalopen = False
zoom = 1
Hide
'Unload kc
Unload ttform
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastwidth(Me.name, Me.Width)
Call form1.setmylastheight(Me.name, Me.Height)
exuld:
On Error GoTo 0
End Sub

Sub gotoday(d$)
Dim i%, X%, Y%, col As Long, idx%, d0 As Variant, dn As Variant
Dim wd%, dx As Variant, t$

t$ = form1.inmylanguage(form1.myfirstdayofweek())
For i% = 0 To 6
  If Left$(t$, 2) = wdays$(i%) Then
    fdow% = i%
    i% = 6
  End If
Next i%
col = form1.BackColor
For i% = 0 To 6
  idx% = (fdow% + i%) Mod 7
  Label1(idx%).BackColor = RGB(222, 222, 222)
  Label1(i%).Caption = wdays$(idx%)
  If form1.outmylanguage(Label1(i%).Caption) = "So" Then
    Label1(i%).BackColor = RGB(255, 60, 0)
  End If
  Label1(i%).Width = 100
  Label1(i%).Top = 0
  Label1(i%).Left = i% * 100
Next i%
d0 = CDate(tag0$) - 7
While Weekday(CDate(d0), vbMonday) - 1 <> fdow%
  d0 = d0 - 1
Wend
olday$ = d0
dn = d0
dx = CDate(tag0$)
For i% = 0 To 41
  X% = (i% Mod 7) * 100
  Y% = Int(i% / 7) * 100 * zoom + Label1(0).Height
  p1(i%).Width = 100
  p1(i%).Height = 100 * zoom
  p1(i%).Top = Y%
  p1(i%).Left = X%
  p1(i%).BackColor = col
  p1(i%).ScaleHeight = 100 * zoom
  p1(i%).ScaleWidth = 100
  p1(i%).ToolTipText = ""
  If dn = Date Then p1(i%).BackColor = RGB(0, 255, 0)
  If dn = dx Then
    p1(i%).ForeColor = RGB(0, 0, 255)
    If dn <> Date Then p1(i%).BackColor = RGB(255, 255, 0)
  Else
    p1(i%).ForeColor = RGB(0, 0, 0)
  End If
  If Left(dn, 2) = "01" Then
    p1(i%).Font.Bold = True
  Else
    p1(i%).Font.Bold = False
  End If
  p1(i%).AutoRedraw = True
  p1(i%).Cls
  p1(i%).Print dn
  dn = dn + 1
Next i%

If lipv% = 0 Then
  Call rdsels
Else
  Call c_draw
End If

End Sub

Private Sub p1_DblClick(Index As Integer)
Dim d As Variant, id$

'd2infile = "k3": d2insub = "p1_DblClick"
If nsel% = 0 Then
  Call kc.List1_DblClick
End If
End Sub

Private Sub p1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim o%, rrr

If Button = 2 Then
  tmrstrt.Enabled = False
  If nalert <> "" Then
    If form1.weckerpresent Then
      o% = FreeFile
      On Error Resume Next
      Open form1.s00dir() + "\wecker.ini" For Output As #o%
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        Print #o%, nalert
        Print #o%, nalertcap
        Close #o%
        If form1.weckerpresent Then tmrstrt.Enabled = True
      Else
        form1.weckerpresent = False
      End If
    End If
  End If
  PopupMenu neu
End If

End Sub

Private Sub p1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Variant, i%, dy%, yp%, tt$, r As ADODB.Recordset, rrr, ta$, j%, rest$
Dim rtmp As ADODB.Recordset, d2 As Variant, d0 As Variant
Dim ctrl As Control, beginn$, tb$, fx As Double, fy As Double, pfx As Double, pfy As Double
Dim ttfT, ttfL, wd As Integer, ttedlm As String, ttlines As Integer
Dim d2infile As String, d2insub As String, ttdelm As String
Dim bmx%, bmy%, wdw As Long, extt As Boolean

d2infile = "k3": d2insub = "p1_MouseMove"
If olday$ = "" Or form1.poplock Then Exit Sub
If mmblock Then Exit Sub
extt = False
tt$ = LCase(form1.getusersetting("extendedcalendartooltips", "no"))
If tt$ = "ja" Or tt$ = "yes" Then extt = True
mmblock = True
d0 = now
If form1.ttmode = "1" Then
  ttdelm = ", "
Else
  ttdelm = vbCrLf
End If
If prvi% <> Index Then
  prvi% = Index
  d = CDate(olday$)
  k3date = CDate(d + prvi%)
  k3.Caption = form1.inmylanguage("Kalender") & " (" & wdays(Weekday(CDate(d + prvi%), vbMonday) - 1) & ", " & k3date & ")"
  kc.List1.Clear
  For i% = 0 To lip%(prvi%)
    kc.List1.AddItem transe(c_typ$(prvi%, i%)) & " " & c_bez$(prvi%, i%) & Space$(80) & " (AID:" & c_id$(prvi%, i%)
  Next i%
End If
nsel% = 1
If kc.List1.ListCount <= 0 Then
  ttform.Hide
  DoEvents
Else
  yp% = Y - 10
  dy% = ypixperentry%
  'dy% = p1(prvi%).Height / kc.List1.ListCount
  dy% = Int(yp% / dy%)
  If dy% > kc.List1.ListCount - 1 Then dy% = kc.List1.ListCount - 1
  If dy% >= 0 Then
    kc.List1.ListIndex = dy%
    tt$ = kc.List1.List(kc.List1.ListIndex)
    tt$ = Mid$(tt$, InStr(tt$, "(AID:") + 5)
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rrr = form1.adoopen(rtmp, "SELECT * FROM auftritt where id= '" & tt$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    'If rrr = 0 Then
    If Not rtmp.EOF Then
      On Error Resume Next
      beginn$ = trm(rtmp!zeit)
      d2 = CDate(datfromsql(trm(rtmp!datum)) + " " + beginn$)
      nalert = datfromsql(trm(rtmp!datum)) + " " + beginn$
      nalertcap = trm(rtmp!bezeichnung)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        d2 = now
        nalert = ""
        nalertcap = ""
      End If
      If d2 <= d0 Then
        nalert = ""
        nalertcap = ""
      End If
      If trm(rtmp!auftrittstyp) = "" Then
        rrr = -1
      Else
        tt$ = "select * from usr_" & utabn(rtmp!auftrittstyp) & " where id='" & tt$ & "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
        rrr = form1.adoopen(r, tt$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      End If
      If rrr = 0 Then
        If Not r.EOF Then
          tt$ = "": wd = 0
          For i% = 1 To r.Fields.count - 1
            On Error Resume Next
            ta$ = trm(r.Fields(i%).value)
            tb$ = trm(r.Fields(i%).name)
            rrr = Err
            On Error GoTo 0
            If rrr <> 0 Then ta$ = ""
            If ta$ <> "" Then
              If tt$ <> "" Then
                tt$ = tt$ & ttdelm
              End If
'              tt$ = tt$ & tb$ + ": " + ta$
              If ttdelm = vbCrLf Then
                tt$ = tt$ + tb$ + ": " + ta$
                rest$ = ta$
                If InStr(rest$, vbCrLf) = 0 Then
                  If Len(tb$ + ": " + ta$) > wd Then wd = Len(tb$ + ": " + ta$)
                Else
                  Do
                    j% = InStr(rest$, vbCrLf)
                    If j% > 0 And j% < Len(rest$) Then
                      If j > wd Then wd = j%
                      rest$ = Mid(rest$, j% + 1)
                    Else
                      j% = 0
                    End If
                  Loop Until j% = 0
                End If
              Else
                tt$ = tt$ + ta$
              End If
            End If
            If InStr(tb$, "Programm") = 1 And trm(ta$) <> "" And extt Then
              tt$ = tt$ + vbCrLf + strrepl(form1.getwerke(ta$), vbCrLf, vbCrLf + "  ")
            End If
            ta$ = ""
            If (Not extt) And Len(tt$) > 200 And ttdelm$ <> vbCrLf Then Exit For
          Next i%
'          While InStr(tt$, vbCrLf) > 0: tt$ = strrepl(tt$, vbCrLf, " - "): Wend
          ta$ = ""
          If tmrstrt.Enabled Then
            On Error Resume Next
            ta$ = fixeur(d2 - d0) + " Tg."
            rrr = Err
            On Error GoTo 0
            If rrr <> 0 Then ta$ = ""
          End If
          If ta$ <> "" Then
            If tt$ <> "" Then
              If ttdelm$ <> "" And ttdelm$ <> vbCrLf Then tt$ = tt$ & ttdelm$
            End If
            If tt$ <> "" And Right(tt$, 1) <> vbCrLf Then tt$ = tt$ + vbCrLf
            tt$ = tt$ + ta$
          End If
'          p1(Index).ToolTipText = tt$ + " " + ta$
        End If
      Else
        tt$ = ""
      End If
    End If
    nsel% = 0
  End If
  If ttdelm = vbCrLf Then
    If beginn$ <> "" Then
      tt$ = beginn$ + vbCrLf + tt$
    End If
    If tt$ = "" Then
      ttform.Hide
      prvtt$ = ""
    Else
      fx = k3.Width / k3.ScaleWidth
      fy = (k3.Height - 450) / k3.ScaleHeight
      pfx = p1(Index).Width / p1(Index).ScaleWidth
      pfy = p1(Index).Height / p1(Index).ScaleHeight
      bmx% = Screen.Width / 2
      bmy% = Screen.Height / 2
      ttfT = Y * pfy * fy + k3.Top + 450 + p1(Index).Top * fy + 400
      ttfL = X * pfx * fx + k3.Left + p1(Index).Left * fx + 200
      wdw = wd: If wdw > 100 Then wdw = 100
      ttform.Width = wdw * 100
      ttform.Text1.Width = ttform.Width
      If ttfT > bmy% Then
        ttfT = (Y * pfy * fy + k3.Top + (p1(Index).Top * fy + 400)) - ttform.Height
      End If
      If ttfL > bmx% Then
        ttfL = (X * pfx * fx + k3.Left + p1(Index).Left * fx - 200) - ttform.Width
      End If
'      If ttfL + ttform.Width > Screen.Width Then
'        ttfL = Screen.Width - ttform.Width
'      End If
'      If ttfT + ttform.Height > Screen.Height Then
'        ttfT = Screen.Height - ttform.Height
'      End If
      ttform.Top = ttfT
      ttform.Left = ttfL
      If prvtt$ <> tt$ Then
        prvtt$ = tt$
        ttlines = linesof(tt$) + 1
        If ttlines > 160 Then ttlines = 160
        ttform.Hide
        DoEvents
        ttform.Text1.text = tt$
        ttform.Height = 200 * ttlines
        ttform.Text1.Height = ttform.Height
        ttform.Show
      End If
    End If
  Else
    p1(Index).ToolTipText = strrepl(tt$, vbCrLf, " ")
  End If
'    Call m_cTT.AddTool(p1(Index))
'    m_cTT.ToolText(p1(Index)) = tt$
End If
mmblock = False
End Sub

Sub rdsels()
Dim gw$, fsel$, bisi%, kid$, old As Variant, tpid$
Dim dv$, db$, selstr$, cmd$, nosel As Integer, shwpriv As Boolean
Dim r As ADODB.Recordset, rrr, c_stat0 As Long, optcol As Boolean, col As Long
Dim prvid$, offs%, i%, cbz$, gw1$, ent2 As Boolean, wasalles As String
Dim dkz As Boolean, noshow As Boolean, tpidokcache$, tpidnokcache$
Dim prz As Boolean, pvon, pbis, idat, d0

Dim d2infile As String, d2insub As String
d2infile = "k3": d2insub = "rdsels"
c_stat0 = RGB(255, 255, 255)
dv$ = datum2sql(olday$)
db$ = datum2sql(CDate(dv$) + 41)
For offs% = 0 To 41: lip%(offs%) = -1: Next offs%
d0 = CDate(tag0$) - 7
While Weekday(CDate(d0), vbMonday) - 1 <> fdow%
  d0 = d0 - 1
Wend
olday$ = d0
old = CDate(olday$)

prz = False
If form1.getusersetting("Projektezeigen", "nein") = "ja" Then
  prz = True
  cmd$ = "select * from tplan where (Hauptperson<>'Dekade') and (von>='" + dv$ + "' and von<='" + db$ + "') or (bis>='" + dv$ + "' and bis<='" + db$ + "') or (von<'" + dv$ + "' and bis>'" + db$ + "')"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
  While Not r.EOF
    On Error Resume Next
    pvon = Max(CDate(r!von), CDate(old))
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then pvon = CDate(old)
    On Error Resume Next
    pbis = MyMin(CDate(r!bis), CDate(old) + 41)
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      pbis = MyMin(CDate(r!von), CDate(old) + 41)
    End If
    For idat = pvon To pbis
      offs% = idat - old
      lip%(offs%) = lip%(offs%) + 1
      c_id$(offs%, lip%(offs%)) = r!id
      c_typ$(offs%, lip%(offs%)) = r!hauptperson
      c_stat(offs%, lip%(offs%)) = c_stat0
      c_col(offs%, lip%(offs%)) = form1.projektfarbe(trm(r!hauptperson))
      c_bez$(offs%, lip%(offs%)) = "Projekt: " + trm(r!id)
      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry%)-(80, (lip%(offs%) + 1) * ypixperentry%), c_stat(offs%, lip%(offs%)), BF
      p1(offs%).Line (80, (lip%(offs%) + 2) * ypixperentry%)-(0, (lip%(offs%) + 1) * ypixperentry%), c_col(offs%, lip%(offs%)), BF
      p1(offs%).Print c_bez$(offs%, lip%(offs%))
      'Linie drüber
      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry% - 1)-(0, (lip%(offs%) + 2) * ypixperentry% - 1), 0
    Next idat
    DoEvents
    r.MoveNext
  Wend
  End If
End If
On Error GoTo exrds
selstr$ = ""
selstr$ = selstr$ + "((datum>='" + dv$ + "' and datum<='" + db$ + "')) "
gw1$ = kc.getwho()
optcol = False
If kc.selct(2).ListCount = 0 And gw1$ = "" Then
  gw$ = kc.getwhere()
  wasalles = "id as aid,astatus,TourneeplanID,datum as adatum,auftritt.zeit as azeit, bezeichnung as abez,ort as aort, auftrittstyp as atyp "
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    wasalles = wasalles + ", optkalcolor as tf "
    optcol = True
  End If
  cmd$ = "SELECT " + wasalles + " from auftritt "
  If gw$ = "" Then
    gw$ = "where "
  Else
    gw$ = gw$ + " and "
  End If

  cmd$ = cmd$ + gw$ + selstr$
Else
  wasalles = "auftritt.id as aid,astatus,auftritt.TourneeplanID,auftritt.datum as adatum,auftritt.zeit as azeit,auftritt.bezeichnung as abez,auftritt.ort as aort, auftritthigru.auftrittstyp as atyp, auftritthigru.FeldName, auftritthigru.Felddaten "
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    wasalles = wasalles + ", auftritt.optkalcolor as tf "
    optcol = True
  End If
  cmd$ = "SELECT " + wasalles
  cmd$ = cmd$ + " FROM auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid "
  gw$ = kc.getwhere()
  If gw$ = "" Then
    cmd$ = cmd$ + " Where "
  Else
    cmd$ = cmd$ + gw$ + " and "
  End If
  nosel = 1
  For i% = 0 To kc.selct(2).ListCount - 1
    If kc.selct(2).Selected(i%) = True Then
      nosel = 0
      fsel$ = kc.selct(2).List(i%)
      i% = kc.selct(2).ListCount
    End If
  Next i%
  bisi% = 30
  If nosel = 1 Then
    kid$ = Trim("" & kc.selct(2).List(0))
    kid$ = "(instr(FeldDaten,'" + kid$ + "')>0) "
    For i% = 1 To kc.selct(2).ListCount - 1
      kid$ = kid$ + "or (instr(FeldDaten,'" + Trim("" & kc.selct(2).List(i%)) + "')>0) "
      bisi% = bisi% - 1
      If bisi% < 0 Then i% = kc.selct(2).ListCount - 1
    Next i%
  Else
    kid$ = fsel$
    kid$ = "(instr(FeldDaten],'" + kid$ + "')>0) "
    bisi% = kc.selct(2).ListCount - 1: If bisi% > 20 Then bisi% = 20
    For i% = 0 To bisi%
      If kc.selct(2).Selected(i%) = True And kc.selct(2).List(i%) <> fsel$ Then
        kid$ = kid$ + "or (instr(FeldDaten,'" + Trim("" & kc.selct(2).List(i%)) + "')>0) "
        bisi% = bisi% - 1
        If bisi% < 0 Then i% = kc.selct(2).ListCount - 1
      End If
    Next i%
  End If
  cmd$ = cmd$ + " ( " + kid$ + ") and  "
  cmd$ = cmd$ + selstr$
End If
cmd$ = cmd$ + " ORDER BY auftritt.datum,auftritt.zeit"
dkz = False
If form1.getusersetting("Dekadenzeigen", "nein") = "ja" Then dkz = True
If form1.getusersetting("Privateszeigen", "nein") = "ja" Then shwpriv = True
'daten selektieren
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  cmd$ = strrepl(cmd$, ",astatus", "")
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Sub
End If
prvid$ = "-x"
adrgrpselcache$ = "|"
adrgrpnoselcache$ = "|"
tpidokcache$ = ""
tpidnokcache$ = ""
While Not r.EOF
  ent2 = True
  noshow = False
  If Not dkz Then
    tpid$ = trm(r!TourneeplanID)
    If InStr(tpidnokcache$, tpid$) > 0 Then noshow = True
    If Not noshow And tpid$ <> "" And tpid$ <> "-1" And InStr(tpidokcache$, tpid$) = 0 Then
      If form1.projekttyp(tpid$) = "Dekade" Then
        noshow = True
        tpidnokcache$ = tpidnokcache$ + " " + tpid$
      Else
        tpidokcache$ = tpidokcache$ + " " + tpid$
      End If
    Else
      tpidokcache$ = tpidokcache$ + " " + tpid$
    End If
  End If
  If gw1$ <> "" Then ent2 = adrisinselectedgroup(trm(r!felddaten), gw1$)
  If r!aid <> prvid$ And ent2 Then
    prvid$ = r!aid
    If Not shwpriv And trm(r!atyp) = "Privat" Then noshow = True
    If Not noshow Then
      On Error Resume Next
      offs% = CDate(datfromsql(r!adatum)) - CDate(old)
      On Error GoTo 0
Call form1.dbg2f("zeige: " + trm(r!adatum) + " " + trm(r!atyp) + " " + trm(r!aid) + " " + trm(r!abez) + " offset: " + trm(offs%))
      If offs% < 0 Then offs% = 0
      If offs% > 41 Then offs% = 41
      lip%(offs%) = lip%(offs%) + 1
      c_id$(offs%, lip%(offs%)) = r!aid
      If Not IsNull(r!atyp) Then
        c_typ$(offs%, lip%(offs%)) = r!atyp
        If optcol Then
          c_col(offs%, lip%(offs%)) = Val(trm0(r!tf))
        Else
          c_col(offs%, lip%(offs%)) = -1
        End If
        If c_col(offs%, lip%(offs%)) <= 0 Then c_col(offs%, lip%(offs%)) = form1.get_eventcolor(r!atyp)
        On Error Resume Next
        c_stat(offs%, lip%(offs%)) = form1.get_eventstatuscolor(r!astatus)
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then c_stat(offs%, lip%(offs%)) = c_stat0
      End If
      cbz$ = trm(r!aort & " " & r!abez)
      If r!TourneeplanID <> -1 Then cbz$ = cbz$ & " " & r!TourneeplanID
      If trm(r!azeit) <> "" Then cbz$ = cbz$ & " " & r!azeit & " h"
      c_bez$(offs%, lip%(offs%)) = cbz$
      p1(offs%).Line (100, (lip%(offs%) + 2) * ypixperentry%)-(80, (lip%(offs%) + 1) * ypixperentry%), c_stat(offs%, lip%(offs%)), BF
      p1(offs%).Line (80, (lip%(offs%) + 2) * ypixperentry%)-(0, (lip%(offs%) + 1) * ypixperentry%), c_col(offs%, lip%(offs%)), BF
      p1(offs%).Print c_bez$(offs%, lip%(offs%))
      DoEvents
    End If
  End If
  r.MoveNext
Wend
lipv% = 1
Call buttonset
exrds:
On Error GoTo 0

End Sub

Sub c_draw()
Dim i%, j%

'd2infile = "k3": d2insub = "c_draw"
ttform.Hide
ypixperentry% = 16
For i% = 0 To 41
  For j% = 0 To lip%(i%)
Call form1.dbg2f("c_draw  (i/j)=" + trm(i%) + "/" + trm(j%) + ": " + c_bez$(i%, j%))
    p1(i%).Line (100, (j% + 2) * ypixperentry%)-(80, (j% + 1) * ypixperentry%), c_stat(i%, j%), BF
    p1(i%).Line (80, (j% + 2) * ypixperentry%)-(0, (j% + 1) * ypixperentry%), c_col(i%, j%), BF
    p1(i%).Print c_bez$(i%, j%)
  Next j%
Next i%
Call buttonset
End Sub

Private Sub pred_Click()
Dim id$, prjid$, prj As Boolean

'd2infile = "k3": d2insub = "pred_Click"
If kc.List1.ListIndex < 0 Then Exit Sub
id$ = kc.List1.List(kc.List1.ListIndex)
If InStr(id$, " Projekt: ") > 0 Then prj = True
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
If Not prj Then
  prjid$ = form1.get_projectid_by_aid(id$)
Else
  prjid$ = id$
End If
Load tplan
tplan.setcaption ("Crossover - Projekt")
Call tplan.SetFocus
Call tplan.showrec(prjid$)

End Sub

Private Sub project_Click()
Dim neuid As String, d As Variant, s$

'd2infile = "k3": d2insub = "project_Click"
d = CDate(olday$)
neuid = CDate(d + prvi%)
neuid = " " & Mid(neuid, 4, 2) & " " & apyear(neuid)
neuid = trm(InputBox(transe("Neue Projekt-ID:"), transe("Neues Projekt erstellen"), neuid))
If trm(neuid) = "" Then Exit Sub

s$ = "insert into tplan (id,kuerzel,hauptperson,von) values('" + neuid$ + "','" + Left$(neuid$, 4) + "','Künstler','" + datum2sql(trm(CDate(d + prvi%))) + "')"
Call form1.sqlqry(s$)

On Error Resume Next
Load tplan
tplan.SetFocus
On Error GoTo 0
DoEvents

tplan.Text2.text = neuid$

End Sub

Private Sub termin_Click()
Dim d As Variant, id$

'd2infile = "k3": d2insub = "termin_Click"
  d = CDate(olday$)
  id$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 id$ + "','-1'" + _
                 ",'Neuer Auftritt','" + transe("Neuer Auftritt") + "','" + _
                 datum2sql(CDate(d + prvi%)) + "')")
  Unload auftritt
  DoEvents
  Load auftritt
  Call auftritt.SetFocus
  Call auftritt.showrec(id$, 0)

End Sub

Private Sub tmrstrt_Click()
Dim X, o%

If form1.weckerpresent Then
  X = Shell(form1.s00dir() + "\wecker.exe", vbNormalFocus)
Else
  tmrstrt.Enabled = False
End If
End Sub

Private Sub todo_Click()
Dim neuid As String, d As Variant

'd2infile = "k3": d2insub = "todo_Click"
d = CDate(olday$)
neuid = CDate(d + prvi%)
Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), "" _
             , "", neuid, Left(Time, 5))
create2do.Text1(1).Enabled = False
Call create2do.SetFocus

End Sub

Private Sub trmed_Click()
'd2infile = "k3": d2insub = "trmed_Click"
Call p1_DblClick(prvi%)
End Sub


Sub buttonset()
'up4.Left = ScaleWidth - up4.Width
dwn12.Top = ScaleHeight - dwn12.Height
dwn12.Left = ScaleWidth - dwn12.Width
dwn4.Left = dwn12.Left - dwn4.Width: dwn4.Top = dwn12.Top
up4.Top = dwn4.Top: up4.Left = dwn4.Left - up4.Width
up12.Top = dwn4.Top: up12.Left = up4.Left - up12.Width
zoomin.Top = up12.Top: zoomout.Top = up12.Top
zoomin.Left = up12.Left - zoomin.Width: zoomout.Left = zoomin.Left - zoomout.Width:
End Sub

Private Sub up12_Click()
Call kc.Command7_Click
End Sub

Private Sub up4_Click()
Call kc.Command4_Click
End Sub

Private Sub zoomin_Click()
zoom = zoom + 1
zoomout.Enabled = True
If zoom > 3 Then zoom = 3
If zoom = 3 Then zoomin.Enabled = False
Call form1.setusersetting("zkalzoom", trm(zoom))
Call kc.Command1_Click
End Sub

Private Sub zoomout_Click()
zoom = zoom - 1
zoomin.Enabled = True
If zoom < 1 Then zoom = 1
If zoom = 1 Then zoomout.Enabled = False
Call form1.setusersetting("zkalzoom", trm(zoom))
Call kc.Command1_Click
End Sub

Private Sub mvwk(n%)
Dim tg1
tg1 = datum2sql(CDate(tag0) + n%)
Call settag0(trm(tg1))
End Sub

