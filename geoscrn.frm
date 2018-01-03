VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form geoscrn 
   Caption         =   "OpenGeoDB"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   4320
      TabIndex        =   7
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4680
      Picture         =   "geoscrn.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Ansicht aktualisieren"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "geoscrn.frx":0B66
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Formular schliessen"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Text            =   "0,3"
      Top             =   4680
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   2760
      Top             =   4440
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label centerid 
      Caption         =   "centerid"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "geoscrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currval As Boolean, cl0 As Double, cb0 As Double, csc As Double

Private Sub centerid_Change()
Dim id$, rtmp As ADODB.Recordset, popu As Double
Dim b As Double, l As Double, c$, scle As Double
Dim b0 As Double, l0 As Double, pb As Single, pl As Single

Dim d2infile As String, d2insub As String

d2infile = "geoscrn": d2insub = "centerid_Change"
id$ = centerid.Caption
p1.Cls
c$ = "SELECT breite,laenge,name,name_int FROM geodb_locations where id='" + id$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f(c$)
rtmp.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly
If rtmp.EOF Then Exit Sub
b = trm0(rtmp!breite)
l = trm0(rtmp!laenge)
scle = var2dbl(trm0(Text1.Text))
b0 = trm0(rtmp!breite) - scle / 2
l0 = trm0(rtmp!laenge) - scle / 2
cl0 = l0: cb0 = b0: csc = scle
currval = True
Me.Caption = "OpenGeoDB: " + trm(rtmp!name_int) + " L:" + fixeur(l) + " B:" + fixeur(b)
'c$ = "SELECT * FROM geodb_locations " & _
'         "where breite > '" + d2db(b - 1) + "' and breite < '" + d2db(b + 1) + _
'         "' and laenge > '" + d2db(l - 1) + "' and laenge < '" + d2db(l + 1) + ";"
p1.ScaleWidth = 1000
p1.ScaleHeight = 1000

c$ = "SELECT * FROM geodb_locations " & _
         "where typ = 6 and breite > " + d2db(b - scle / 2) + " and breite < " + d2db(b + scle / 2) + _
         " and laenge > " + d2db(l - scle / 2) + " and laenge < " + d2db(l + scle / 2) + ";"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f(c$)
rtmp.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly
p1.FontSize = 6
While Not rtmp.EOF
  pb = 1000 - (trm0(rtmp!breite) - b0) * 1000 / scle
  pl = (trm0(rtmp!laenge) - l0) * 1000 / scle
  popu = geogetpop(rtmp!id) / 20000
  p1.Circle (pl, pb), imax(Int(popu), 2), RGB(0, 0, 0)
  p1.Print trm(rtmp!name)
  DoEvents
  rtmp.MoveNext
Wend

End Sub

Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command2_Click()
Call centerid_Change
End Sub

Private Sub Command21_Click()
Dim z As Double

z = trm0(Text1.Text) - 0.1
If z < 0.05 Then z = 0.05
Text1.Text = trm(z)
Call Command2_Click
End Sub

Private Sub Command3_Click()
Dim z As Double

z = trm0(Text1.Text) + 0.1
Text1.Text = trm(z)
Call Command2_Click
End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Label1.Caption = ""

End Sub

Private Sub Form_Resize()
axsResizer1.Resize
Call centerid_Change

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub

Function geogetpop(id$) As Double
Dim c$
Dim rtmp As ADODB.Recordset

geogetpop = 0

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT population_min FROM geodb_population where id='" + id$ + "'"
Call form1.dbg2f("geogetpop:" + c$)
rtmp.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly

If rtmp.EOF Then Exit Function
'If IsNull(rtmp!Name) Or IsNull(rtmp!vornamen) Then Exit Function

geogetpop = rtmp!population_min

End Function

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim geob As Double, geol As Double, M As Double
Dim geob_m As Integer, geob_s As Integer

If Not currval Then Exit Sub

M = 1000 / csc
geob = cb0 + (1000 - Y) / M
geob_m = 60 * (geob - Int(geob))
geol = cl0 + X / M
Label1.Caption = fixeur(geob) + "/" + fixeur(geol)

End Sub

Private Sub p1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim geob As Double, geol As Double, M As Double, c$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String

d2infile = "geoscrn": d2insub = "p1_MouseUp"
If Not currval Then Exit Sub
MousePointer = 11: DoEvents
M = 1000 / csc
geob = cb0 + (1000 - Y) / M
geol = cl0 + X / M
c$ = "select id,name,abs(" + trm(d2db(geob)) + " - breite) + abs(" + trm(d2db(geol)) + " - laenge) as dist from geodb_locations where typ<4712 order by dist;"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("p1_MouseUp:" + c$)
rtmp.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly

If Not rtmp.EOF Then
  centerid.Caption = rtmp!id
End If
MousePointer = 0

End Sub
