VERSION 5.00
Begin VB.Form auftrittrepeat 
   Caption         =   "Wiederholungen des Auftritts"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   120
      Picture         =   "auftrittrepeat.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   16
      ToolTipText     =   "Abbrechen"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      Picture         =   "auftrittrepeat.frx":0164
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Speichern"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "auftrittrepeat.frx":06A8
      Left            =   3600
      List            =   "auftrittrepeat.frx":06F4
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      ItemData        =   "auftrittrepeat.frx":07A0
      Left            =   2160
      List            =   "auftrittrepeat.frx":07C2
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      ItemData        =   "auftrittrepeat.frx":07E5
      Left            =   3240
      List            =   "auftrittrepeat.frx":07F5
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      ItemData        =   "auftrittrepeat.frx":0819
      Left            =   2640
      List            =   "auftrittrepeat.frx":083E
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   1
      ItemData        =   "auftrittrepeat.frx":0865
      Left            =   3600
      List            =   "auftrittrepeat.frx":0867
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "am"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "um"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Uhr"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "in"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Termin"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Zeitraum"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   4800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "danach"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "mal, alle"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1455
      Left            =   600
      Shape           =   4  'Gerundetes Rechteck
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "auftrittrepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d0 As Variant, delta As Long, deltaunit$
Dim perdeltau$
Dim perdelta As Integer


Private Sub Combo2_Change(Index As Integer)
d2infile = "auftrittrepeat": d2insub = "Combo2_Change"
Call deltaset(Index)
End Sub

Private Sub Combo2_Click(Index As Integer)
d2infile = "auftrittrepeat": d2insub = "Combo2_Click"
Call deltaset(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)

d2infile = "auftrittrepeat": d2insub = "Combo2_LostFocus"
Call deltaset(Index)

End Sub

Private Sub Combo3_Change(Index As Integer)
d2infile = "auftrittrepeat": d2insub = "Combo3_Change"
Call deltaset(Index)
End Sub

Private Sub Combo3_Click(Index As Integer)
d2infile = "auftrittrepeat": d2insub = "Combo3_Click"
Call deltaset(Index)
End Sub

Private Sub Combo3_LostFocus(Index As Integer)

d2infile = "auftrittrepeat": d2insub = "Combo3_LostFocus"
Call deltaset(Index)

End Sub

Private Sub Command2_Click()
Dim ast As Integer

d2infile = "auftrittrepeat": d2insub = "Command2_Click"
form1.fastsave_copy = True
poft = Val(Text2.Text)
ast = form1.auftrittsstatus(trm(auftritt.Text1(0).Text))
While poft > 0
  poft = poft - 1
  If perdelta <> 0 And trm(perdeltau$) <> "" Then
    auftritt.setast = ast
    Call auftritt.Command7_Click
    auftritt.setast = -1
    DoEvents
    dat$ = datum2sql(auftritt.Text1(2).Text)
    yyyy% = Val(Left$(dat$, 4))
    mm% = Val(Mid$(dat$, 6, 2))
    dd% = Val(Right$(dat$, 2))
    d0 = CDate(datfromsql(dat$))
    Select Case perdeltau$
      Case "ta": d0 = CDate(d0 + perdelta)
      Case "mo": ti% = mm%
                 While ti% > 0
                   mm% = mm% + 1
                   If mm% > 12 Then
                     yyyy% = yyyy% + 1
                   End If
                 Wend
                 d0 = CDate(trm(yyyy%) & "-" & trm(mm%) & "-" & trm(dd%))
      Case "ja": ti% = perdelta
                 While ti% > 0
                   ti% = ti% - 1
                   yyyy% = yyyy% + 1
                 Wend
                 d0 = CDate(trm(yyyy%) & "-" & trm(mm%) & "-" & trm(dd%))
      Case "wo": d0 = CDate(d0 + perdelta * 7)
      Case Default:
    End Select
    Call auftritt.Text1_GotFocus(2): DoEvents
    auftritt.Text1(2).Text = d0: DoEvents
    Call auftritt.Text1_LostFocus(2): DoEvents
    Call auftritt.Command10_Click: DoEvents
  End If
Wend
form1.fastsave_copy = False
Call Command3_Click
End Sub

Private Sub Command3_Click()
d2infile = "auftrittrepeat": d2insub = "Command3_Click"
Unload auftrittrepeat
End Sub

Private Sub Form_Load()
Dim i As Integer

d2infile = "auftrittrepeat": d2insub = "Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
auftrittrepeat.Caption = transe("Wiederholungen des Auftritts")
Command3.ToolTipText = transe("Formular schliessen")
Command2.ToolTipText = transe("Speichern")
Label1(5).Caption = transe("am")
Label1(6).Caption = transe("um")
Label2.Caption = transe("Uhr")
Label3.Caption = transe("in")
Label6.Caption = transe("Termin")
Label7.Caption = transe("Zeitraum")
Label8.Caption = transe("danach")
Label9.Caption = transe("mal, alle")
Show
Combo2(1).ToolTipText = transe("0=nie")
For i = 0 To 1
  Combo3(i).Clear
  Combo3(i).AddItem transe("Tage")
  Combo3(i).AddItem transe("Woche")
  Combo3(i).AddItem transe("Monate")
  Combo3(i).AddItem transe("Jahre")
Next i
Text1.Text = auftritt.Text1(2).Text
Combo1.Text = Left(auftritt.Text1(3).Text, 5)

End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "auftrittrepeat": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
form1.fastsave_copy = False
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Sub deltaset(i%)
Dim df As Double

d2infile = "auftrittrepeat": d2insub = "deltaset"
On Error Resume Next
delta = Val(Combo2(i%).Text)
rrr = Err
On Error GoTo 0
deltaunit$ = Combo3(i%).Text
If trm(deltaunit$) = "" Or rrr <> 0 Then Exit Sub
Select Case LCase$(Left$(deltaunit$, 2))
  Case "ta": df = 1
  Case "wo": df = 7
  Case "mo": df = 30
  Case "ja": df = 365.25
  Case Default: df = 0
End Select

perdeltau$ = LCase$(Left$(deltaunit$, 2))
perdelta = delta

On Error GoTo 0
End Sub

