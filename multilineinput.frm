VERSION 5.00
Begin VB.Form multilineinput 
   Caption         =   "Form2"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Go"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Abbruch"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK, speichern"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   0
      Text            =   "multilineinput.frx":0000
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "multilineinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txt As String, cmd$, p%, r$

Dim d2infile As String, d2insub As String

d2infile = "multilineinput": d2insub = "Command1_Click"
cmd$ = Text2.text
p% = InStr(Text2.text, "$$$")
If p% > 0 Then
  If Len(Text1.text) > 0 Then
    txt = Text1.text
    If txt = "=GUI" Then txt = GUID()
    txt = strrepl(txt, "'", "´")
    cmd$ = Left$(cmd$, p% - 1) + txt + Mid$(cmd$, p% + 3)
  Else
    r$ = Mid$(cmd$, p% + 3)
    cmd$ = Left$(cmd$, p% - 2) + "NULL"
    If Len(r$) > 1 Then cmd$ = cmd$ + Mid$(r$, 2)
  End If
End If
form1.sqlqry (cmd$)
txt = trm(Text3.text)
If txt <> "" Then
  Call form1.sqlqry(txt)
End If

Call shwAdrDetail.reshow4

Call Command2_Click
End Sub

Private Sub Command2_Click()
Hide
Unload multilineinput
End Sub

Private Sub Command3_Click()
Dim l$, p%, t0 As Double, tn As Double, l1$, t$, X

l$ = trm(Text1.text)
If Command3.Caption = "Anwahl" Then
  Load AutoAnwahl
  Call AutoAnwahl.SetFocus
  AutoAnwahl.nummer.text = l$
  DoEvents
  Call AutoAnwahl.cmdDial_Click
Else
  While Len(l$) > 0
    p% = InStr(l$, vbCrLf)
    If p% > 0 Then
      l1$ = Left$(l$, p% - 1)
      l$ = trm(Mid$(l$, p% + 2))
    Else
      l1$ = l$
      l$ = ""
    End If
    l1$ = trm(l1$)
    Select Case LCase(word1(l1$))
      Case "shell": t$ = trm(Mid$(l1$, InStr(l1$, " ") + 1))
         X = Shell(t$, 1)
         t0 = Date + Time: While (Date + Time) - t0 < 2 / 86400: DoEvents: Wend
      Case "sendkeys": t$ = trm(Mid$(l1$, InStr(l1$, " ") + 1))
         Call form1.SendKys(t$, 1)
         t0 = Date + Time: While (Date + Time) - t0 < 1 / 86400: DoEvents: Wend
    Case Else:
    End Select
  Wend
End If
Call Command2_Click

End Sub

Private Sub Form_Load()

Text3.text = ""
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim d2infile As String, d2insub As String
d2infile = "multilineinput": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0


End Sub
Public Sub init(p$)
Dim d2infile As String, d2insub As String, z
Dim offset
d2infile = "multilineinput": d2insub = "Init"
Text1.text = p$
z = Val(Left$(p$, InStr(p$, ":") - 1))
p$ = Mid$(p$, InStr(p$, ":") + 1)
Text2.text = p$
offset = (z - 1) * 240
Text1.Height = 285 + offset
Command1.Top = Text1.Height + Text1.Top + 80
Command2.Top = Command1.Top
Command3.Top = Command1.Top
DoEvents
multilineinput.Height = Text1.Height + Text1.Top + Command1.Height + 660

End Sub
Public Sub setdeflt(p$)
Dim d2infile As String, d2insub As String, z
d2infile = "multilineinput": d2insub = "setdeflt"
Text1.text = p$

End Sub

Public Sub setcap(p$)
Dim d2infile As String, d2insub As String, z
d2infile = "multilineinput": d2insub = "setcap"
Me.Caption = p$
If p$ = "Action" Then
  Command3.Visible = True
  Command3.Caption = "Go"
End If
If Left$(LCase(p$), 4) = "tel-" Then
  Command3.Visible = True
  Command3.Caption = "Anwahl"
End If

End Sub

