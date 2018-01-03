VERSION 5.00
Begin VB.Form fselect 
   Caption         =   "Datei auswählen"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox fqn 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Abbrechen"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Dateien"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Verzeichnisse"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Laufwerk:"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "fselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn$, fn$
Dim fqfn$

Private Sub Combo1_Click()
d2infile = "fselect": d2insub = "Combo1_Click"
Text1.text = Combo1.List(Combo1.ListIndex) + "\"
End Sub

Private Sub Command1_Click()

d2infile = "fselect": d2insub = "Command1_Click"
Hide

End Sub

Private Sub Command2_Click()
d2infile = "fselect": d2insub = "Command2_Click"
If Text2.text <> "" Then
  fqfn$ = dn$ + Text2.text
  fqn.text = fqfn$
End If
Hide
End Sub

Private Sub Form_Load()
d2infile = "fselect": d2insub = "Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
fselect.Caption = transe("Datei auswählen")
Command2.Caption = transe("&OK")
Command1.Caption = transe("&Abbrechen")
Label3.Caption = transe("Dateien")
Label2.Caption = transe("Verzeichnisse")
Label1.Caption = transe("Laufwerk:")
Show

Combo1.Clear
Combo1.text = Left$(form1.s0dir(), 3)
Text1.text = form1.s0dir()
For i% = 2 To 25
  tr = Chr$(Asc("A") + i%) + ":\*.*"
  On Error Resume Next
  tr = Dir(tr)
  rrr = Err
  On Error GoTo 0
  If tr <> "" And rrr = 0 Then Combo1.AddItem Chr$(Asc("A") + i%) + ":"
Next i%
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "fselect": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub


Private Sub List1_DblClick()

d2infile = "fselect": d2insub = "List1_DblClick"
dd$ = List1.List(List1.ListIndex)
If dd$ <> ".." Then
  Text1.text = dn$ + dd$
Else
  dn$ = Left$(dn$, Len(dn$) - 1)
  If InStr(dn$, "\") > 0 Then
    While Right$(dn$, 1) <> "\"
      dn$ = Left$(dn$, Len(dn$) - 1)
    Wend
  End If
  Text1.text = dn$
End If
End Sub

Private Sub List2_Click()
d2infile = "fselect": d2insub = "List2_Click"
Text2.text = List2.List(List2.ListIndex)

End Sub

Private Sub List2_DblClick()
d2infile = "fselect": d2insub = "List2_DblClick"
Call Command2_Click
End Sub

Private Sub Text1_Change()
Dim res, rrr

d2infile = "fselect": d2insub = "Text1_Change"
List1.Clear
List2.Clear
Text2.text = ""
t$ = Text1.text
If Right$(t$, 1) <> "\" Then t$ = t$ + "\"
dn$ = t$
tr = Dir(t$, vbDirectory)
Do While tr <> ""
  On Error Resume Next
  res = GetAttr(dn$ + tr) And vbDirectory
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
  If res = vbDirectory Then
     List1.AddItem tr
  End If
  End If
  tr = Dir
Loop
t$ = Text1.text
If Right$(t$, 1) <> "\" Then t$ = t$ + "\"
t$ = t$ + "*.*"
tr = Dir(t$)
Do While tr <> ""
  List2.AddItem tr
  tr = Dir
Loop

End Sub

Private Sub Text2_Change()
d2infile = "fselect": d2insub = "Text2_Change"
fn$ = Text2.text
End Sub
