VERSION 5.00
Begin VB.Form fehler 
   Caption         =   "aufgetretene Fehler:"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3315
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "alle zeigen"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "alle löschen"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "fehler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload fehler
End Sub

Private Sub Command2_Click()
Dim tr

tr = Dir(form1.s0dir() & "\" + form1.docs() + "\" & form1.getuserid() & "*.err")
While tr <> ""
  On Error Resume Next
  Kill form1.s0dir() & "\" + form1.docs() + "\" & tr
  On Error GoTo 0
  tr = Dir
Wend
Call Command1_Click

End Sub

Private Sub Command3_Click()
Dim tr, o As Integer, p As Integer, fo As String, fi As String
Dim l As String, x, rrr

o = FreeFile
fo = form1.s0dir() & "\" + form1.docs() + "\" & form1.getuserid() & "fehlerliste.err"
Open fo For Output As #o%
tr = Dir(form1.s0dir() & "\" + form1.docs() + "\" & form1.getuserid() & "*.err")
While tr <> ""
  p = FreeFile
  fi = form1.s0dir() & "\" + form1.docs() + "\" & tr
  On Error Resume Next
  Open fi For Input As #p
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    While Not EOF(p)
      Line Input #p, l
      Print #o, l
    Wend
    Close #p
  End If
  tr = Dir
Wend
Close #o
x = Shell("notepad.exe " + fo, 1)
DoEvents
On Error Resume Next
Kill fo
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim tr

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

fehler.Caption = transe("aufgetretene Fehler:")
Command2.Caption = transe("alle löschen")
Command1.Caption = transe("&Schliessen")

Show
List1.Clear
tr = Dir(form1.s0dir() & "\" + form1.docs() + "\" & form1.getuserid() & "*.err")
While tr <> ""
  List1.AddItem tr
  DoEvents
  tr = Dir
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0
End Sub

Private Sub List1_DblClick()
Dim x

x = Shell("notepad.exe " + form1.s0dir() & "\" + form1.docs() + "\" & List1.List(List1.ListIndex), 1)

End Sub
