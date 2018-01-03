VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form memoview 
   Caption         =   "Memo Anzeige"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows-Standard
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Text            =   "memoview.frx":0000
      Top             =   120
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "memoview.frx":0006
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   4320
      Width           =   8535
   End
End
Attribute VB_Name = "memoview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

d2infile = "memoview": d2insub = "Command1_Click"
Unload memoview
End Sub

Private Sub Form_Load()
d2infile = "memoview": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Show
Call settext("")
End Sub

Private Sub Form_Resize()
d2infile = "memoview": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "memoview": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub
Public Sub settext(t$)

d2infile = "memoview": d2insub = "settext"
If t$ = "" Then Exit Sub
Text1.Text = ""
Text1.Text = t$
End Sub

