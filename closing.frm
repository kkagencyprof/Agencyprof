VERSION 5.00
Begin VB.Form closing 
   Caption         =   "POP-Client"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2775
   LinkTopic       =   "Form2"
   ScaleHeight     =   1020
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "... wird beendet ..."
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "closing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = popmain.Top
Me.Left = popmain.Left
Show
Timer1.Interval = 2000
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload closing
DoEvents
Call Form1.Command1_Click

End Sub
