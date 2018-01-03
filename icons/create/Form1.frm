VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   915
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   1335
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   915
      Index           =   2
      Left            =   3000
      Picture         =   "Form1.frx":3DE2
      ScaleHeight     =   855
      ScaleWidth      =   1335
      TabIndex        =   4
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   915
      Index           =   1
      Left            =   1560
      Picture         =   "Form1.frx":7BC4
      ScaleHeight     =   855
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   915
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":B9A6
      ScaleHeight     =   855
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()

For z = 0 To 25
For i = 0 To 1
  P1(i).Cls
  P1(i).Picture = P1(2 + i).Picture
  P1(i).FontSize = 10:  P1(i).Print
  P1(i).FontSize = 12:  P1(i).Print "         ";
  P1(i).FontSize = 18
  P1(i).Print Chr$(Asc("A") + z)
  DoEvents
  fn$ = "lbl-" & Chr$(Asc("A") + z) & "-" & Trim(Str(i)) & ".bmp"
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
  SavePicture P1(i).Image, fn$
Next i
Next z
End Sub

