VERSION 5.00
Begin VB.Form diverse 
   Caption         =   "diverse importe"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Ende"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "diverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load import_schessl
Call import_schessl.SetFocus

End Sub

Private Sub Command2_Click()

End
End Sub

Private Sub Form_Load()

Show
End Sub
