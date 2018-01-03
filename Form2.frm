VERSION 5.00
Begin VB.Form pwenter 
   Caption         =   "Passwort eingeben"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2910
   LinkTopic       =   "Form2"
   ScaleHeight     =   975
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "&Abbruch"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Passwort:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "pwenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
form1.pwentered = trm(Text1.Text)
Unload Me
End Sub

Private Sub Command2_Click()
form1.pwentered = "passwo2345rteingabeabge4356brochen11223476"
Unload Me

End Sub

Private Sub Form_Load()
Show
End Sub
