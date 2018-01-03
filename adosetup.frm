VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ADOSetup durchgeführt."
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"adosetup.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End

End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer

End Sub
