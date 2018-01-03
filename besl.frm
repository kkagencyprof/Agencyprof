VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Besetzungen"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   5040
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "besl.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   4440
      Width           =   495
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   5
      Left            =   8640
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   4
      Left            =   6960
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   3
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox splt 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4215
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim besmax

Private Sub Combo2_Change()

End Sub

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rrr

rrr = 0
besmax = 0
Do

Loop Until rrr



  
End Sub
