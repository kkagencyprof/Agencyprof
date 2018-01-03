VERSION 5.00
Begin VB.Form startup 
   Caption         =   "Agencyprof wird gestartet ..."
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   3360
      Picture         =   "startup.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2715
      TabIndex        =   2
      ToolTipText     =   "Herzlich willkommen!"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5760
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6255
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Contains cryptography software by David Ireland of DI Management Services Pty Ltd <www.di-mgt.com.au>."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: kk@agencyprof.de"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.Agencyprof.de"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2004,2005  Karsten Kaus"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3615
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d2infile = "startup": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Form_Load()
d2infile = "startup": d2insub = "Form_Load"
Me.Top = 5
Me.Left = 5
Label2.Caption = "Copyright (C) 2001-" + trm(Year(Date)) + " Karsten Kaus"
Me.Caption = form1.inmylanguage("Agencyprof wird gestartet ...")
Command1.Caption = form1.inmylanguage("Schliessen")
Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "startup": d2insub = "Form_Unload"
Hide

End Sub

Private Sub Timer1_Timer()
d2infile = "startup": d2insub = "Timer1_Timer"
If List1.ListCount < 0 Then Exit Sub

If List1.List(List1.ListCount - 1) = transe("fertig.") Then
  Unload Me
End If
End Sub
