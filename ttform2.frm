VERSION 5.00
Begin VB.Form ttform 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   0  'Kein
   Caption         =   "Form2"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      Height          =   2850
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4670
   End
End
Attribute VB_Name = "ttform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wichtige Eigenschaften

'Form2 BorderStyle 0-Kein

'Textbox    Appearance 0-2D
'           BorderStyle 1-Fest Einfach
'           BackColor -> System -> QuickInfo
'           MultiLine  True


'Um die Textbox dem Formular anzupasen
Private Sub Form_Resize()
Text1.Top = 0
Text1.Left = 0
'Text1.Height = Me.ScaleHeight
'Text1.Width = Me.ScaleWidth
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Hide
End Sub
