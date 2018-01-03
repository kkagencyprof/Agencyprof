VERSION 5.00
Begin VB.Form tlform 
   BorderStyle     =   0  'Kein
   Caption         =   "Form2"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "tlform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
form1.tlopen = False
Hide
End Sub

Private Sub Form_Resize()
List1.Top = 0
List1.Left = Command1.Width + 40
'Text1.Height = Me.ScaleHeight
'Text1.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
form1.tlopen = False
Hide
End Sub

Private Sub List1_DblClick()
Dim p%, i%, tid$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
p% = InStr(List1.List(i%), "(ID:")
If p% = 0 Then Exit Sub
tid$ = trm(Mid(List1.List(i%), p% + 4))
Call auftritt.showrec(tid$, 0)

End Sub
