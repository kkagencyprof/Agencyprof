VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form dbupgrade 
   Caption         =   "Datenbankupgrade"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "dbupgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Unload dbupgrade
End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
dbupgrade.Caption = transe("Datenbankupgrade")
Command1.Caption = transe("&Schliessen")
Show
End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Public Sub addline(text As String)
  List1.AddItem text
  List1.ListIndex = List1.ListCount - 1
End Sub
