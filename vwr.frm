VERSION 5.00
Begin VB.Form vwr 
   AutoRedraw      =   -1  'True
   Caption         =   "Viewer"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Image i1 
      Height          =   4935
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "vwr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rzno As Integer
Private Sub Form_Load()

'd2infile = "vwr": d2insub = "Form_Load"
rzno = 1
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
rzno = 0

Show
End Sub


Private Sub Form_Resize()
Dim fak As Double

'd2infile = "vwr": d2insub = "Form_Resize"
If rzno = 1 Then Exit Sub
tbw = 100
wmx = Screen.Width - tbw - Left
fak = i1.Width / i1.Height
rzno = 1
Width = fak * Height
If Width + Left > wmx Then
  'Width = wmx - Left
  Width = wmx
  Height = Width / fak
End If
i1.Width = Width
i1.Height = Height
i1.Stretch = True
'If exist(Caption) <> 0 Then i1 = LoadPicture(Caption)

rzno = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "vwr": d2insub = "Form_Unload"
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

Hide
End Sub

Public Sub rdrw()

rzno = 1
fn$ = Caption
If exist(fn$) <> 0 Then
  i1.Stretch = False
  i1 = LoadPicture(fn$)
  Width = i1.Width
  i0h = i1.Height
  h1 = Me.Height
  Height = Max(0, Int(h1))
  Height = Max(Height, (Screen.Height - 0.05 * Screen.Height) * 0.8)
  i1.Height = Height - i1.Top
  fx = i1.Height / i0h
  i1.Width = i1.Width * fx
  Width = i1.Width
  i1.Stretch = True
End If
rzno = 0
End Sub
Public Sub mxh()

'd2infile = "vwr": d2insub = "mxh"
Height = Screen.Height - 0.05 * Screen.Height


End Sub

