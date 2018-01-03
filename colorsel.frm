VERSION 5.00
Begin VB.Form colorsel 
   Caption         =   "Farbe auswählen"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Alignment       =   2  'Zentriert
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
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
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
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   2760
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Abbruch"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   1440
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "colorsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currcol As Long

Private Sub Command1_Click()
d2infile = "colorsel": d2insub = "Command1_Click"
Call form1.setcolorselected(-1)
Unload colorsel

End Sub

Private Sub Command2_Click()

d2infile = "colorsel": d2insub = "Command2_Click"
Call form1.setcolorselected(currcol)

Unload colorsel
End Sub

Private Sub Form_Load()
d2infile = "colorsel": d2insub = "Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

Show
Picture1.ScaleHeight = 255
Picture1.ScaleWidth = 255
Picture2.ScaleHeight = 255
Picture2.ScaleWidth = 255
Picture3.ScaleHeight = 255
Picture3.ScaleWidth = 255
Picture4.ScaleHeight = 255
Picture4.ScaleWidth = 255

Call form1.setcolorselected(-20)

For x = 0 To 250 Step 5
  For y = 0 To 250 Step 5
    Picture1.Line (x, y)-(x + 5, y + 5), RGB(x, y, 0), BF
    Picture2.Line (x, y)-(x + 5, y + 5), RGB(x, 0, y), BF
    Picture3.Line (x, y)-(x + 5, y + 5), RGB(0, x, y), BF
  Next y
  Picture4.Line (x, 0)-(x + 5, 255), RGB(x, x, x), BF
Next x
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "colorsel": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Public Sub updc(c As Long)
Dim w As Long, r As Long, g As Long, b As Long

d2infile = "colorsel": d2insub = "updc"
If c <= 0 Then Exit Sub
Picture5.BackColor = c
Text1.ForeColor = c
b = c / 65536
w = c Mod 65536
g = w / 256
r = w Mod 256
Text1.Text = r & " /" & g & " /" & b
Text2.Text = trm(Text1.ForeColor)

End Sub

Private Sub Picture1_Click()
d2infile = "colorsel": d2insub = "Picture1_Click"
Call Command2_Click
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

d2infile = "colorsel": d2insub = "Picture1_MouseMove"
If Shift = 0 Then
  currcol = Picture1.Point(x, y)
  Call updc(currcol)
End If

End Sub

Private Sub Picture2_Click()
d2infile = "colorsel": d2insub = "Picture2_Click"
Call Command2_Click
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
d2infile = "colorsel": d2insub = "Picture2_MouseMove"
If Shift = 0 Then
  currcol = Picture2.Point(x, y)
  Call updc(currcol)
End If
End Sub

Private Sub Picture3_Click()
d2infile = "colorsel": d2insub = "Picture3_Click"
Call Command2_Click
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

d2infile = "colorsel": d2insub = "Picture3_MouseMove"
If Shift = 0 Then
  currcol = Picture3.Point(x, y)
  Call updc(currcol)
End If

End Sub

Private Sub Picture4_Click()
d2infile = "colorsel": d2insub = "Picture4_Click"
Call Command2_Click
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

d2infile = "colorsel": d2insub = "Picture4_MouseMove"
If Shift = 0 Then
  currcol = Picture4.Point(x, y)
  Call updc(currcol)
End If

End Sub
