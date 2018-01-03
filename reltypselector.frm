VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form reltypselector 
   Caption         =   "Adresstyp(en) auswählen"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   120
      Picture         =   "reltypselector.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Formular schiessen"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "neu"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   15
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "löschen"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "wählen"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      Caption         =   "neuer Typ:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   15
      Width           =   855
   End
End
Attribute VB_Name = "reltypselector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
form1.d2infile="reltypselector": form1.d2insub="Command1_Click"
Hide
End Sub

Private Sub Command2_Click()
form1.d2infile="reltypselector": form1.d2insub="Command2_Click"
If List1.ListIndex >= 0 Then Call List1_DblClick
End Sub

Private Sub Command3_Click()

form1.d2infile="reltypselector": form1.d2insub="Command3_Click"
If List1.ListIndex < 0 Then Exit Sub

id$ = List1.List(List1.ListIndex)
form1.sqlqry ("delete from adresstypen where id='" & id$ & "'")
Call rlist1

End Sub

Private Sub Command4_Click()

form1.d2infile="reltypselector": form1.d2insub="Command4_Click"
If trm(Text1.Text) = "" Then Exit Sub

id$ = trm(Text1.Text)
form1.sqlqry ("insert into adresstypen (id) values('" & id$ & "')")
Call rlist1

End Sub

Private Sub Form_Load()


form1.d2infile="reltypselector": form1.d2insub="Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
axsResizer1.SaveControlPositions


reltypselector.Caption = transe("Adresstyp(en) auswählen")
Command1.ToolTipText = transe("Formular schiessen")
Command4.Caption = transe("neu")
Command3.Caption = transe("löschen")
Command2.Caption = transe("wählen")
Label1.Caption = transe("neuer Typ:")
Show
Call rlist1

End Sub

Sub rlist1()
Dim rtmp As ADODB.Recordset

form1.d2infile="reltypselector": form1.d2insub="rlist1"
List1.Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT id FROM adresstypen", form1.adoc, adOpenDynamic, adLockReadOnly

While Not rtmp.EOF
  List1.AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend
Text1.Text = ""

End Sub

Private Sub Form_Resize()
form1.d2infile="reltypselector": form1.d2insub="Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
form1.d2infile="reltypselector": form1.d2insub="Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub

Private Sub List1_DblClick()

form1.d2infile="reltypselector": form1.d2insub="List1_DblClick"
If List1.ListIndex < 0 Then Exit Sub

Call shwAdrDetail.addtyp(transo(List1.List(List1.ListIndex)))

End Sub
