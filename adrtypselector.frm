VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form adrtypselector 
   Caption         =   "Adresstyp(en) auswählen"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "neu"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   120
      Picture         =   "adrtypselector.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Formular schiessen"
      Top             =   4080
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
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "wählen"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
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
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label4 
      Caption         =   "neue Bez.:"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   15
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Beziehungen"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Adressgruppen"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1335
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
Attribute VB_Name = "adrtypselector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Hide
End Sub

Private Sub Command2_Click()
If List1.ListIndex >= 0 Then
  Call List1_DblClick
Else
  If List2.ListIndex >= 0 Then Call List2_DblClick
End If
End Sub

Private Sub Command3_Click()
Dim id$

If List1.ListIndex >= 0 Then
  id$ = List1.List(List1.ListIndex)
  form1.sqlqry ("delete from adresstypen where id='" & id$ & "'")
  Call rlist1
Else
  If List2.ListIndex >= 0 Then
    id$ = List2.List(List2.ListIndex)
    form1.sqlqry ("delete from adresstypen where id='rel:" & id$ & "'")
    Call rlist1
  End If
End If

End Sub

Private Sub Command4_Click()
Dim id$

If trm(Text1.Text) = "" Then Exit Sub

id$ = trm(Text1.Text)
form1.sqlqry ("insert into adresstypen (id) values('" & id$ & "')")
Text1.Text = ""
Call rlist1

End Sub

Private Sub Command5_Click()
Dim id$

If trm(Text2.Text) = "" Then Exit Sub

id$ = trm(Text2.Text)
form1.sqlqry ("insert into adresstypen (id) values('rel:" & id$ & "')")
Text2.Text = ""
Call rlist1

End Sub

Private Sub Form_Load()


Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
axsResizer1.SaveControlPositions

adrtypselector.Caption = transe("Adresstyp(en) auswählen")
Command1.ToolTipText = transe("Formular schliessen")
Command4.Caption = transe("neu")
Command3.Caption = transe("löschen")
If form1.getusersetting("adresstypschutz", "nein") <> "nein" Then
  Command3.Enabled = False
End If
Command2.Caption = transe("wählen")
Label1.Caption = transe("neuer Typ:")
Label4.Caption = transe("neue Bez.:")
Label2.Caption = transe("Adressgruppen")
Label3.Caption = transe("Beziehungen")
Show
Call rlist1

End Sub

Sub rlist1()
Dim rtmp As ADODB.Recordset, rrr
Dim d2infile As String, d2insub As String
d2infile = "adrtypselector": d2insub = "rlist1"
List1.Clear
List2.Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresstypen", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
While Not rtmp.EOF
  If Left(rtmp!id, 4) <> "rel:" Then
    List1.AddItem transe(rtmp!id)
  Else
    List2.AddItem Mid(rtmp!id, 5)
  End If
  rtmp.MoveNext
Wend
Text1.Text = ""

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

Private Sub List1_Click()
Dim i As Integer

i = List1.ListIndex
If i < 0 Then Exit Sub

List2.ListIndex = -1
List1.ListIndex = i
DoEvents
End Sub

Private Sub List1_DblClick()

If List1.ListIndex < 0 Then Exit Sub
Call shwAdrDetail.addtyp(transo(List1.List(List1.ListIndex)))

End Sub

Private Sub List2_Click()
Dim i As Integer

i = List2.ListIndex
If i < 0 Then Exit Sub

List1.ListIndex = -1
List2.ListIndex = i
DoEvents

End Sub

Private Sub List2_DblClick()
Dim neukwert As String, neuawert As String, neuwert As String
If List2.ListIndex < 0 Then Exit Sub

Unload bezlist
Load adrselect
Call adrselect.sel_init("", "")
Call adrselect.SetFocus
Do
  DoEvents
Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
If adrselect.sel_brk() = 0 Then
  neukwert = adrselect.get_kontsel()
  neuwert = adrselect.sel_getselected(): neuawert = neuwert
  If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
  Call shwAdrDetail.addrel(List2.List(List2.ListIndex), neuwert)
  Unload adrselect
End If

End Sub
