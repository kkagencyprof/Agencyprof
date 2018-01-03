VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "resizer.ocx"
Begin VB.Form dialselect 
   Caption         =   "Telefonnummern "
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command35 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      MaskColor       =   &H00000000&
      Picture         =   "dialselect.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Markierte Nummer wählen"
      Top             =   3360
      Width           =   495
   End
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
      Picture         =   "dialselect.frx":018A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   3360
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ListBox List1 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Doppelklick waehlt die Nummer"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label dialthis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "dialselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cvid$

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command35_Click()
Dim l$

l$ = trm(dialthis.Caption)
If l$ <> "" Then
  Load AutoAnwahl
  Call AutoAnwahl.SetFocus
  AutoAnwahl.nummer.text = l$
  DoEvents
  Call AutoAnwahl.cmdDial_Click
  Call Command11_Click
End If
End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Caption = transe("Telefonnummern")
Call form1.formpos(Me)
Command35.ToolTipText = transe("Markierte Nummer wählen")
Command11.ToolTipText = transe("Dieses Formular schliessen")
List1.ToolTipText = transe("Doppelklick waehlt die Nummer")
cvid$ = ""
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

Private Sub List1_Click()
Dim i%, d$, z$, vid$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
dialthis.Caption = ""
d$ = trm(cut_d2bis(List1.List(i%), ":"))
vid$ = ""
If InStr(d$, ":") > 0 Then
  vid$ = trm(cut_d2bis(d$, ":"))
  d$ = trm(cut_d1(d$, ":"))
End If
cvid$ = vid$
If Left(d$, 1) = "~" Then dialthis.Caption = "~"
If Left(d$, 1) = "+" Then dialthis.Caption = "00"
i% = 1
While i% <= Len(d$) And i% > 0
  z$ = Mid$(d$, i%, 1)
  If istnum(z$) > 0 Then
    dialthis.Caption = dialthis.Caption & z$
  End If
  i% = i% + 1
Wend
End Sub

Private Sub List1_DblClick()
DoEvents
Call Command35_Click
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c$, i%

i% = List1.ListIndex
If i% < 0 Then Exit Sub

If KeyCode = 8 Or KeyCode = 46 Then
  If cvid$ = "" Then
    MsgBox (transe("nicht löschbar"))
  Else
    c$ = "delete from opt_numbers where id='" + cvid$ + "'"
    Call form1.sqlqry(c$)
    List1.RemoveItem i%
  End If
End If
End Sub
