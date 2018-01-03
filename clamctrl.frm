VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form clamctrl 
   Caption         =   "Clamav Virenscanner"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command18 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "nur Viren melden (--infected)"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Unterverzeichnisse durchsuchen (-r)"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      Picture         =   "clamctrl.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Datei scannen"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   5160
      Picture         =   "clamctrl.frx":117A
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Verzeichnis scannen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   840
      Picture         =   "clamctrl.frx":22F4
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Starte Freshclam"
      Top             =   3120
      Width           =   615
   End
   Begin VB.ListBox filelist 
      Height          =   2355
      IntegralHeight  =   0   'False
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.ListBox dirlist 
      Height          =   2355
      IntegralHeight  =   0   'False
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "clamctrl.frx":2E5A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Formular schliessen"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.ListBox drvlist 
      Height          =   2355
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   2280
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dateien"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Verzeichnisse"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Laufwerk"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "clamctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d2infile = "clamctrl": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command18_Click()
d2infile = "clamctrl": d2insub = "Command18_Click"
Call form1.handbuchcall("03-Benutzereinstellungen.htm")

End Sub

Private Sub Command2_Click()
Dim fn As String
Dim i As Integer

d2infile = "clamctrl": d2insub = "Command2_Click"
i = filelist.ListIndex
If i < 0 Then Exit Sub

fn = Text1.Text
If Right(fn, 1) <> "\" Then fn = fn + "\"
fn = fn + filelist.List(i)
xx$ = form1.clamscanfile(fn)
If xx$ <> "0" Then
  MsgBox transe("Mindestens ein Virus gefunden in") + " " + fn
Else
  MsgBox transe("Kein Virus gefunden in") + " " + fn
End If

End Sub

Private Sub Command3_Click()
Dim fn As String

d2infile = "clamctrl": d2insub = "Command3_Click"
fn = Chr$(34) + Text1.Text + Chr$(34)
If Check1.value = 1 Then fn = "-r " + fn
If Check2.value = 1 Then fn = "--infected " + fn
xx$ = form1.clamscandir(fn)

End Sub

Private Sub Command4_Click()
d2infile = "clamctrl": d2insub = "Command4_Click"
MousePointer = 11: DoEvents
Call form1.freshclam(1)
MousePointer = 0
End Sub

Private Sub dirlist_DblClick()
Dim i As Integer, e As String

d2infile = "clamctrl": d2insub = "dirlist_DblClick"
i = dirlist.ListIndex
If i < 0 Then Exit Sub

e = dirlist.List(i)
If e = ".." Then
  e = DirName(Text1.Text)
  Text1.Text = e
Else
  If Right(Text1.Text, 1) <> "\" Then e = "\" + e
  Text1.Text = Text1.Text + e
End If

End Sub

Private Sub drvlist_DblClick()
Dim i As Integer

d2infile = "clamctrl": d2insub = "drvlist_DblClick"
i = drvlist.ListIndex
If i < 0 Then Exit Sub

Text1.Text = drvlist.List(i)


End Sub

Private Sub filelist_Click()
d2infile = "clamctrl": d2insub = "filelist_Click"
Command2.Enabled = True
End Sub

Private Sub Form_Load()
Dim drvl As String

d2infile = "clamctrl": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
drvl = trm(GetDriveStrings())
While drvl <> ""
  drvlist.AddItem cut_d1(drvl, Chr$(0))
  drvl = trm(cut_d2bis(drvl, Chr$(0)))
Wend
Show
End Sub
Private Sub Form_Resize()
d2infile = "clamctrl": d2insub = "Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "clamctrl": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub Text1_Change()
Dim tr, t$, rrr

d2infile = "clamctrl": d2insub = "Text1_Change"
dirlist.Clear
filelist.Clear
Command2.Enabled = False
t$ = Text1.Text
On Error Resume Next
tr = Dir(t$ + "\*.*", vbDirectory)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then tr = ""
Do While tr <> ""
  If (GetAttr(t$ + "\" + tr) And vbDirectory) = vbDirectory Then
    If tr <> "." Then
      dirlist.AddItem tr
    End If
  Else
    filelist.AddItem tr
  End If
  tr = Dir
Loop

End Sub
