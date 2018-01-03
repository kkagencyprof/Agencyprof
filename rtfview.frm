VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form rtfview 
   Caption         =   "Dateiansicht"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows-Standard
   Begin Resizer.axsResizer axsResizer1 
      Left            =   2520
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Drucken"
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   7080
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Editor"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      MaskColor       =   &H00000000&
      Picture         =   "rtfview.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Speichern"
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "rtfview.frx":066C
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Formular schiessen"
      Top             =   5760
      Width           =   495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"rtfview.frx":11EE
   End
End
Attribute VB_Name = "rtfview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

d2infile = "rtfview": d2insub = "Command1_Click"
Call chksave
Unload rtfview

End Sub

Private Sub Command2_Click()
Dim ext$
d2infile = "rtfview": d2insub = "Command2_Click"
Call chksave
ext$ = FileExtension(form1.fixfilename(RichTextBox1.FileName))
On Error Resume Next
X = Shell(form1.getmyeditor(ext$) & " " & form1.fixfilename(RichTextBox1.FileName), 1)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Editor kann nicht gestartet werden."
Else
  Call Command1_Click
End If

End Sub

Private Sub Command3_Click()

d2infile = "rtfview": d2insub = "Command3_Click"
Call cdlg1.ShowPrinter

End Sub

Private Sub Command4_Click()

d2infile = "rtfview": d2insub = "Command4_Click"
RichTextBox1.SaveFile (RichTextBox1.FileName)
Command4.Enabled = False
BackColor = form1.cleancolor()

End Sub

Private Sub Form_Load()
d2infile = "rtfview": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Command4.Enabled = False
Show

End Sub

Private Sub Form_Resize()
d2infile = "rtfview": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "rtfview": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub
Public Sub loadtext(fn$)

d2infile = "rtfview": d2insub = "loadtext"
RichTextBox1.FileName = fn$
rtfview.Caption = fn$
Command4.Enabled = False
BackColor = form1.cleancolor()

End Sub

Private Sub RichTextBox1_Change()
d2infile = "rtfview": d2insub = "RichTextBox1_Change"
Command4.Enabled = True
BackColor = form1.dirtycolor()

End Sub

Sub chksave()
d2infile = "rtfview": d2insub = "chksave"
If BackColor = form1.cleancolor() Then Exit Sub
If form1.ask2save() = vbNo Then Exit Sub
Call Command4_Click

End Sub
