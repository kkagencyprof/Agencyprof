VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form infobox 
   Caption         =   "Info"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      Picture         =   "infobox.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "alle exportieren"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4320
      Width           =   2535
   End
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
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Hilfeseite öffnen"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "infobox.frx":0250
      Top             =   2760
      Width           =   2775
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1800
      Left            =   120
      Picture         =   "infobox.frx":0256
      ScaleHeight     =   1740
      ScaleWidth      =   2700
      TabIndex        =   2
      Top             =   120
      Width           =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4920
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright (C) 2003  Karsten Kaus, Kontakt: info@AgencyProf.com, URL:http://www.AgencyProf.com"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
End
Attribute VB_Name = "infobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command18_Click()
Call form1.handbuchcall("15-Andere_Dinge.htm")

End Sub

Private Sub Command2_Click()
For i% = 0 To List1.ListCount - 1
  List1.ListIndex = i%
  DoEvents
  Call List1_DblClick
Next i%

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Call Command1_Click
End Sub


Private Sub Form_Load()
axsResizer1.SaveControlPositions
Show
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.Name)
Me.Left = form1.mylastleft(Me.Name)


Text1.Text = "Contains cryptography software by David Ireland of DI Management Services Pty Ltd <www.di-mgt.com.au>."

Label1.Caption = "Agencyprof " & App.Major & "." & App.Minor & " - Build #" & App.Revision & Chr$(13) & Chr$(10) & Label1.Caption

If InStr(LCase(App.EXEName), "apadmin") > 0 Then
  Command2.Enabled = False
  List1.Enabled = False
  Timer1.Enabled = False
Else
  Call Timer1_Timer
End If

End Sub


Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.Name, Me.Top)
Call form1.setmylastleft(Me.Name, Me.Left)
Hide
End Sub

Private Sub Label1_Click()
Call Picture1_Click
End Sub


Private Sub List1_DblClick()
MsgBox App.EXEName
i% = List1.ListIndex
If i% < 0 Then Exit Sub

MousePointer = 11: DoEvents
Call form1.ExportOneTableToExcel(word1(List1.List(i%)))
MousePointer = 0

End Sub


Private Sub Picture1_Click()
Unload frmBrowser
frmBrowser.StartingAddress = "http://www.agencyprof.de/site/contact.html"
Load frmBrowser

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
Dim r As Recordset

List1.Clear
For i% = 0 To form1.sqla.TableDefs.Count - 1
  If Left$(LCase(sqla.TableDefs(i%).Name), 4) <> "msys" Then
    Set r = form1.sqla.OpenRecordset( _
      "SELECT count(*) as cnt FROM " + form1.sqla.TableDefs(i%).Name, dbOpenDynaset, dbReadOnly)

    ad$ = form1.sqla.TableDefs(i%).Name & " " & r!cnt & " recs"
    'Debug.Print form1.sqla.TableDefs(i%).Indexes(0)
    List1.AddItem ad$
  End If
Next i%

End Sub
