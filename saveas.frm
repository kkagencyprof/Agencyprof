VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form saveas 
   Caption         =   "Datei speichern unter..."
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   ScaleHeight     =   1815
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   4800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Height          =   495
      Left            =   4440
      Picture         =   "saveas.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Dateiauwahlbox - immer öffnen: saveas=comdlg in den Benutzereinstellungen"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "temp.rtf"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox verzname 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox dateiname 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abbruch"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok, speichern"
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox fname 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Datei"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Verzeichnis"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vorschlag"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "saveas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim noupcombo As Boolean, noupfn As Boolean
Dim fl_rl3%
Dim SelOK As Boolean

Public Property Get SelectionOK() As Boolean
  SelectionOK = SelOK
End Property

Public Property Get SelectedName() As String
  SelectedName = fname.Text
End Property

Private Sub Command1_Click()
'd2infile = "saveas": d2insub = "Command1_Click"
SelOK = True
Hide
End Sub

Private Sub Command2_Click()
'd2infile = "saveas": d2insub = "Command2_Click"
SelOK = False
Hide
End Sub

Private Sub Command3_Click()
'd2infile = "saveas": d2insub = "Command3_Click"
dateiname.Text = "temp.rtf"
If form1.vorlagencache <> "" Then verzname.Text = form1.vorlagencache
DoEvents
SelOK = True
Hide
End Sub

Private Sub Command4_Click()

'd2infile = "saveas": d2insub = "Command4_Click"
    On Error Resume Next
    With cdlg1
    'Bei "Abbruch" Fehler raisen lassen:
    .CancelError = True
    'Suchpfad einstellen:
    .InitDir = DirName(fname.Text)
    .FileName = FileName(fname.Text)
    .DialogTitle = "Speichern unter ..."
    'und endlich den Dialog anzeigen:
    .ShowOpen

    'Auswertung:
    If Err = cdlCancel Then
      On Error GoTo 0
      Call Command2_Click
      Exit Sub
    End If
    On Error GoTo 0
    Call Command1_Click
    fname.Text = .FileName

    End With
    On Error GoTo 0


End Sub

Private Sub dateiname_Change()
'd2infile = "saveas": d2insub = "dateiname_Change"
Call fnupd
End Sub

Private Sub fname_Change()
'd2infile = "saveas": d2insub = "fname_Change"
If noupfn = False Then
  noupcombo = True
  verzname.Text = DirName(trm(fname.Text))
  dateiname.Text = FileName(trm(fname.Text))
  noupcombo = False
End If

End Sub

Private Sub Form_Load()
'd2infile = "saveas": d2insub = "Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Me.Caption = form1.inmylanguage("saveas")
Call form1.formpos(Me)
Label1.Caption = form1.inmylanguage("savorschlag")
Label2.Caption = form1.inmylanguage("Verzeichnis")
Label3.Caption = form1.inmylanguage("Datei")
Command1.Caption = form1.inmylanguage("Ok, speichern")
Command2.Caption = form1.inmylanguage("Abbruch")

SelOK = False
noupcombo = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
'd2infile = "saveas": d2insub = "Form_QueryUnload"
  If (UnloadMode = vbFormControlMenu) Then
    Cancel = True: Hide
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'd2infile = "saveas": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0


End Sub


Private Sub fnupd()
Dim mename$
'd2infile = "saveas": d2insub = "fnupd"
If noupcombo = False Then
  noupfn = True
  mename$ = trm(verzname.Text)
  If Right$(mename$, 1) = "\" Then mename$ = Left(mename$, Len(mename$) - 1)
  fname.Text = mename$ & "\" & trm(dateiname.Text)
  noupfn = False
End If

End Sub


Private Sub verzname_Change()
'd2infile = "saveas": d2insub = "verzname_Change"
Call fnupd
End Sub

Private Sub verzname_Click()
'd2infile = "saveas": d2insub = "verzname_Click"
Call verzname_Change
End Sub

Private Sub verzname_DropDown()
Dim rtmp As ADODB.Recordset
Dim c$, rrr
Dim uId$

Dim d2infile As String, d2insub As String
d2infile = "saveas": d2insub = "verzname_DropDown"
verzname.Clear

uId$ = form1.getuserid()
verzname.AddItem form1.s0dir() & "\" & form1.docs() & "\" & uId$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_DokumentenVerzeichnis')=1 or instr(owner,'sysvar_system_DokumentenVerzeichnis')=1", form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  'o$ = rtmp!Owner
  'o$ = Mid$(o$, InStr(o$, uId$) + Len(uId$) + 1)
  verzname.AddItem rtmp!wert
  rtmp.MoveNext
Wend

End Sub
