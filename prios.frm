VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form prios 
   Caption         =   "Prioritäten"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   720
      Picture         =   "prios.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   29
      ToolTipText     =   "aktualisieren"
      Top             =   4440
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid ptr2 
      Height          =   1095
      Index           =   2
      Left            =   7080
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid ptr2 
      Height          =   1095
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid ptr2 
      Height          =   1095
      Index           =   0
      Left            =   1560
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   240
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   14
      Left            =   8640
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   13
      Left            =   8040
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   12
      Left            =   7440
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   11
      Left            =   6840
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   10
      Left            =   6240
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   5640
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   5040
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   4440
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   3840
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   3240
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   2640
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   2040
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1440
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbx 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   840
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      ToolTipText     =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<--"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-->"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "prios.frx":0672
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Formular schliessen"
      Top             =   4440
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   8640
      Top             =   4440
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   3735
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6588
      _Version        =   393216
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   3735
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6588
      _Version        =   393216
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid fg2 
      Height          =   3735
      Index           =   2
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6588
      _Version        =   393216
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Werkzeuge"
      Height          =   855
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Label prio0 
      Caption         =   "Label3"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4215
      Left            =   240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "prios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pptr As Integer, currx As Integer, curry As Integer, execall(14), selid As String

Private Sub Command1_Click()
'd2infile = "prios": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command2_Click()
Dim cs$, i%

'd2infile = "prios": d2insub = "Command2_Click"
i% = Asc(prio0.Caption) + 1
If i% > 255 Then i% = 255
cs$ = Chr$(i%)
If cs$ > "X" Then cs$ = "X"
prio0.Caption = cs$
Call nulldsp
Call initdsp

End Sub

Public Sub Command20_Click()
Call nulldsp
Call initdsp

End Sub

Private Sub Command3_Click()
Dim cs$, i%

'd2infile = "prios": d2insub = "Command3_Click"
i% = Asc(prio0.Caption) - 1
If i% < 0 Then i% = 0
cs$ = Chr$(i%)
If cs$ < "A" Then cs$ = "A"
prio0.Caption = cs$
Call nulldsp
Call initdsp

End Sub

Private Sub fg2_Click(Index As Integer)
On Error Resume Next
selid = ptr2(Index).TextMatrix(curry, currx)
On Error GoTo 0
End Sub

Private Sub fg2_DblClick(Index As Integer)
Dim trgid As String, sid$, rrr

On Error Resume Next
trgid = ptr2(Index).TextMatrix(curry, currx)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
Select Case Left(trgid, 1)
  Case "A":
    sid$ = Mid(trgid, 3)
    shwAdrDetail.srchit% = 0
    Load shwAdrDetail
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(sid$, "")
    shwAdrDetail.Combo3.text = sid$
    Call shwAdrDetail.SetFocus
    shwAdrDetail.srchit% = 1
  Case "T":
    Load tplan
    Call tplan.rlists
    Call tplan.nulldsp
    Call tplan.showrec(Mid(trgid, 3))
    On Error Resume Next
    Call tplan.SetFocus
    On Error GoTo 0
  Case "E":
    Unload auftritt
    DoEvents
    Load auftritt
    Call auftritt.SetFocus
    Call auftritt.showrec(Mid(trgid, 3), 0)
  Case Else: Exit Sub
End Select
End Sub

Private Sub fg2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If fg2(Index).MouseRow >= fg2(Index).Rows Then Exit Sub

currx = fg2(Index).MouseCol
curry = fg2(Index).MouseRow
fg2(Index).col = currx
fg2(Index).Row = curry
fg2(Index).CellBackColor = RGB(0, 0, 0)

End Sub

Private Sub Form_Load()
'd2infile = "prios": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
selid = ""
prio0.Caption = "A"
Call nulldsp
Call initdsp
pptr = 0
form1.priosopen = True
Show
'Call initlinks
End Sub

Sub initlinks()
Dim i%, fn$, csl As cShellLink, j%, exf$, n As Integer, icf As String
Dim rrr
Dim ExeFile As String, WorkDir As String, ExeArgs As String, IconFile As String, IconIdx As Long, ShowCmd As Long
Dim descr As String
Dim d2infile, d2insub

d2infile = "prios": d2insub = "initlinks"
pptr = 0
For i% = 0 To 14: pbx(i%).Visible = False: Next i%
For i% = 0 To 14
  fn$ = form1.getusersetting("shelllink" + trm(i%), "")
  If fn$ <> "" Then
    If exist(fn$) Then
      Set csl = New cShellLink
      rrr = csl.GetShellLinkInfo(fn$, ExeFile, WorkDir, ExeArgs, IconFile, IconIdx, ShowCmd, descr)
      If rrr Then
        Call form1.setusersetting("shelllink" + trm(pptr), fn$)
        j% = InStr(ExeFile, Chr$(0))
        If j% > 1 Then exf$ = Left(ExeFile, j% - 1)
        If exf$ <> "" Then
          execall(pptr) = exf$
          If Left(IconFile, 1) <> Chr$(0) Then
            icf = IconFile
          Else
            icf = exf
          End If
          pbx(pptr).Picture = GetIconFromFile(icf, CLng(IconIdx), True)
          pbx(pptr).Visible = True
          fn$ = FileName(exf)
          pbx(pptr).ToolTipText = basename(fn$, FileExtension(fn$)) + ":" + descr
          DoEvents
          pptr = pptr + 1
        End If
      End If
    End If
  End If
Next i%
i% = pptr
While i% < 15
  Call form1.delusersetting("shellink" + trm(i%))
  i% = i% + 1
Wend
End Sub
Sub nulldsp()
Dim i%, X%

'd2infile = "prios": d2insub = "nulldsp"
For i% = 0 To 2
  Label1(i%).Caption = ""
  fg2(i%).Clear
  fg2(i%).Cols = 2
  fg2(i%).Rows = 1
  ptr2(i%).Rows = 1
  fg2(i%).ColWidth(0) = fg2(i%).Width / 5
  'For x% = 1 To fg2(i%).Cols - 1: fg2(i%).ColWidth(x%) = fg2(i%).ColWidth(x%) * 3 / 2: Next x%
  fg2(i%).ColWidth(1) = fg2(i%).Width - fg2(i%).ColWidth(0) - 60
Next i%

End Sub
Private Sub Form_Resize()
'd2infile = "prios": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "prios": d2insub = "Form_Unload"
form1.priosopen = False
Hide
On Error GoTo exuld1
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld1:
On Error GoTo 0

End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rrr, fn$, i%, csl As cShellLink, j%, c$, exf$, n As Integer, icf As String
Dim ExeFile As String, WorkDir As String, ExeArgs As String, IconFile As String, IconIdx As Long, ShowCmd As Long
Dim descr As String, fm$

'd2infile = "prios": d2insub = "Frame1_OLEDragDrop"
ExeFile = Space$(260)
WorkDir = Space$(260)
ExeArgs = Space$(260)
IconFile = Space$(260)
IconIdx = -1
ShowCmd = -1
descr = Space$(260)

On Error Resume Next
n = Data.Files.Count
rrr = Err
On Error GoTo 0
If rrr <> 0 Or pptr > 14 Then Exit Sub

For i% = 1 To n
If pptr < 15 Then
  fn$ = Data.Files(i%): fm$ = fn$
  exf$ = ""
  Set csl = New cShellLink
  rrr = csl.GetShellLinkInfo(fn$, ExeFile, WorkDir, ExeArgs, IconFile, IconIdx, ShowCmd, descr)
  If rrr Then
    Call form1.setusersetting("shelllink" + trm(pptr), fm$)
    j% = InStr(ExeFile, Chr$(0))
    If j% > 1 Then exf$ = Left(ExeFile, j% - 1)
    'Debug.Print exf$
    'Debug.Print IconIdx; ":"; IconFile
    If exf$ <> "" Then
      execall(pptr) = exf$
      If Left(IconFile, 1) <> Chr$(0) Then
        icf = IconFile
      Else
        icf = exf
      End If
      pbx(pptr).Picture = GetIconFromFile(icf, CLng(IconIdx), True)
      pbx(pptr).Visible = True
      fn$ = FileName(exf)
      pbx(pptr).ToolTipText = basename(fn$, FileExtension(fn$)) + ":" + descr
      pptr = pptr + 1
    End If
  End If
End If
Next i%

End Sub

Sub initdsp()
Dim rrr, bez As String
Dim r As ADODB.Recordset, r1 As ADODB.Recordset, c$, f2ptr As Integer, prvprio As String
Dim cp$, i%, s As String

Dim d2infile As String, d2insub As String
d2infile = "prios": d2insub = "initdsp"
s = prio0.Caption
Label1(0).Caption = s
For i% = 1 To 2
  Label1(i%).Caption = Chr$(Asc(Label1(i% - 1).Caption) + 1)
Next i%
c$ = "select evnt,prio from opt_prios where userid='" + form1.getuserid() + "'"
If s <> "" Then c$ = c$ + " and prio>='" + s + "'"
c$ = c$ + " order by prio;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
f2ptr = 0: prvprio = ""
While Not r.EOF And f2ptr < 3
  cp$ = trm(r!prio)
  If cp$ <> prvprio Then
    While cp$ <> Label1(f2ptr).Caption And f2ptr < 2
      f2ptr = f2ptr + 1
    Wend
    prvprio = cp
  End If
  If Label1(f2ptr).Caption = cp$ Then
    fg2(f2ptr).Rows = fg2(f2ptr).Rows + 1
    ptr2(f2ptr).Rows = fg2(f2ptr).Rows
    bez = ""
    Select Case UCase(Left(r!evnt, 1))
      Case "E": c$ = "select bezeichnung from auftritt where id='" + Mid(trm(r!evnt), 3) + "'"
                Set r1 = New ADODB.Recordset
                r1.CursorLocation = adUseServer
                On Error Resume Next
rrr = form1.adoopen(r1, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                bez = trm(r1!bezeichnung)
                On Error GoTo 0
      Case "T": bez = Mid(trm(r!evnt), 3)
      Case "A": bez = Mid(trm(r!evnt), 3)
      Case Else: bez = "Fehler:" + trm(r!evnt)
    End Select
    fg2(f2ptr).TextMatrix(fg2(f2ptr).Rows - 1, 1) = bez
    ptr2(f2ptr).TextMatrix(ptr2(f2ptr).Rows - 1, 1) = r!evnt
  End If
  r.MoveNext
Wend

End Sub

Private Sub pbx_DblClick(Index As Integer)
Dim rrx, r As ADODB.Recordset, sid$, c$, rrr
Dim d2infile As String, d2insub As String, anm$

d2infile = "prios": d2insub = "pbx_DblClick"
If selid <> "" Then
  sid$ = Mid(selid, 3)
  Select Case Left(selid, 1)
    Case "A":
      On Error Resume Next
      MkDir form1.s0dir() + "\" + form1.medien() + "\"
      c$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(sid$)
      MkDir c$
      On Error GoTo 0
    Case "T":
      On Error Resume Next
      MkDir form1.s0dir() + "\" + form1.medien() + "\"
      MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\"
      c$ = form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + form1.medienname(sid$)
      MkDir c$
      On Error GoTo 0
    Case "E":
      c$ = "select * from auftritt where id='" + sid$ + "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
      rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr <> 0 Then Exit Sub
      If Not r.EOF Then
        c$ = form1.medienname(r!tpid)
        anm$ = form1.medienname(form1.get_atabkz(trm(r!auftrittstyp) & "_" & sid$))
        On Error Resume Next
        MkDir form1.s0dir() + "\" + form1.medien() + "\"
        MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\"
        MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + c$
        c$ = form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + c$ & "\" & anm$
        MkDir c$
        On Error GoTo 0
      Else
        Exit Sub
      End If
  Case Else: Exit Sub
  End Select
  rrx = Shell(execall(Index), vbNormalFocus)
End If
End Sub

Private Sub pbx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
  Call form1.delusersetting("shelllink" + trm(Index))
  'Call initlinks
End If
End Sub

Private Sub pbx_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fn$, nf%, rrr, M$, rrx

'd2infile = "prios": d2insub = "pbx_OLEDragDrop"
M$ = "Index=" + trm(Index)
If Data.GetFormat(vbCFFiles) Then
  nf% = 1
  Do
    On Error Resume Next
    fn$ = Data.Files(nf%)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      M$ = M$ + vbCrLf + fn$
      rrx = Shell(execall(Index) + " " + fn$, vbNormalFocus)
      DoEvents
    End If
    nf% = nf% + 1
  Loop Until rrr <> 0
  'MsgBox M$
End If

End Sub
