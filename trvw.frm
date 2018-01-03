VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form trvw 
   Caption         =   "TreeViewer"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox alwopen 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      Picture         =   "trvw.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "löschen"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      Picture         =   "trvw.frx":12D6
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Lege neue Nachricht an"
      Top             =   720
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   4200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      Picture         =   "trvw.frx":1942
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Ihr Dokumentenverzeichnis öffnen"
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   3855
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6800
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "trvw.frx":1F6C
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   4080
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3480
      Top             =   720
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label alwopenl 
      BackStyle       =   0  'Transparent
      Caption         =   "immer öffnen"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label currnod 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4695
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "trvw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currmode As String

Private Sub alwopen_Click()

If alwopen.value = 0 Then
  Call form1.setmylastFormVar(Me.name, "mailimmertreeview", "0")
Else
  Call form1.setmylastFormVar(Me.name, "mailimmertreeview", "1")
End If

End Sub

Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command19_Click()
Dim X

X = Shell("explorer.exe " & currnod.Caption, vbNormalFocus)

End Sub

Private Sub Command5_Click()
Dim neuv$, nnn, rrr

neuv$ = InputBox("Neues Verzeichnis erstellen", "Verzeichnisname", "")
If neuv$ <> "" Then
  neuv$ = strrepl(neuv$, " ", "_")
  neuv$ = strrepl(neuv$, "'", "")
  On Error Resume Next
  MkDir currnod.Caption + "\" + neuv$
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    nnn = tv1.Nodes.add(currnod.Caption, tvwChild, currnod.Caption + "\" + neuv$, neuv$)
  Else
    MsgBox transe("Das Verzeichnis") + " " + neuv$ + vbCrLf + "(" + currnod.Caption + "\" + neuv$ + ")" + vbCrLf + transe("kann nict erstellt werden.")
  End If
End If

End Sub

Private Sub delme_Click()
Dim tr, nnn As Node, pfad As String, rrr

pfad = currnod.Caption + "\"
tr = Dir(pfad + "*.*")
If tr = "" Then
  tr = Dir(pfad, vbDirectory)
  Do While tr <> ""
    If (GetAttr(pfad + tr) And vbDirectory) = vbDirectory Then
      If tr <> "." And tr <> ".." Then
        GoTo delme_nodel
      End If
    End If
    tr = Dir
  Loop
  On Error Resume Next
  RmDir pfad
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    tv1.Nodes.Remove tv1.SelectedItem.Index
    Exit Sub
  End If
End If
delme_nodel:
MsgBox transe("Das Verzeichnis kann nicht gelöscht werden")

End Sub

Private Sub Form_Load()
Dim klrv%
currmode = ""
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
pg1.Visible = False
klrv% = Val(form1.mylastFormVar(Me.name, "mailimmertreeview", "0"))
If klrv% <> 0 Then klrv% = 1
alwopen.value = klrv%
Show

End Sub
Private Sub Form_Resize()

axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case currmode:
  Case "mozillamail"
    frmMain.mozopen = False
    Call frmMain.cmdCheckMail_Click
    frmMain.cmdView.Enabled = True
    frmMain.cmdCheckMail.Enabled = True
  Case "mail"
    msafe.Command5.Enabled = True
    msafe.Command6.Enabled = False
    msafe.Caption = transe("Mailsafe")
  Case Else
End Select

Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Sub rtree1(mde$)
Dim nod As Node
Dim t0$, t$, tr, i%, liste As String, l$, w$, getlist As String, w1$
Dim pgmx, pgm, abst%, trx$
Dim lvitem

Select Case mde$
  Case "mozillamail"
    MousePointer = 11: DoEvents
    Me.Caption = "Mailverzeichnisse"
    t0$ = form1.getusersetting("Netscape47inbox", "")
    If t0$ <> "" Then
      t0$ = DirName(t0$)
      t$ = t0$
      Set nod = tv1.Nodes.add(, , t$, FileName(t$))
      nod.Expanded = True
      frmMain.listMessages.ListItems.Clear
      currnod.Caption = t$
      tr = Dir(t$ + "\*.*", vbDirectory)
      Do While tr <> ""
        If (GetAttr(t$ + "\" + tr) And vbDirectory) = vbDirectory Then
          If tr <> "." And tr <> ".." Then
            Set nod = tv1.Nodes.add(t$, tvwChild, t$ + "\" + tr, tr)
          End If
        Else
          trx$ = LCase(Right(tr, 4))
          If trx$ <> ".msf" And trx$ <> ".dat" Then
            Set lvitem = frmMain.listMessages.ListItems.add(, , tr)
          End If
        End If
        tr = Dir
      Loop
    End If
  Case "mail"
    msafe.ftrg.Clear
    MousePointer = 11
    pg1.value = 0
    pg1.Max = 100
    pg1.Visible = True
    DoEvents
    Me.Caption = "Mailverzeichnisse"
    t0$ = form1.mydir()
    t$ = t0$
    Set nod = tv1.Nodes.add(, , t$ + "\mail", "mail")
    nod.Expanded = True
    getlist = t$ + "\mail"
    abst% = Len(t$ + "\mail") + 2
    Do
      w$ = cut_d1(getlist, "|")
      getlist = cut_d2bis(getlist, "|")
      If w$ <> "" Then
        liste = dirlist(w$ + "\")
        Do
          w1$ = cut_d1(liste, "|")
          liste = cut_d2bis(liste, "|")
          If w1$ <> "" Then
            If getlist <> "" Then getlist = getlist + "|"
            getlist = getlist + w$ + "\" + w1$
            Set nod = tv1.Nodes.add(w$, tvwChild, w$ + "\" + w1$, w1$)
            msafe.ftrg.AddItem Mid$(w$ + "\" + w1$, abst%)
'            DoEvents
          End If
          pgm = Len(getlist + liste)
          If pgmx < pgm Then
            pgmx = pgm
            pg1.Max = pgmx
          End If
          pg1.value = pgmx - pgm
        Loop Until liste = ""
      End If
      pgm = Len(getlist)
      If pgmx < pgm Then
        pgmx = pgm
        pg1.Max = pgmx
      End If
      pg1.value = pgmx - pgm
      DoEvents
    Loop Until getlist = ""
  Case Else
End Select
MousePointer = 0
pg1.Visible = False

End Sub
Public Sub setmode(mode$)
currmode = mode$
Call rtree1(mode$)
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim t0$, abst%, tr, t$, nod As Node, trx$, lvitem

MousePointer = 0
Select Case currmode
  Case "mozillamail":
    t0$ = form1.getusersetting("Netscape47inbox", "")
    If t0$ <> "" Then
      t$ = DirName(DirName(t0$)) + "\" + Node.FullPath
      currnod.Caption = t$
      Node.Expanded = True
      frmMain.listMessages.ListItems.Clear
      tr = Dir(t$ + "\*.*", vbDirectory)
      Do While tr <> ""
        If (GetAttr(t$ + "\" + tr) And vbDirectory) = vbDirectory Then
          If tr <> "." And tr <> ".." Then
            On Error Resume Next
            Set nod = tv1.Nodes.add(t$, tvwChild, t$ + "\" + tr, basename(trm(tr), ".sbd"))
            On Error GoTo 0
          End If
        Else
          trx$ = LCase(Right(tr, 4))
          If trx$ <> ".msf" And trx$ <> ".dat" Then
            Set lvitem = frmMain.listMessages.ListItems.add(, , tr)
          End If
        End If
        tr = Dir
      Loop
    End If
  Case "mail":
    t0$ = form1.mydir() + "\"
    abst% = Len(t0$ + "mail") + 2
    currnod.Caption = t0$ + Node.FullPath
    msafe.Command6.ToolTipText = transe("verschiebt markierte Mails nach") + " " + currnod.Caption
    msafe.Command6.Enabled = True
    msafe.Caption = currnod.Caption
    msafe.ftrg.text = Mid(currnod.Caption, abst%)
    DoEvents
    Call msafe.rlist1(currnod.Caption)
  Case Else:
End Select

End Sub

