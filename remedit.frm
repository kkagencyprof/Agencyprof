VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "resizer.ocx"
Begin VB.Form remedit 
   Caption         =   "Edit Reminder"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "remedit.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "exit"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      Picture         =   "remedit.frx":0250
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "save, exit"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox rmytext 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox rdatum 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox roffset 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   615
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label rdeftext 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "Date:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Offset:"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.Label abez 
      Caption         =   "..."
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label aid 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label remid 
      Caption         =   "..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "remedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dtgedit As Boolean, offedit As Boolean

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
Dim c$, rd$, neuw$, csid$

csid$ = trm(remid.Caption)
If csid$ <> "" Then
  If rdeftext.Visible Then
    neuw$ = Trim("" & rmytext.text)
    If neuw$ <> "" Then
      c$ = "update opt_checks set confirmed='ok, deleted' where id='" + csid$ + "'"
      Call form1.sqlqry(c$)
      csid$ = form1.newid("opt_checks", "id", 22)
      c$ = "insert into opt_checks (id,auftrittsid) values('" + csid$ + "','" + aid.Caption + "')"
      Call form1.sqlqry(c$)
      c$ = "update opt_checks set checkid='',checkpoint='" + neuw$ + "' where id='" + csid$ + "'"
      Call form1.sqlqry(c$)
    End If
  End If
  c$ = "update opt_checks set dtg='" + datum2sql(trm(rdatum.text)) + "' where id='" + csid$ + "'"
  Call form1.sqlqry(c$)
  If form1.auftrittisopen Then Call auftritt.shw_reminders
  If form1.todoisopen Then Call todolist.Command4_Click
End If
Command10.Enabled = False
End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
dtgedit = False: offedit = False
Call form1.formpos(Me)


End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub

Private Sub rdatum_Change()
Dim diffd, rrr, dtg

Command10.Enabled = True
If dtgedit Then
  On Error Resume Next
  dtg = CDate(rdatum.text)
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    diffd = dtg - CDate(trm(cut_d1(abez.Caption, ",")))
    roffset.text = trm(diffd)
  End If
End If
End Sub

Private Sub rdatum_DblClick()
Dim id$

  With frmCalendar
    .Init rdatum, rdatum.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      rdatum.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With

End Sub

Private Sub rdatum_GotFocus()
dtgedit = True: offedit = False
End Sub

Private Sub remid_Change()
Dim rrr
Dim r As ADODB.Recordset, ra As ADODB.Recordset, c$
Dim csid$, i%, csvalid As Boolean, cpid$

csid$ = trm(remid.Caption)
csvalid = True
c$ = "select * from opt_checks where id='" + csid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If Not r.EOF Then
  aid.Caption = r!auftrittsid
  c$ = "select * from auftritt where id='" + aid.Caption + "'"
  Set ra = New ADODB.Recordset
  ra.CursorLocation = adUseServer
  rrr = form1.adoopen(ra, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
  If Not ra.EOF Then
    abez.Caption = trm(ra!datum) + ", " + trm(ra!bezeichnung)
    rdatum.text = r!dtg
    roffset.text = trm(CDate(trm(r!dtg)) - CDate(trm(ra!datum)))
  Else
    csvalid = False
  End If
  cpid$ = trm(r!checkid)
  If cpid$ = "" Then
    rmytext.Top = rdeftext.Top
    rdeftext.Visible = False
    rmytext.text = r!checkpoint
  Else
    c$ = "select * from opt_checklists where id='" + cpid$ + "'"
    Set ra = New ADODB.Recordset
    ra.CursorLocation = adUseServer
    rrr = form1.adoopen(ra, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    If Not ra.EOF Then
      rdeftext.Caption = trm(ra!checkpoint)
'      rmytext.Text = rdeftext.Caption
      rmytext.text = ""
    End If
  End If
End If
Command10.Enabled = False

End Sub

Private Sub rmytext_Change()
Command10.Enabled = True
End Sub

Private Sub roffset_Change()
Command10.Enabled = True
If offedit Then
  On Error Resume Next
  dtg = Val(roffset.text)
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    diffd = CDate(trm(cut_d1(abez.Caption, ","))) + dtg
    rdatum.text = trm(diffd)
  End If
End If
End Sub

Private Sub roffset_GotFocus()
dtgedit = False: offedit = True
End Sub

Private Sub savecheck()
If Command10.Enabled Then
  antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  If antw = vbYes Then
    Call Command10_Click
  End If
End If

End Sub
