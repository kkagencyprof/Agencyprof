VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form tpsel 
   Caption         =   "Select Project"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   ScaleHeight     =   2880
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4680
      Picture         =   "tpsel.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Liste der aufgerufenen Projekte löschen. (Löscht NICHT das Projekt)"
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   2445
      IntegralHeight  =   0   'False
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2445
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "tpsel.frx":12D6
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Formular schiessen"
      Top             =   0
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      Caption         =   "shows in:"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "tpsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rabort As Boolean, foraid As String

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub init(aid$)
foraid = aid$
Call rlists
End Sub

Private Sub delme_Click()
Dim i%, c$

i% = List2.ListIndex
If i% < 0 Then Exit Sub
c$ = "delete from opt_othertplans where aid='" + foraid + "' and tpid='" + List2.List(i%) + "'"
Call form1.sqlqry(c$)
Call rlist2

End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
rabort = False
Show
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld1
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld1:
On Error GoTo 0

End Sub

Sub rlists()
Dim rtmp As ADODB.Recordset, i%, c$, ttxt As String, rrr

rabort = False
MousePointer = 11: DoEvents
Set rtmp = New ADODB.Recordset

rtmp.CursorLocation = adUseServer
c$ = "select ID from tplan"
If Text1.text <> "" Then
  ttxt = ""
  ttxt = trm(Text1.text)
  c$ = c$ + " where ID like '%" + ttxt + "%' or ID like '%" + ttxt + "' or ID like '" + ttxt + "%'"
End If
c$ = c$ + " limit 0,99"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then
  Call form1.dbg2f("Fehler " + trm(rrr) + " " + Error$(rrr))
  MousePointer = 0
  Exit Sub
End If
List1.Clear
List2.Clear
While Not rtmp.EOF And Not rabort
  List1.AddItem trm(rtmp!id)
  rtmp.MoveNext
  DoEvents
Wend
If Not rabort Then Call rlist2
rabort = False
MousePointer = 0
End Sub

Sub rlist2()
Dim rtmp As ADODB.Recordset, i%, c$, ttxt As String, rrr

rabort = False
MousePointer = 11: DoEvents
Set rtmp = New ADODB.Recordset

rtmp.CursorLocation = adUseServer
List2.Clear
c$ = "select tpid from opt_othertplans where aid='" + foraid + "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then
  Call form1.dbg2f("Fehler " + trm(rrr) + " " + Error$(rrr))
  MousePointer = 0
  Exit Sub
End If
While Not rtmp.EOF And Not rabort
  List2.AddItem trm(rtmp!tpid)
  rtmp.MoveNext
  DoEvents
Wend
rabort = False
MousePointer = 0
End Sub

Private Sub List1_DblClick()
Dim i%, c$

i% = List1.ListIndex
If i% < 0 Then Exit Sub
c$ = "insert into opt_othertplans (aid,tpid) values('" + foraid + "','" + List1.List(i%) + "')"
Call form1.sqlqry(c$)
Call rlist2

End Sub

Private Sub List2_DblClick()
Dim i%, c$, tpid$

i% = List2.ListIndex
If i% < 0 Then Exit Sub
  
  tpid$ = List2.List(i%)
  If Len(tpid$) <> 0 Then
    Load tplan
    Call tplan.rlists
    Call tplan.nulldsp
    Call tplan.showrec(tpid$)
    On Error Resume Next
    Call tplan.SetFocus
    On Error GoTo 0
  End If

End Sub

Private Sub Text1_Change()
rabort = True
DoEvents
DoEvents
DoEvents
rabort = False
Call rlists
End Sub
