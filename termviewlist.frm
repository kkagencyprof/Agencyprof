VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form termviewlist 
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton svme 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      Picture         =   "termviewlist.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Auftritt speichern"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   1
      Left            =   4440
      Picture         =   "termviewlist.frx":03A7
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "löschen"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "termviewlist.frx":167D
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "löschen"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   3135
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   4200
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   3135
      IntegralHeight  =   0   'False
      Left            =   2520
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3135
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   240
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "termviewlist.frx":2953
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   4320
      Width           =   6135
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label aid 
      Caption         =   "Label3"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4095
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "termviewlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aid_Change()
Dim id$, c$, rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "termviewlist": d2insub = "aid_Change"
id$ = aid.Caption
  c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisviz';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not rtmp.EOF
    List1(0).AddItem trm(rtmp!felddaten)
    rtmp.MoveNext
  Wend
  c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisinviz';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not rtmp.EOF
    List1(1).AddItem trm(rtmp!felddaten)
    rtmp.MoveNext
  Wend
End Sub

Private Sub Command1_Click()
'd2infile = "termviewlist": d2insub = "Command1_Click"
Unload Me
End Sub

Private Sub Command5_Click(Index As Integer)
Dim i As Integer, j As Integer, isgrp As Boolean
Dim r As ADODB.Recordset, didit As Boolean

Dim d2infile As String, d2insub As String
d2infile = "termviewlist": d2insub = "Command5_Click"
isgrp = True: didit = False
For i = 0 To List3.ListCount - 1
  If InStr(List3.List(i), "---") > 0 Then
    isgrp = False
  Else
    If List3.Selected(i) Then
'      If isgrp Then
'        Set r = New ADODB.Recordset
'        r.CursorLocation = adUseServer
'        r, "SELECT userid FROM benutzergruppen where groupid='" + List3.List(i) + "';", form1.adoc, adOpenDynamic, adLockReadOnly)
'        While Not r.EOF
'          For j = 0 To List1(Index).ListCount - 1
'            If List1(Index).List(j) = r!userid Then j = List1(Index).ListCount + 10
'          Next j
'          If j <= List1(Index).ListCount Then
'            List1(Index).AddItem r!userid
'            didit = True
'          End If
'          r.MoveNext
'        Wend
'      Else
        For j = 0 To List1(Index).ListCount - 1
          If List1(Index).List(j) = List3.List(i) Then j = List1(Index).ListCount + 10
        Next j
        If j <= List1(Index).ListCount Then
          List1(Index).AddItem List3.List(i)
          didit = True
        End If
'      End If
      List3.Selected(i) = False
      DoEvents
    End If
  End If
Next i%
If didit Then
  Me.BackColor = form1.dirtycolor()
  svme.Enabled = True
End If

End Sub

Private Sub delme_Click(Index As Integer)
Dim i, didit As Boolean

'd2infile = "termviewlist": d2insub = "delme_Click"
didit = False
For i = List1(Index).ListCount - 1 To 0 Step -1
  If List1(Index).Selected(i) Then
    List1(Index).RemoveItem i
    didit = True
  End If
Next i
If didit Then
  Me.BackColor = form1.dirtycolor()
  svme.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "termviewlist": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Label1.Caption = transe("Sichtbar für: (leer=alle)")
Label2.Caption = transe("unsichtbar für:")
svme.ToolTipText = transe("Einstellungen speichern")
delme(0).ToolTipText = transe("Markierte löschen")
delme(1).ToolTipText = delme(0).ToolTipText
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT gid FROM gruppennamen order by gid;", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
List3.Clear
While Not r.EOF
  List3.AddItem r!gid
  r.MoveNext
Wend
List3.AddItem "----------"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id FROM benutzerdaten order by id;", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  List3.AddItem r!id
  r.MoveNext
Wend
Me.BackColor = form1.cleancolor()
svme.Enabled = False
Show

End Sub

Private Sub Form_Resize()
'd2infile = "termviewlist": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

'd2infile = "termviewlist": d2insub = "Form_Unload"
Call savecheck
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub svme_Click()
Dim c$, id$, i, j, fnam$

'd2infile = "termviewlist": d2insub = "svme_Click"
id$ = aid.Caption
c$ = "delete from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisviz';"
Call form1.sqlqry(c$)
c$ = "delete from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisinviz';"
Call form1.sqlqry(c$)
fnam$ = "zzzsysisviz"
For i = 0 To 1
  For j = 0 To List1(i).ListCount - 1
    c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + _
       form1.newid("auftritthigru", "id", 18) + "','" + _
       id$ + "','" + _
       form1.auftrittstyp(id$) + "','" + _
       fnam$ + "','" + List1(i).List(j) + "')"
    Call form1.sqlqry(c$)
  Next j
  fnam$ = "zzzsysisinviz"
Next i
If form1.dayvopen Then Call dayvw.Command4_Click
Me.BackColor = form1.cleancolor()
svme.Enabled = False

End Sub

Sub savecheck()
Dim antw

'd2infile = "termviewlist": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then Call svme_Click
End If

End Sub
