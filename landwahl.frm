VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form landwahl 
   Caption         =   "Land wählen ..."
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form2"
   ScaleHeight     =   3150
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List2 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "landwahl.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Formular schiessen"
      Top             =   2760
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   600
      Top             =   2640
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "landwahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim target%, l2s As Boolean, l2sstart As Integer

Private Sub Command1_Click()
'd2infile = "landwahl": d2insub = "Command1_Click"
Hide
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
Dim mw

'd2infile = "landwahl": d2insub = "Form_Load"
l2s = False
If Not form1.geodbok Then
  List2.Visible = False
  mw = List1.Width
  Me.Width = List1.Left + List1.Width + 80
  List1.Width = mw
  Label1.Visible = False
  Check1.Visible = False
  Command1.Width = 375
End If
axsResizer1.SaveControlPositions

Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
landwahl.Caption = transe("Land wählen ...")
Command1.ToolTipText = transe("Formular schliessen")
Label1.Caption = transe("Karte zeigen")
Show

End Sub
Sub rlist1()
Dim rrr
Dim s As ADODB.Recordset, c$, lk$

Dim d2infile As String, d2insub As String
d2infile = "landwahl": d2insub = "rlist1"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
c$ = "SELECT * FROM sysvars where instr(owner,'sysvar_system_landeskennung_')=1"
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not s.EOF
  lk$ = Mid(s!Owner, 29)
  List1.AddItem lk$ & "=" & s!wert
  DoEvents
  s.MoveNext
Wend

End Sub

Private Sub Form_Resize()
'd2infile = "landwahl": d2insub = "Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "landwahl": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub

Private Sub List1_Click()
Dim c$, e$, p%, plz$, rrr, shtml$, l$, o%
Dim s As ADODB.Recordset, rdmde As Boolean, brk As Boolean
Dim sp As ADODB.Recordset, ort$

Dim d2infile As String, d2insub As String
d2infile = "landwahl": d2insub = "List1_Click"
If List1.ListIndex < 0 Then Exit Sub

MousePointer = 11: DoEvents
List2.Clear
e$ = List1.List(List1.ListIndex)
plz$ = word1(Mid$(trm(e$), 3))
p% = InStr(e$, "|")
If p% > 0 Then
  l2sstart = -1
End If
MousePointer = 0
End Sub

Public Sub List1_DblClick()
Dim e$, p%, ask%, Src$, nrad$, i%
Dim sp As ADODB.Recordset, ttest As Boolean, ftest As Boolean, tok As Boolean, fok As Boolean
Dim srct$, srcf$, c$, numlist$

Dim d2infile As String, d2insub As String
d2infile = "landwahl": d2insub = "List1_DblClick"
If List1.ListIndex < 0 Then Exit Sub

e$ = List1.List(List1.ListIndex)
p% = InStr(e$, "=")
If p% > 0 Then e$ = Mid$(e$, p% + 1)
Select Case target%
    Case 1: shwAdrDetail.datf(14).text = e$
    Case 2: shwAdrDetail.kadat(1).text = e$
    Case 3:
    Case 4:
    Case Else:
End Select
Unload Me
End Sub

Public Sub settarget(t%)
Dim mw

'd2infile = "landwahl": d2insub = "settarget"
target = t%
If target < 3 Then
  List2.Visible = False
  mw = List1.Width
  Me.Width = List1.Left + List1.Width + 80
  List1.Width = mw
  Label1.Visible = False
  Check1.Visible = False
  Command1.Width = 375
  axsResizer1.SaveControlPositions
  Call rlist1
End If
End Sub

Private Sub List2_Click()
Dim i%, s$, plst$, j%

'd2infile = "landwahl": d2insub = "List2_Click"
i% = List2.ListIndex
If i% < 0 Then Exit Sub
s$ = List2.List(i%)
i% = InStr(s$, "PLZ: ")
If i% < 2 Then Exit Sub
plst$ = LCase(trm(Left$(s$, i% - 1)))
s$ = trm(Mid$(s$, i% + 5))
For i% = 0 To List1.ListCount - 1
  If InStr(List1.List(i%), " " + s$ + " ") > 0 Then
    List1.ListIndex = i%
    DoEvents
    For j% = 0 To List2.ListCount - 1
      If InStr(LCase(List2.List(j%)), plst$) > 0 Then
        List2.ListIndex = j%
        Exit For
      End If
    Next j%
    Exit For
  End If
Next i%

End Sub

Private Sub List2_DblClick()
Dim i%

'd2infile = "landwahl": d2insub = "List2_DblClick"
i% = List2.ListIndex
If i% < 0 Then Exit Sub
If i% >= l2sstart Then l2s = True
Call List1_DblClick

End Sub
