VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form emailadrselect 
   Caption         =   "Emailadresse auswählen"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows-Standard
   Begin Resizer.axsResizer axsResizer1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Schliessen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   5175
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "emailadrselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim wrkJet As Workspace
'Dim sqla As Database
Dim break%
Dim callhim As String



Private Sub Command1_Click()
d2infile = "emailadrselect": d2insub = "Command1_Click"
Unload emailadrselect

End Sub

Private Sub Form_Load()
d2infile = "emailadrselect": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
emailadrselect.Caption = transe("Emailadresse auswählen")
Command1.Caption = transe("&Schliessen")
Show

End Sub

Private Sub Form_Resize()
d2infile = "emailadrselect": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "emailadrselect": d2insub = "Form_Unload"
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Private Sub List1_Click()
d2infile = "emailadrselect": d2insub = "List1_Click"
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
d2infile = "emailadrselect": d2insub = "List1_DblClick"
Call List2_DblClick
End Sub

Private Sub List2_Click()

d2infile = "emailadrselect": d2insub = "List2_Click"
If List1.ListIndex <> List2.ListIndex Then List1.ListIndex = List2.ListIndex

End Sub

Public Sub List2_DblClick()
Dim r As Recordset

d2infile = "emailadrselect": d2insub = "List2_DblClick"
t$ = List2.List(List2.ListIndex)
id$ = List1.List(List1.ListIndex)
If InStr(id$, "(KID:") > 0 Then
  kid$ = Mid$(id$, InStr(id$, "(KID:") + 5)
  id$ = form1.getadridbykontaktid(kid$)
Else
  id$ = Left$(id$, InStr(id$, "(") - 1)
  kid$ = "-1"
End If
Select Case LCase(callhim)
  Case "smtp": Call smtp.callback(id$, kid$, t$)
  Case Default
End Select

End Sub

Private Sub Text1_Change()
d2infile = "emailadrselect": d2insub = "Text1_Change"
break% = 1
Timer1.Enabled = False
Timer1.Interval = form1.getsuchvz()
Timer1.Enabled = True

End Sub

Sub rlist1(s$)
Dim rtmp As ADODB.Recordset, fcnt%
Dim d2infile As String, d2insub As String

d2infile = "emailadrselect": d2insub = "rlist1"
'On Error GoTo errhdl

List1.Clear
List2.Clear

cmd$ = "SELECT * FROM adresse where ( (" + _
       "instr(lcase(strasse),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(ort),'" + LCase(s$) + "')>0) or (" + _
       "instr(telfaxhandy,'" + s$ + "')>0) or (" + _
       "instr(lcase(url),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(hinweise),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(email),'" + LCase(s$) + "')>0) or (" + _
       "instr(lcase(name),'" + LCase(s$) + "')>0) )"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If rrr = 0 And Not rtmp.EOF Then
  fcnt% = 0
  rtmp.MoveFirst
  While Not rtmp.EOF And break% = 0 And fcnt% < 50
    em$ = form1.getemailbyid(rtmp!id)
    If em$ <> "" Then
      fcnt% = fcnt% + 1
      List1.AddItem rtmp!id & "(" + rtmp!name + ")"
      List2.AddItem em$
    End If
    rtmp.MoveNext
    DoEvents
    If break% > 0 Then
      break% = 0
      Exit Sub
    End If
  Wend
  rtmp.Close
End If

cmd$ = "SELECT ID,name FROM kontakt where (" + _
    " (instr(lcase(name),'*" + LCase(s$) + "')>0) or " + _
    " (instr(lcase(email),'" + LCase(s$) + "')>0) or " + _
    " (instr(telfaxhandy,'" + s$ + "')>0) )"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)

If rrr = 0 And Not rtmp.EOF Then
  rtmp.MoveFirst
  fcnt% = 0
  While Not rtmp.EOF And break% = 0 And fcnt% < 550
    em$ = form1.getkontaktemailbyid(rtmp!id)
    If em$ <> "" Then
      fcnt% = fcnt% + 1
      List1.AddItem rtmp!name & Space$(40) & "(KID:" & rtmp!id
      List2.AddItem em$
    End If
    rtmp.MoveNext
    DoEvents
    If break% > 0 Then
      break% = 0
      Exit Sub
    End If
  Wend
  rtmp.Close
End If
Exit Sub
errhdl:
  rrr = Err
  If rrr <> 0 Then
    If rrr <> 3420 Then MsgBox "Fehler #" & trm(str$(rrr)) & " " + Error$(rrr)
    On Error GoTo 0
    break% = 1
    Exit Sub
  End If
  Resume Next

End Sub

Public Sub Timer1_Timer()

d2infile = "emailadrselect": d2insub = "Timer1_Timer"
Timer1.Enabled = False
su$ = Text1.Text
If su$ <> "" Then
  break% = 0
  Call rlist1(su$)
End If
End Sub
Public Sub callbackto(him$)

d2infile = "emailadrselect": d2insub = "callbackto"
callhim = him$

End Sub
