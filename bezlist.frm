VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form bezlist 
   Caption         =   "Beziehungen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox fid 
      Height          =   1875
      IntegralHeight  =   0   'False
      Left            =   8280
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ersetzen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8400
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox flist 
      Height          =   1875
      IntegralHeight  =   0   'False
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox allin1 
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Aktiviert
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   "Beziehung entfernen"
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Beziehung hinzufügen"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   2760
   End
   Begin VB.ListBox liwerte 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.ListBox bezli 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "bezlist.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Formular schiessen"
      Top             =   2760
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   4440
      Top             =   2760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ListBox allewerte 
      Height          =   2595
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   7440
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategorieansicht"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2820
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "bezlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currenta As String
Dim currentk As String
Dim allin As Boolean

Private Sub allin1_Click()

If allin1.value = 1 Then
  allin = False
  bezli.Visible = True
  liwerte.Visible = True
  allewerte.Visible = False
Else
  allin = True
  bezli.Visible = False
  liwerte.Visible = False
  allewerte.Visible = True
End If
Call rlist1

End Sub

Private Sub bezli_Click()
Dim i%, r$, c$, rtmp As ADODB.Recordset, rrr, sid$, sidp%, sida$, sidk$, fail As Boolean, j%
Dim sidka$

liwerte.Clear
flist.Clear
fid.Clear
Label1.Caption = ""
i% = bezli.ListIndex
If i% < 0 Then Exit Sub
Call getcurrent
r$ = bezli.List(i%)
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,typ,wert FROM adresstyp where typ='rel:" + r$ + "' and vid ='" + currenta + "' and kid='" + currentk + "' order by wert"
Debug.Print c$
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
While Not rtmp.EOF
  c$ = trm(rtmp!wert)
  If c$ <> "" Then
    fail = False
    If InStr(trm(rtmp!wert), "{") > 0 Then
            sid$ = trm(rtmp!wert)
            sidp% = InStr(sid$, "{")
            sida$ = sid$
            sidk$ = trm(Left(sid$, sidp% - 1))
            sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
            sidka$ = sidk$
            sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
            If sidk$ = "" Then
              fail = True
            End If
    Else
      sida$ = trm(rtmp!wert)
      If form1.getidbyid(sida$) <> sida$ Then fail = True
    End If
    
'    If fail Then
'      flist.AddItem c$
'      fid.AddItem rtmp!id
'    Else
      fail = False
      For j% = 0 To liwerte.ListCount - 1
        If liwerte.List(j%) = c$ Then
          fail = True
          Exit For
        End If
      Next j%
      If Not fail Then liwerte.AddItem c$
'    End If
  
  End If
  rtmp.MoveNext
Wend
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim i%, c$, rtmp As ADODB.Recordset, rrr
Dim sid$, cid$, vid$, p%

i% = form1.List2.ListCount
i% = i% - 1
If i% > 1 Then i% = form1.List2.ListIndex
If i% >= 0 Then
  sid$ = form1.List2.List(i%)
  p% = InStr(sid$, "ID:")
  If p% > 0 Then
    cid$ = trm(Left$(sid$, p% - 1))
    sid$ = Mid$(sid$, p% + 3)
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    c$ = "SELECT vid FROM kontakt where id='" + sid$ + "'"
    rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If rtmp.EOF Then Exit Sub
    rtmp.MoveFirst
    If Not IsNull(rtmp!vid) Then
      vid$ = rtmp!vid
    Else
      vid$ = "-1"
    End If
    c$ = "update adresstyp set wert='" + cid + " {" + vid$ + "}' where wert='" + flist.List(flist.ListIndex) + "' and instr(typ,'rel:')=1"
    Label1.Caption = c$
    Debug.Print c$
    Call form1.sqlqry(c$)
    Call bezli_Click
  End If
End If

End Sub

Private Sub Command8_Click()
Call shwAdrDetail.Command8_Click
Unload Me
End Sub

Private Sub Command9_Click()
Dim i As Integer, j As Integer

i = liwerte.ListIndex
If i < 0 Then Exit Sub
j = bezli.ListIndex
If j < 0 Then Exit Sub
Call shwAdrDetail.Command9_Click
Unload Me

End Sub

Private Sub flist_Click()
Dim i%

i% = flist.ListIndex
If i% < 0 Then Exit Sub
fid.ListIndex = i%
form1.Combo1.Text = strrepl(flist.List(i%), ",", "")

End Sub

Private Sub Form_Load()

allewerte.Top = bezli.Top
axsResizer1.SaveControlPositions
currenta = ""
currentk = "-1"
Me.BackColor = form1.cleancolor()
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
allin = True
If form1.getusersetting("beziehungeninkategorien", "ja") <> "ja" Then allin = False
If Not allin Then allin1.value = 0

Call rlist1
Show

End Sub

Private Sub Form_Resize()
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Hide

End Sub

'Public Function setcurrent(ByVal sida$, sidk$, sidkid$)
'Dim kid$

'kid$ = form1.get_kontaktid_by_name(sida$, sidk$)
'Me.Caption = "Beziehungen von: "
'If sidk$ <> "-1" Then
'  Me.Caption = Me.Caption + sidk$ + " {" + sida$ + "}"
'Else
'  Me.Caption = Me.Caption + sida$
'End If
'currenta = sida$
'currentk = sidkid$
'If currentk = "" Then currentk = "-1"

'End Function

Private Function getcurrent()
Dim kid$, sida$, sidk$, sidkid$

sida$ = shwAdrDetail.idshow.Caption
sidk$ = ""
If shwAdrDetail.klist.ListIndex >= 0 Then sidk$ = shwAdrDetail.klist.List(shwAdrDetail.klist.ListIndex)
sidkid$ = form1.get_kontaktid_by_name(sida$, sidk$)
Me.Caption = "Beziehungen von: "
If sidk$ <> "-1" Then
  Me.Caption = Me.Caption + sidk$ + " {" + sida$ + "}"
Else
  Me.Caption = Me.Caption + sida$
End If
currenta = sida$
currentk = sidkid$
If currentk = "" Then currentk = "-1"

End Function

Sub rlist1()
Dim rtmp As ADODB.Recordset, rrr, i As Integer

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresstypen where instr(id,'rel:')=1", form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then
  Unload Me
  Exit Sub
End If
While Not rtmp.EOF
  bezli.AddItem Mid(rtmp!id, 5)
  rtmp.MoveNext
Wend
allewerte.Clear
For i = 0 To shwAdrDetail.List1b.ListCount - 1
  allewerte.AddItem shwAdrDetail.List1b.List(i)
Next i
End Sub

Public Sub currentrel(rel$)
Dim p%, r$, w$, i%

Call getcurrent
r$ = rel$: w$ = ""
liwerte.Clear
flist.Clear
fid.Clear
Label1.Caption = ""
p% = InStr(rel$, ":")
If p% > 0 Then
  r$ = Left(rel$, p% - 1)
  If p% < Len(rel$) Then w$ = Mid$(rel$, p% + 1)
End If
For i% = 0 To bezli.ListCount - 1
  If bezli.List(i%) = r$ Then
    bezli.ListIndex = i%
    DoEvents
    Exit For
  End If
Next i%
If w$ <> "" Then
  For i% = 0 To liwerte.ListCount - 1
    If w$ = liwerte.List(i) Then
      liwerte.ListIndex = i%
      Exit For
    End If
  Next i%
End If
End Sub

Private Sub Label7_Click()
Call allin1_Click
End Sub

Private Sub liwerte_Click()
Dim i As Integer, j As Integer, k As Integer, c$

i = liwerte.ListIndex
If i < 0 Then Exit Sub
j = bezli.ListIndex
If j < 0 Then Exit Sub
c$ = bezli.List(j) + ":" + liwerte.List(i)
For k = 0 To shwAdrDetail.List1b.ListCount - 1
  If shwAdrDetail.List1b.List(k) = c$ Then
    shwAdrDetail.l1bdont = True
    shwAdrDetail.List1b.ListIndex = k
    DoEvents
    shwAdrDetail.l1bdont = False
    Exit For
  End If
Next k
End Sub

Private Sub liwerte_DblClick()
Dim i As Integer, sid$, sida$, sidk$, p As Integer

i = liwerte.ListIndex
If i < 0 Then Exit Sub

sid$ = liwerte.List(i)
sida$ = sid$: sidk$ = ""
p = InStr(sida$, "{")
If p > 0 Then
  sidk$ = trm(Left(sid$, p - 1))
  sida$ = trm(Mid(sid$, p + 1)): sida$ = Left(sida$, Len(sida$) - 1)
End If
If Len(sida$) > 0 Then Call shwAdrDetail.refreshadrdetail(sida$, sidk$)

End Sub

Private Sub Timer1_Timer()

If Not shwAdrDetail.List1b.Visible Then Unload Me

End Sub
