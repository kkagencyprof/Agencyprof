VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Maileingang"
   ClientHeight    =   5985
   ClientLeft      =   4140
   ClientTop       =   3105
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11010
   Begin VB.CommandButton Command9 
      Caption         =   "report spam"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   38
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "quick address"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   37
      ToolTipText     =   "create an address for the selected mail"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton spma 
      Caption         =   "Spam markieren"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pgb2 
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Spamwörer bearbeiten"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "Liste aller Spammerworte bearbeiten. Nach Änderung bitte neu starten oder erneut hier klicken."
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Domain z. schwz. Liste"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox autoblack 
      Height          =   255
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "Markiert beim Öffnen alle Einträge der Schwarzen Liste"
      Top             =   4320
      Width           =   255
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
      Height          =   375
      Left            =   480
      TabIndex        =   30
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox tout 
      Height          =   285
      Left            =   10320
      TabIndex        =   27
      Text            =   "30"
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox brlle 
      Height          =   615
      Index           =   1
      Left            =   4680
      Picture         =   "Frmmain.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox brlle 
      Height          =   615
      Index           =   0
      Left            =   4080
      Picture         =   "Frmmain.frx":117A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      Picture         =   "Frmmain.frx":22F4
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "zeigen"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      Picture         =   "Frmmain.frx":346E
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Markierte Nachrichten löschen"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Frmmain.frx":4744
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "vom Server laden"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "zur schwarzen Liste"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Schwarze Liste"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "alle von 'dort'"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "invert"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   6480
   End
   Begin VB.TextBox pin 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   16
      Top             =   6360
      Width           =   735
   End
   Begin VB.ComboBox txtServer 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   4440
      TabIndex        =   14
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Bekannte"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   2640
      Top             =   6480
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Mailstatus auf allen Servern, Doppelklick zum Aktualisieren"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "Frmmain.frx":510A
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtMessage 
      Height          =   5775
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   8655
   End
   Begin MSComctlLib.ListView listMessages 
      Height          =   5775
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10186
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   8160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6360
      Width           =   795
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "&Mail testen"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   9360
      TabIndex        =   1
      Text            =   "110"
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label bbox 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOut:"
      Height          =   255
      Left            =   9840
      TabIndex        =   28
      Top             =   6360
      Width           =   435
   End
   Begin VB.Label cmdViewl 
      Caption         =   "Zeigen"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label lblPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   9000
      TabIndex        =   6
      Top             =   6360
      Width           =   435
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "Pop Server:"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   6360
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nomess%, wvlt%, pcode
Dim poplistok%
Dim prevfire%, currentmozfile As String
Dim tot As Long
Dim cmdDelete_ask%
Public mozopen As Boolean

Const LVCFROM = 0
Const LVCSUBJECT = 1
Const LVCSIZE = 2
Const LVCDATE = 3
Const LVCTO = 4
Const LVCID = 5
Const LVCFILE = 6
Const LVCCC = 7

Private Sub autoblack_Click()
'd2infile = "Frmmain": d2insub = "autoblack_Click"
Call form1.setmylastFormVar(Me.name, "autoblack", trm(autoblack.value))

End Sub

Public Sub cmdCheckMail_Click()
Dim intNum As Integer
Dim msgn%, intMessageNum As Integer 'the number of messages
Dim strMessageHeader As String, sbj$, sbf$
Dim i As Integer
Dim lvitem, msgid$, trg$
Dim gl As ListItem, anyconn As Boolean, pos%
Dim fn$, o%, dellst(99) As String, delptr As Integer

fn$ = form1.s0dir() + "\" + form1.docs() + "\" + form1.getuserid() + "\autospam.txt"
If Not nexist(fn$) Then
  o% = FreeFile: delptr = 0
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, fn$: fn$ = trm(fn$)
    If Len(fn$) > 1 Then
      If Left(fn$, 1) <> ";" Then
        dellst(delptr) = fn$
        delptr = delptr + 1: If delptr > 99 Then delptr = 99
      End If
    End If
  Wend
  Close #o%
End If
anyconn = False
Timer1.Enabled = True
Timer1.Interval = 100

End Sub

Private Sub cmdDelete_Click()
'd2infile = "Frmmain": d2insub = "cmdDelete_Click"
    'Connect to the server and delete any selected messages
    Dim lvitem
    Dim i As Integer
    Dim intNum As Integer
    Dim bSelected As Boolean

    form1.ihavemail = False
    bSelected = False

    If cmdDelete_ask% = 1 Then
      If MsgBox(transe("Wollen Sie diese Nachricht(en) wirklich löschen?"), vbQuestion + vbYesNo, App.Title) = vbNo Then
        On Error Resume Next
        listMessages.SetFocus
        On Error GoTo 0
        Exit Sub
      End If
    End If
    For i = listMessages.ListItems.Count To 1 Step -1
        If (listMessages.ListItems(i).Selected = True) Then
          On Error Resume Next
          Kill listMessages.ListItems(i).SubItems(LVCFILE)
          On Error GoTo 0
          bSelected = True
        End If
    Next
    If cmdDelete_ask% = 1 Then listMessages.SetFocus

    If (bSelected) Then
        'If messages were deleted, then use cmdCheckMail to redisplay messages
        cmdCheckMail_Click
    Else
        MsgBox "Es sind keine Nachrichten ausgewählt."
    End If

    If cmdDelete_ask% = 1 Then listMessages.SetFocus

End Sub

Private Sub cmdView_Click()
'd2infile = "Frmmain": d2insub = "cmdView_Click"
    'Views the first selected message
    Dim i As Integer, rrr
    Dim intResultCode As Integer
    Dim intNum As Integer
    Dim strMessage As String


    If (cmdViewl.Caption = "Schliessen") Then
        cmdView.Picture = brlle(0).Picture
        listMessages.Visible = True
        txtMessage.Visible = False
        cmdDelete.Enabled = True
        cmdCheckMail.Enabled = True
        cmdViewl.Caption = "Zeigen"
        txtMessage.text = ""
        listMessages.SetFocus
    Else
        Call cmdViewfromFile
    End If
End Sub

Private Sub Command1_Click()

'd2infile = "Frmmain": d2insub = "Command1_Click"
Call form1.dbg2f("closef rmMain (Command1_Click)")
form1.ihavemail = False
If (cmdViewl.Caption = "Schliessen") Then Call cmdView_Click

Unload Me

End Sub

Private Sub Command11_Click()
Dim frm$, c$, r As ADODB.Recordset, u$, o%, i As Integer, lvitem, ad1 As Boolean, X
Dim rrr

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Command11_Click"
u$ = form1.getuserid()
On Error GoTo exsux
ad1 = False
For i = 1 To listMessages.ListItems.Count
  Set lvitem = listMessages.ListItems(i)
  If (listMessages.ListItems(i).Selected = True) Then
'frm$ = domainofemail(emailonly(strrepl(listMessages.SelectedItem, """", "")))
    frm$ = domainofemail(emailonly(strrepl(listMessages.ListItems(i), """", "")))
    On Error GoTo 0
    If frm$ <> "" Then
      c$ = "select * from sysvars where owner='blacklistdom:" + u$ + "' and wert='" + frm$ + "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If r.EOF Then
        c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 36) + _
           "','blacklistdom:" + u$ + _
           "','" + frm$ + "')"
        Call form1.sqlqry(c$)
        o% = FreeFile
        Open form1.mydir() + "\spamdoms.txt" For Append As #o%
        c$ = "insert into spammer (id) values('" + frm$ + "@web1p1');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web10p1');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web2p2');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web7p1');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web7p2');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web7p3');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web7p4');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@web7p5');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@juliaalbrecht');": Print #o%, c$
        c$ = "insert into spammer (id) values('" + frm$ + "@mch');": Print #o%, c$
        ad1 = True
        Close #o%
      End If
    End If
  End If
Next i
exsux:
'If ad1 Then x = Shell("notepad.exe " + form1.mydir() + "\spamdoms.txt", vbNormalFocus)
Call Command6_Click


End Sub

Private Sub Command12_Click()
Dim fn$, X

'd2infile = "Frmmain": d2insub = "Command12_Click"
Call form1.read_spmlst
fn$ = form1.s0dir() + "\" + "spamwrds.txt"
X = Shell("notepad.exe " + fn$, 1)

End Sub

Private Sub Command18_Click()

'd2infile = "Frmmain": d2insub = "Command18_Click"
Call form1.handbuchcall("13-Email.htm")

End Sub

Private Sub Command2_Click()
Dim i As Integer

Dim from$, frome$, c$, lvitem
'd2infile = "Frmmain": d2insub = "Command2_Click"
Call listMessages.SetFocus
On Error GoTo erroutc2
For i = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i).Selected = False Then
    listMessages.ListItems(i).Selected = True
  Else
    listMessages.ListItems(i).Selected = False
  End If
  DoEvents
Next i
erroutc2:
On Error Resume Next
End Sub

Private Sub Command3_Click()

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Command3_Click"
Call getMsgfromFile

End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim dh As ADODB.Recordset, rrr
Dim from$, frome$, c$, lvitem

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Command4_Click"
MousePointer = 11
Call listMessages.SetFocus
For i = 1 To listMessages.ListItems.Count
  listMessages.ListItems(i).Selected = False
Next i
DoEvents
For i = 1 To listMessages.ListItems.Count
  Set lvitem = listMessages.ListItems(i)
  from$ = listMessages.ListItems(i)
  lvitem.Selected = form1.knownaddress(from$)
Next i
fastx:
MousePointer = 0
End Sub

Private Sub Command5_Click()
Dim i As Integer, rrr
Dim frm$, lvitem


'd2infile = "Frmmain": d2insub = "Command5_Click"
On Error Resume Next
frm$ = listMessages.SelectedItem
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

Call listMessages.SetFocus
For i = 1 To listMessages.ListItems.Count
  listMessages.ListItems(i).Selected = False
Next i
DoEvents
For i = 1 To listMessages.ListItems.Count
  Set lvitem = listMessages.ListItems(i)
  If frm$ = listMessages.ListItems(i) Then listMessages.ListItems(i).Selected = True
  DoEvents
Next i


End Sub

Private Sub Command6_Click()
Dim i As Integer, j%
Dim dh As ADODB.Recordset, rrr
Dim from$, frome$, c$, lvitem, fromdom$, sbj$

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Command6_Click"
MousePointer = 11
Call listMessages.SetFocus
For i = 1 To listMessages.ListItems.Count
  listMessages.ListItems(i).Selected = False
Next i
DoEvents
For i = 1 To listMessages.ListItems.Count
If Not listMessages.ListItems(i).Selected Then
  Set lvitem = listMessages.ListItems(i)
  from$ = emailonly(strrepl(listMessages.ListItems(i), """", ""))
  fromdom$ = domainofemail(from$)
  sbj$ = lvitem.SubItems(LVCSUBJECT)
  If form1.spambetreff(sbj$) Then
    listMessages.ListItems(i).Selected = True
  Else
    c$ = "SELECT * FROM sysvars where (((owner='blacklist:" & form1.getuserid() + "') and (wert='" & from$ & "')) "
    If fromdom <> "" Then
      c$ = c$ & "or ((owner='blacklistdom:" & form1.getuserid() + "') and (wert='" & fromdom$ & "'))"
    End If
    c$ = c$ & ")"
    Set dh = New ADODB.Recordset
    dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr = 0 Then
      If Not dh.EOF Then
        listMessages.ListItems(i).Selected = True
        DoEvents
      End If
    End If
  End If
  If listMessages.ListItems(i).Selected = True Then
    c$ = "SELECT * FROM adresse where ( (instr(lcase(email),'" + LCase(from$) + "')>0) )"
    Set dh = New ADODB.Recordset
    dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr <> 0 Then GoTo fa1stx
    If Not dh.EOF Then
      listMessages.ListItems(i).Selected = False
      DoEvents
    End If
    If listMessages.ListItems(i).Selected = True Then
      c$ = "SELECT ID FROM kontakt where ( (instr(lcase(email),'" + LCase(from$) + "')>0)  )"
      Set dh = New ADODB.Recordset
      dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not dh.EOF Then
        listMessages.ListItems(i).Selected = False
        DoEvents
      End If
    End If
  End If
  If listMessages.ListItems(i).Selected = True Then
    For j% = i + 1 To listMessages.ListItems.Count
      If InStr(listMessages.ListItems(j%), from$) > 0 And listMessages.ListItems(j%).Selected = False Then
        listMessages.ListItems(j%).Selected = True
      End If
    Next j%
  End If
End If
fa1stx:
Next i
MousePointer = 0
End Sub

Private Sub Command7_Click()
Dim frm$, c$, r As ADODB.Recordset, u$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Command7_Click"
u$ = form1.getuserid()
On Error GoTo exsu
frm$ = emailonly(strrepl(listMessages.SelectedItem, """", ""))
On Error GoTo 0
If frm$ <> "" Then
  c$ = "select * from sysvars where owner='" & u$ & "' and wert='" & frm$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then Exit Sub
  c$ = "insert into sysvars (id,owner,wert) values('" & form1.newid("sysvars", "id", 36) + _
         "','blacklist:" & u$ + _
         "','" & frm$ & "')"
  Call form1.sqlqry(c$)
End If
exsu:
Call Command6_Click

End Sub


Private Sub Command8_Click()
Dim frm$, cmd$

frm$ = trm(form1.Combo1.text)
Call form1.sqlqry("insert into adresse (id,name,email) values('" + frm$ + "','" + frm$ + "','" + frm$ + "')")
Call form1.sqlqry("insert into adresstyp (id,vid,typ,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + frm$ + "','Person','-1')")
Call Command4_Click

End Sub

Private Sub Command9_Click()
Dim i As Integer, rd As Integer, l$, strMessage As String
Dim b$, o%, eml$, knt$, rrr, sbj$

For i = 1 To listMessages.ListItems.Count
  If (listMessages.ListItems(i).Selected = True) Then
    Call form1.dbg2f("report as spam: " + listMessages.ListItems(i).SubItems(LVCFILE))
    
    Load smtp
    smtp.Visible = True
    smtp.txtSendTo = "blockspam@spamlab.biz"
    smtp.adrid = ""
    smtp.kid = ""
    Call smtp.txtMessageSubject.SetFocus
    smtp.txtServer.Enabled = False
    smtp.txtMailFrom.Enabled = False
    smtp.txtMessageText.text = smtp.txtMessageText.text & "Dear Sir or Madam," & vbCrLf & vbCrLf & "please find attached an undetected spam e-mail." & vbCrLf & vbCrLf
    smtp.txtMessageText.text = smtp.txtMessageText.text & "Regards" & vbCrLf
    smtp.txtMessageText.text = smtp.txtMessageText.text & form1.uname$ & vbCrLf
    Call form1.signaturinclude
    sbj$ = "Reporting a spam e-mail"
    smtp.txtMessageSubject = sbj$
    Name listMessages.ListItems(i).SubItems(LVCFILE) As listMessages.ListItems(i).SubItems(LVCFILE) + ".eml"
    Call smtp.attachfile(listMessages.ListItems(i).SubItems(LVCFILE) + ".eml")
    On Error Resume Next
    Kill listMessages.ListItems(i).SubItems(LVCFILE) + ".eml"
    On Error GoTo 0
    DoEvents
    'Call smtp.cmdSend_Click
    Call listMessages.ListItems.Remove(i)
    Exit Sub
  End If
Next

End Sub

Private Sub Form_Load()
Dim colHeader
Dim r As ADODB.Recordset, rrr, stst$, i%, s%, klrv%, c$

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Form_Load"

spma.Visible = True
List1.Height = cmdCheckMail.Top - List1.Top
DoEvents

axsResizer1.SaveControlPositions

i% = FreeFile
c$ = Environ$("HOMEDRIVE") + Environ$("HOMEPATH") + "\appop.dat"
Open c$ For Output As #i%
Print #i%, form1.getuserid()
Print #i%, form1.mylastFormVar(Me.name, "delrecvd", "0")
Close #i%
wvlt% = 0
prevfire% = 0
cmdView.Picture = brlle(0).Picture
poplistok% = 0
pin.Visible = False
Label2.Visible = False
List1.Visible = False
s% = form1.myfontsize()
List1.Font.Size = s%
txtMessage.Font.Size = s%
listMessages.Font.Size = s%
mozopen = False
klrv% = Val(form1.mylastFormVar(Me.name, "autoblack", "0"))
If klrv% <> 1 Then klrv% = 0
autoblack.value = klrv%

cmdDelete_ask% = 1
cmdViewl.Caption = "Zeigen"
cmdView.Enabled = True

c$ = "SELECT id FROM poplist where instr(id,'" & form1.getuserid() & "_')>0 and id<>'" & form1.getuserid() & "_PDFServer'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  r.Close
  poplistok% = 1
  pin.Visible = True
  Label2.Visible = True
  List1.Visible = True
End If

nomess% = 0
listMessages.View = lvwReport
Set colHeader = listMessages.ColumnHeaders.add(, , "Von", 1600)
Set colHeader = listMessages.ColumnHeaders.add(, , "Betreff", 4350)
Set colHeader = listMessages.ColumnHeaders.add(, , "kB", 800)
Set colHeader = listMessages.ColumnHeaders.add(, , "Datum", 2300)
Set colHeader = listMessages.ColumnHeaders.add(, , "To", 1200)
Set colHeader = listMessages.ColumnHeaders.add(, , "Message-ID", 2)
Set colHeader = listMessages.ColumnHeaders.add(, , "MessageFile", 2)
Set colHeader = listMessages.ColumnHeaders.add(, , "CC", 2)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

tout.text = form1.getusersetting("popmailtimeout", "30")
tot = Val(tout.text)
lblStatus.Caption = "ok"
frmMain.Caption = transe("Maileingang")
Command12.Caption = transe("Spamwörer bearbeiten")
Command12.ToolTipText = transe("Liste aller Spammerworte bearbeiten. Nach Änderung bitte neu starten oder erneut hier klicken.")
Command11.Caption = transe("Domain z. schwz. Liste")
autoblack.ToolTipText = transe("Markiert beim Öffnen alle Einträge der Schwarzen Liste")
Command18.Caption = transe("?")
Command18.ToolTipText = transe("Hilfeseite öffnen")
cmdView.ToolTipText = transe("zeigen")
cmdDelete.ToolTipText = transe("Markierte Nachrichten löschen")
Command3.ToolTipText = transe("vom Server laden")
Command7.Caption = transe("zur schwarzen Liste")
Command6.Caption = transe("Schwarze Liste")
Command5.Caption = transe("alle von dort")
Command2.Caption = transe("invert")
Command4.Caption = transe("&Bekannte")
List1.ToolTipText = transe("Mailstatus auf allen Servern, Doppelklick zum Aktualisieren")
Command1.ToolTipText = transe("Dieses Formular schliessen")
cmdCheckMail.Caption = transe("&Mail testen")
Label3.Caption = transe("TOut:")
cmdViewl.Caption = transe("Zeigen")
Label2.Caption = transe("PIN:")
Label1.Caption = transe("Password:")
lblUserName.Caption = transe("User")
lblPort.Caption = transe("Port:")
lblServer.Caption = transe("Pop Server:")

Show
End Sub


Public Sub chkm()
'd2infile = "Frmmain": d2insub = "chkm"
DoEvents
Call cmdCheckMail_Click
End Sub

Private Sub Form_Resize()
'd2infile = "Frmmain": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub


Private Sub Form_Unload(Cancel As Integer)

'd2infile = "Frmmain": d2insub = "Form_Unload"
If mozopen Then Unload trvw
Hide
Call form1.dbg2f("unloading frmMain")
On Error Resume Next
Kill form1.s0dir() & "\debug2file_" & form1.getuserid() & "_frmMain_Socket1.txt"
On Error GoTo 0
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub


Private Sub List1_Click()
Dim id$, i%

'd2infile = "Frmmain": d2insub = "List1_Click"
If List1.ListIndex < 0 Then Exit Sub
id$ = List1.List(List1.ListIndex)
id$ = Mid$(id$, InStr(id$, "auf ") + 4)
Call txtServer_DropDown
For i% = 0 To txtServer.ListCount - 1
  If txtServer.List(i%) = id$ Then
    txtServer.ListIndex = i%
    Exit For
  End If
Next i%

DoEvents
End Sub

Private Sub List1_DblClick()
'd2infile = "Frmmain": d2insub = "List1_DblClick"
List1.Clear
End Sub

Private Sub listMessages_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub listMessages_Click()
Dim frm$, p%, rrr, i%, o%, l$, sbf$, sbj$, trg$, msgid$, pos%
Dim lvitem, hdm As Boolean, strMessageHeader As String
Dim r As ADODB.Recordset, c$, n%

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "listMessages_Click"
On Error Resume Next
frm$ = listMessages.SelectedItem
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

p% = 0: n% = 0
Command5.Enabled = False
For i% = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i%).Selected = True Then
    p% = p% + 1
    If p% > 1 Then Exit For
  End If
Next i%
If p% = 1 Then
  Command5.Enabled = True
End If
Command8.Enabled = True
Command9.Enabled = False
If form1.getusersetting("expurgateprotected", "nein") = "ja" Then Command9.Enabled = True
If form1.knownaddress(frm$) Then
  Command8.Enabled = False
End If
If Not mozopen Then
  currentmozfile = ""
  p% = InStr(frm$, "<")
  If p% > 0 Then
    frm$ = Mid$(frm$, p% + 1)
    frm$ = Left$(frm$, InStr(frm$, ">") - 1)
  Else
    p% = InStr(frm$, "(")
    If p% > 0 Then
      frm$ = trm(Left$(frm$, p% - 1))
    End If
  End If
  form1.Combo1.text = frm$
Else
  hdm = False
  frm$ = trvw.currnod.Caption + "\" + frm$
  If Not nexist(frm$) Then
    currentmozfile = frm$
    frmMain.listMessages.ListItems.Clear
    o% = FreeFile
    Open frm$ For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      If Left$(l$, 5) = "From " Then
        sbf$ = word2bis(l$)
        strMessageHeader = ""
        hdm = True
      End If
      If l$ = "" And hdm Then hdm = False
      If hdm Then strMessageHeader = strMessageHeader + vbCrLf + l$
      If (Not hdm) And strMessageHeader <> "" Then
        sbj$ = GetHeaderValue(strMessageHeader, "Subject")
        sbf$ = GetHeaderValue(strMessageHeader, "From")
        pos% = InStr(sbf$, "@")
        If pos% = 0 Then
          sbf$ = emailonly(GetHeaderValue(strMessageHeader, "Return-Path"))
        End If
        trg$ = GetHeaderValue(strMessageHeader, "To")
        n% = n% + 1
        Set lvitem = listMessages.ListItems.add(, , sbf$)
        lvitem.SubItems(LVCSUBJECT) = sbj$
        lvitem.SubItems(LVCTO) = trg$
        lvitem.SubItems(LVCSIZE) = trm(Int(FileLen(frm$) / 1000))
        lvitem.SubItems(LVCDATE) = GetHeaderValue(strMessageHeader, "Date")
        msgid$ = Left(GetHeaderValue(strMessageHeader, "Message-ID"), 128)
        lvitem.SubItems(LVCID) = msgid$
        strMessageHeader = ""
        c$ = "SELECT * FROM mailsafe where instr(id,'-" + msgid$ + "')>0;"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        lvitem.Selected = False
        If r.EOF Then
          lvitem.Selected = True
        End If
        DoEvents
      End If
    Wend
    Close #o%
    Call listMessages.SetFocus
    lblStatus.Caption = n% & " " + transe("Nachrichten")
  Else
    p% = InStr(frm$, "<")
    If p% > 0 Then
      frm$ = Mid$(frm$, p% + 1)
      frm$ = Left$(frm$, InStr(frm$, ">") - 1)
    Else
      p% = InStr(frm$, "(")
      If p% > 0 Then
        frm$ = trm(Left$(frm$, p% - 1))
      End If
    End If
    form1.Combo1.text = frm$
  End If
End If
End Sub
Private Sub listMessages_DblClick()

'd2infile = "Frmmain": d2insub = "listMessages_dblClick"
If Not mozopen Then
  Call cmdView_Click
End If

End Sub

Private Sub listMessages_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, j%
Dim lvitem

'd2infile = "Frmmain": d2insub = "listMessages_KeyDown"
'<strg>a
If KeyCode = 65 And pcode = 17 Then
  For i = listMessages.ListItems.Count To 1 Step -1
    listMessages.ListItems(i).Selected = True
  Next i
End If
If KeyCode = 8 Or KeyCode = 46 Then Call cmdDelete_Click

pcode = KeyCode

End Sub

Private Sub spma_Click()
Dim fn$, o%, X

fn$ = form1.s0dir() + "\" + form1.docs() + "\" + form1.getuserid() + "\autospam.txt"
If nexist(fn$) Then
  o% = FreeFile
  Open fn$ For Output As #o%
  Print #o%, "; Sie können Mails automatisch markieren lassen."
  Print #o%, "; Tragen Sie dazu hier die Worte ein, mit denen der Betreff beginnen muss."
  Print #o%, "; Die Eingabe mehrerer Begriffe ist erlaubt, ein Begriff pro Zeile, max. 100 Begriffe."
  Print #o%,
  Print #o%, "; Leerzeilen und Zeilen, die mit ; beginnen werden ignoriert."
  Print #o%,
  Print #o%, "; In Verbindung mit einem externen Spamfilter lassen sich Spams so leicht löschen."
  Print #o%, "; Wenn Sie z.B. SpamAssain verwenden, so tragen Sie hier ***SPAM*** ein."
  Print #o%, "; Bei Verwendung von spamfence (empfohlen, siehe spamfence.net) tragen Sie hier [eX-Spam] ein."
  Print #o%,
  Print #o%, ";***SPAM***"
  Print #o%, ";[eX-Spam]"
  Close #o%
End If
X = Shell("notepad.exe " + fn$, 1)

End Sub

Private Sub Timer1_Timer()
Dim r As ADODB.Recordset, u$, ap$, c$
Dim aKey() As Byte, rc$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "Timer1_Timer"
Timer1.Enabled = False
Timer1.Interval = 0
u$ = form1.getuserid()
Call cmdCheckInbox

Call mrkbysize
End Sub


Private Sub tout_Change()
'd2infile = "Frmmain": d2insub = "tout_Change"
tot = Val(tout.text)
End Sub


Private Sub txtServer_DropDown()
Dim r As ADODB.Recordset, u$, ap$, c$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "txtServer_DropDown"
If poplistok% = 0 Then Exit Sub

ap$ = trm(pin.text)
If trm(ap$) = "" Then
  MsgBox "Bitte geben Sie das Passwort ein mit dem die Passwörter in der Datenbank entschüsselt werden."
  Call pin.SetFocus
  Exit Sub
End If
txtServer.Clear
txtServer.AddItem "dir:Inbox"
u$ = form1.getuserid()
c$ = "SELECT id FROM poplist where instr(id,'" + u$ + "_')>0 and id<>'" & u$ & "_PDFServer'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  While Not r.EOF
    If r!id <> "PDFServer" Then
      txtServer.AddItem Mid$(r!id, Len(u$) + 2)
    End If
    r.MoveNext
  Wend
End If
End Sub
Sub rlist1()
Dim intNum As Integer, rrr
Dim intMessageNum As Integer 'the number of messages
Dim ap$, r As ADODB.Recordset, u$, pw$, aKey() As Byte, c$

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "rlist1"
If pin.Visible = True And trm(pin.text) = "" Then
  If form1.pin.Visible = True And trm(form1.pin.text) <> "" Then
    pin.text = form1.pin.text
  End If
End If
Call form1.noerrshow
MousePointer = 11
DoEvents
List1.Clear
Call form1.errshow
MousePointer = 0
DoEvents

End Sub

Public Sub popfire(r As Long, t As Long)
Dim l$, p%, dr As Double

'd2infile = "Frmmain": d2insub = "popfire"
dr = r
l$ = lblStatus.Caption
p% = InStr(l$, vbCrLf)
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
If t = 0 Then Exit Sub
If Int(dr * 20 / t) <> prevfire% Then
  prevfire% = Int(dr * 20 / t)
  lblStatus.Caption = l$ & vbCrLf & trm(dr) & "/" & trm(t) & " (" & trm(Int(dr * 100 / t)) & "%)"
  pgb2.value = Int(dr * 100 / t)
  DoEvents
End If
End Sub

Sub hostnameupdate()
Dim dbn$, dbp$, cmd$, webhostupd As Boolean, rtmp As QueryDef, rrr
Dim rloc As ADODB.Recordset, rrem As ADODB.Recordset, ipn$

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "hostnameupdate"
cmd$ = "select id,info from webhosts where dnsname=id"
Set rloc = New ADODB.Recordset
rloc.CursorLocation = adUseServer
rrr = form1.adoopen(rloc, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While Not rloc.EOF
 ipn$ = GetHostFromIP(rloc!id)
 cmd$ = "update webhosts set dnsname='" & ipn$ & "',info='' where id='" & rloc!id & "'"
 Call form1.sqlqry(cmd$)
 rloc.MoveNext
Wend
rloc.Close

End Sub

Public Sub cmdCheckInbox()
Dim intNum As Integer, ut1$, an$, anl$, c$, cc$, ccl$
Dim msgn%, intMessageNum As Integer 'the number of messages
Dim strMessageHeader As String, sbj$, sbf$, pos%
Dim i As Integer, dn As String, tr, lckd As Integer
Dim lvitem, msgid$, trg$, u$, dc As String
Dim gl As ListItem, anyconn As Boolean
Dim fn$, o%, dellst(99) As String, delptr As Integer

fn$ = form1.s0dir() + "\" + form1.docs() + "\" + form1.getuserid() + "\autospam.txt"
Call form1.dbg2f("spamdefs: " + fn$)
If Not nexist(fn$) Then
  o% = FreeFile: delptr = 0
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, fn$: fn$ = trm(fn$)
    If Len(fn$) > 1 Then
      If Left(fn$, 1) <> ";" Then
        dellst(delptr) = fn$
        delptr = delptr + 1: If delptr > 99 Then delptr = 99
      End If
    End If
  Wend
  Close #o%
End If

anyconn = False

Call form1.noerrshow
MousePointer = 11
DoEvents
listMessages.ListItems.Clear
u$ = form1.getuserid()
dn = form1.mylocaldatadir() + "\mail\inbox"
dn = form1.getusersetting("mailinboxdir", dn)
Call form1.dbg2f("datadir: " + dn)
    intMessageNum = form1.InboxMessageCount(dn)
Call form1.dbg2f("msgcount: " + trm(intMessageNum))
    If intMessageNum = POP_SOCKET_ERROR Then
        MsgBox ("Fehler beim Mailstatus: Error " + trm(intMessageNum))
        MousePointer = 0
        Exit Sub
    End If
    lblStatus.Caption = transe("Empfange Mailheader ...")
    If (intMessageNum <> 0) Then
        cmdDelete.Enabled = True
        cmdViewl.Enabled = True
    Else
        cmdDelete.Enabled = False
        cmdViewl.Enabled = False
        MousePointer = 0
        lblStatus.Caption = ""
        Exit Sub
    End If

    anyconn = True
    msgn% = Val(form1.getusersetting("maxpopindex", "1000"))
    If msgn% > intMessageNum Then msgn% = intMessageNum
    On Error Resume Next
    pgb1.Max = msgn%
    pgb1.value = 0
    pgb1.Visible = True
    On Error GoTo 0
    lckd = 0
    For i = 1 To msgn%
      pgb1.value = i
      DoEvents
      tr = InboxGetMessageHeader(dn, i, strMessageHeader)
      If tr <> "" Then
        sbj$ = GetHeaderValue(strMessageHeader, "Subject")
        If InStr(LCase(sbj$), "utf") > 0 Then
          ut1$ = utf8sbjdecode(sbj$): If ut1$ <> "" Then sbj$ = ut1$
        End If
        If InStr(LCase(sbj$), "iso-8859") > 0 Then
          sbj$ = QuotedPrintableDecode(sbj$)
        End If
        sbf$ = GetHeaderValue(strMessageHeader, "From")
        pos% = InStr(sbf$, "@")
        If pos% = 0 Then
          sbf$ = GetHeaderValue(strMessageHeader, "Return-Path")
          If Left$(sbf$, 1) = "<" Then sbf$ = Mid$(sbf$, 2)
          If Right$(sbf$, 1) = ">" Then sbf$ = Left$(sbf$, Len(sbf$) - 1)
        End If
        an$ = GetHeaderValue(strMessageHeader, "To"): anl$ = ""
        While an$ <> ""
          c$ = cut_d1(an$, ","): an$ = cut_d2bis(an$, ",")
          If anl$ <> "" And Right(trm$(anl$), 1) <> "," Then
            anl$ = anl$ + ","
          End If
          anl$ = anl$ + emailonly(c$)
        Wend
        trg$ = anl$
        cc$ = GetHeaderValue(strMessageHeader, "CC"): ccl$ = ""
        While cc$ <> ""
          c$ = cut_d1(cc$, ","): cc$ = cut_d2bis(cc$, ",")
          If ccl$ <> "" And Right(trm$(ccl$), 1) <> "," Then
            ccl$ = ccl$ + ","
          End If
          ccl$ = ccl$ + emailonly(c$)
        Wend

        Set lvitem = listMessages.ListItems.add(, , sbf$)
        lvitem.Selected = False
        lvitem.SubItems(LVCSUBJECT) = sbj$
        lvitem.SubItems(LVCTO) = trg$
        lvitem.SubItems(LVCSIZE) = trm(Int(FileLen(tr) / 1000))
        lvitem.SubItems(LVCDATE) = GetHeaderValue(strMessageHeader, "Date")
        msgid$ = Left(GetHeaderValue(strMessageHeader, "Message-ID"), 128)
        lvitem.SubItems(LVCID) = msgid$
        lvitem.SubItems(LVCCC) = ccl$
        lvitem.SubItems(LVCFILE) = tr
        lblStatus.Caption = trm(i) + "/" + trm(intMessageNum)
      Else
        lckd = lckd + 1
      End If
    Next i
    pgb1.Visible = False
    intMessageNum = intMessageNum - lckd
    lblStatus.Caption = trm(intMessageNum) + " " + transe("Nachrichten")

schuessdann:
lblStatus.Caption = ""
listMessages.Visible = True
On Error Resume Next
listMessages.SetFocus
On Error GoTo 0
Call form1.errshow
DoEvents
If List1.ListCount = 0 Or List1.List(0) = transe("PIN fehlt") Then Call rlist1
If autoblack.value = 1 Then Call Command6_Click
If delptr > 0 Then
  For i = 1 To listMessages.ListItems.Count
    For o% = 0 To delptr - 1
      If Left(listMessages.ListItems(i).SubItems(LVCSUBJECT), Len(dellst(o%))) = dellst(o%) Then
        listMessages.ListItems(i).Selected = True
        DoEvents
      End If
    Next o%
  Next i
End If
'If anyconn = True Then Call hostnameupdate
MousePointer = 0

End Sub


Public Function InboxGetMessageHeader(inbx As String, intMessage As Integer, strHeader As String) As String
    'Stores the message header into strHeader using the TOP command
    Dim l$, tr, n As Integer, rrr

    strHeader = ""
    InboxGetMessageHeader = ""
    n = intMessage
    On Error GoTo hdrout6767
    tr = Dir(inbx + "\*.amf")
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then Exit Function
    While tr <> "" And n > 0
'      Debug.Print tr
      n = n - 1
      If n > 0 Then tr = Dir
    Wend
    If tr <> "" And nexist(inbx + "\" + tr + ".lck") Then
      n = FreeFile
      Open inbx + "\" + tr For Input As #n
      InboxGetMessageHeader = inbx + "\" + tr
      Do
        On Error Resume Next
        Line Input #n, l$
        rrr = Err
        On Error GoTo 0
        If rrr = 0 Then
          If strHeader <> "" Then strHeader = strHeader + Chr$(13)
          strHeader = strHeader + l$
        Else
          InboxGetMessageHeader = ""
          MsgBox ("Fehler #" + trm(rrr) + ": " + Error$(rrr))
        End If
      Loop Until l$ = "" Or EOF(n) Or rrr <> 0
      Close #n
    End If
    'The server has finished sending the header if it sends the sequence <CRLF>.<CRLF>
hdrout6767:

End Function

Sub cmdViewfromFile()
Dim o%, i As Integer, rd As Integer, l$, strMessage As String

Call form1.dbg2f("cmdViewfromFile()")
For i = 1 To listMessages.ListItems.Count
  If (listMessages.ListItems(i).Selected = True) Then
    o% = FreeFile
    Call form1.dbg2f("opening: " + listMessages.ListItems(i).SubItems(LVCFILE))
    Open listMessages.ListItems(i).SubItems(LVCFILE) For Input As #o%
    Call form1.dbg2f("reading ...")
    While Not EOF(o%) And rd < 30000
      Line Input #o%, l$
      rd = rd + Len(l$)
      If strMessage <> "" Then strMessage = strMessage + vbCrLf
      strMessage = strMessage + l$
    Wend
    Call form1.dbg2f("read " + trm(rd) + " bytes")
    Close #o%
    listMessages.Visible = False
    txtMessage.Visible = True
    cmdDelete.Enabled = False
    cmdViewl.Caption = "Schliessen"
    On Error Resume Next
    txtMessage.text = strMessage
    Exit Sub
  End If
Next

End Sub

Sub getMsgfromFile()
Dim i As Integer, rrr, o%, p%, hd%, f%, idx%, l$, X, cnt%, curr%, n%, dhwr As String
Dim r As ADODB.Recordset, rc As Integer, eid$, u$, frm$, up$, c_c$, xx$, cc$
Dim dh As ADODB.Recordset, sli$, r1c$, j%, w$, P1%, nn$, vn$, rtmp As ADODB.Recordset
Dim intResultCode As Integer, z$, lw$, vncv$, pf%
Dim intNum As Integer, frome$, hid$, msgto$, trgnum$
Dim strMessage As String, fn$, msgid$, dn$, sw$, bnd$, ucf$, bag$
Dim hd_from$, bdcnt%, nowr%, nmget%, c$, gl As ListItem, kcount
Dim from$, sbj$, dtg$, lvitem, mlcl$, mlclf$, mlc$, adl$
Dim mymailadr As String, toadr$, toadre$, betr$, fn0$, dbid$

Dim d2infile As String, d2insub As String
d2infile = "Frmmain": d2insub = "getMsgfromFile"
mlcl$ = form1.getusersetting("mailserver")
mymailadr = form1.getusersetting("email", "")
If listMessages.ListItems.Count = 0 Then Exit Sub
u$ = form1.getuserid()
dn = form1.mylocaldatadir() + "\mail\inbox"
dn = form1.getusersetting("mailinboxdir", dn)

cnt% = 0
For i = 1 To listMessages.ListItems.Count
  If (listMessages.ListItems(i).Selected = True) Then cnt% = cnt% + 1
Next i%
curr% = 0
pgb1.Visible = True
pgb1.Max = listMessages.ListItems.Count
If rrr <> 0 Then Exit Sub
pgb2.Visible = True
pgb2.Max = 100
pgb2.value = 0
For i = 1 To listMessages.ListItems.Count
  pgb1.value = i: DoEvents
  Set lvitem = listMessages.ListItems(i)
  If (listMessages.ListItems(i).Selected = True) Then
    curr% = curr% + 1
    dhwr = "|"
    lblStatus.Caption = "lade " + trm(curr%) + transe(" von ") + trm(cnt%)
    MousePointer = 11
    DoEvents
    'Set lvItem = listMessages.ListItems(i)
    from$ = listMessages.ListItems(i)
    from$ = strrepl(strrepl(from$, """", ""), "'", "")
    frome$ = from$
    If InStr(frome$, "<") > 0 Then
      frome$ = Mid$(frome$, InStr(frome$, "<") + 1)
      frome$ = Left$(frome$, InStr(frome$, ">") - 1)
    End If
    If InStr(frome$, "(") > 0 Then
      frome$ = trm(Left$(frome$, InStr(frome$, "(") - 1))
    End If
    toadr$ = lvitem.SubItems(LVCTO)
    toadr$ = strrepl(strrepl(toadr$, """", ""), "'", "")
    toadre$ = toadr$
    If InStr(toadre$, "<") > 0 Then
      toadre$ = Mid$(toadre$, InStr(toadre$, "<") + 1)
      toadre$ = Left$(toadre$, InStr(toadre$, ">") - 1)
    End If
    If InStr(toadre$, "(") > 0 Then
      frome$ = trm(Left$(toadre$, InStr(toadre$, "(") - 1))
    End If
    sbj$ = strrepl(lvitem.SubItems(LVCSUBJECT), """", "")
    sbj$ = strrepl(sbj$, "'", "´")
    dtg$ = lvitem.SubItems(LVCDATE)
    cc$ = lvitem.SubItems(LVCCC)
    msgid$ = strrepl(form1.getuserid() + "-" + lvitem.SubItems(LVCID), "'", "")
    If Len(msgid$) > 200 Then msgid$ = form1.newid("mailsafe", "id", 80)
    If wvlt% = 1 Then
      Call form1.new2do(u$, u$, "Nachricht ID:" + msgid$, sbj$ + " (" + frome$ + ")", datum2sql(Date), 0, 0, "", 0)
    End If
    DoEvents
'Get the entire message and put it in strMessageFile
    c$ = "select * from mailsafe where id='" & strrepl(msgid$, "'", "") & "' and frm='" & strrepl(from$, "'", "") & "'"
    fn$ = ""
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then fn$ = r!message
    If fn$ = "" Or nexist(fn$) Then
      fn$ = form1.myuniqueinboxname()
      If Len(from$) > 230 Then from$ = Left$(from$, 230)
      If Len(sbj$) > 230 Then sbj$ = Left$(sbj$, 230)
      dbid$ = strrepl(strrepl(msgid$, """", ""), "'", "")
      c$ = emailonly(from$): If c$ <> "" Then from$ = c$
      c$ = "insert into mailsafe (id,frm,subject,message,erstellt,owner) values('" + _
         dbid$ & "','" & from$ & "','" & sbj$ & "','" & fn$ & "','" & datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "')"
      Call form1.sqlqry(c$)
      If Not form1.isfieldmissing("mailsafe", "otpcc") Then
        c$ = emailonly(cc$): If c$ <> "" Then cc$ = c$
        c$ = "update mailsafe set optcc='" + cc$ + "' where id='" + dbid$ + "'"
        Call form1.sqlqry(c$)
      End If
      If Not form1.isfieldmissing("mailsafe", "optan") Then
        c$ = "update mailsafe set optan='" + Left(toadr$, 240) + "' where id='" + dbid$ + "'"
        Call form1.sqlqry(c$)
      End If
'update dochist
      Call form1.dbg2f("from=" + frome$ + ", me=" + mymailadr)
      If domainofemail(frome$) <> domainofemail(trm(mymailadr)) Then
        c$ = "SELECT * FROM adresse where trim(lcase(email))='" + LCase(frome$) + "'"
        betr$ = "Emaileingang"
      Else
        c$ = "SELECT * FROM adresse where trim(lcase(email))='" + LCase(toadre$) + "'"
        betr$ = "Emailausgang"
      End If
      Call form1.dbg2f(c$)
      Set dh = New ADODB.Recordset
      dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not dh.EOF
'        Debug.Print dh!id; ",-1"
        hid$ = form1.newid("dochist", "id", 19)
        If InStr(dhwr, "|" + dh!id + "-1|") = 0 Then
        c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" + _
            hid$ & "','" & dh!id & "','-1','" & fn$ & "','" + _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','" + Left$(sbj$, 79) & "','" & msgid$ & "','" + betr$ + "')"
        Call form1.sqlqry(c$)
        dhwr = dhwr + dh!id + "-1|"
        End If
        dh.MoveNext
      Wend

      If domainofemail(frome$) <> domainofemail(trm(mymailadr)) Then
        c$ = "SELECT ID FROM kontakt where  trim(lcase(email))='" + LCase(frome$) + "'"
        trgnum$ = frome$
        betr$ = "Emaileingang"
      Else
        c$ = "SELECT ID FROM kontakt where trim(lcase(email))='" + LCase(toadre$) + "'"
        trgnum$ = toadre$
        betr$ = "Emailausgang"
      End If
      Set dh = New ADODB.Recordset
      dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not dh.EOF
        eid$ = form1.getadridbykontaktid(dh!id)
        If InStr(dhwr, "|" + eid + dh!id + "|") = 0 Then
        c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" + _
            form1.newid("dochist", "id", 19) & "','" & eid$ & "','" & dh!id & "','" & fn$ & "','" + _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','" & Left$(sbj$, 79) & "','" & msgid$ & "','" + betr$ + "')"
        Call form1.sqlqry(c$)
        dhwr = dhwr + eid + dh!id + "|"
        End If
        dh.MoveNext
      Wend
      
      If Not form1.isfieldmissing("opt_allenummern", "id") Then
        c$ = "SELECT vid,kid FROM opt_allenummern where num='" + trgnum$ + "' and numtyp='email'"
        Call form1.dbg2f(c$)
        Set dh = New ADODB.Recordset
        dh.CursorLocation = adUseServer
        rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not dh.EOF
          If InStr(dhwr, "|" + trm(dh!vid) + dh!kid + "|") = 0 Then
          c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" + _
            form1.newid("dochist", "id", 19) & "','" & trm(dh!vid) & "','" & dh!kid & "','" & fn$ & "','" + _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() & "','" & Left$(sbj$, 79) & "','" & msgid$ & "','" + betr$ + "')"
          Call form1.sqlqry(c$)
          End If
          dh.MoveNext
        Wend
      End If
    
    End If
    bbox.Caption = "download": DoEvents
    fn0$ = listMessages.ListItems(i).SubItems(LVCFILE)
    'set up volltextsuche
    bbox.Caption = "Volltextindex": DoEvents
    If exist(fn0$) = 1 Then
      sli$ = ""
      o% = FreeFile
      Open fn0$ For Input As #o%
      pf% = FreeFile
      Open fn$ For Output As #pf%
      hd% = 1
      While Not EOF(o%)
        Line Input #o%, l$
        Print #pf%, l$
        l$ = trm(l$)
        If hd% = 1 Then
          l$ = LCase(l$)
        End If
        If l$ = "" And hd% = 1 Then
          hd% = 0
        Else
          If hd% = 0 And InStr(l$, " ") > 0 Then
            r1c$ = ""
            For j% = 1 To Len(l$)
              z$ = Mid$(l$, j%, 1)
              If (z$ >= "0" And z$ <= "9") Or (z$ >= "a" And z$ <= "z") Or (z$ >= "A" And z$ <= "Z") Then
                r1c$ = r1c$ + z$
              Else
                r1c$ = r1c$ + " "
              End If
            Next j%
            l$ = r1c$
            While Len(l$) > 0
              w$ = mkalphanum(LCase(word1(l$)))
              P1% = Len(w$)
              If P1% > 0 Then
                l$ = trm(Mid(l$, P1% + 1))
                If P1% > 2 And P1% < 30 Then
                  sli$ = trm(sli$ & " " & w$)
                End If
              Else
                l$ = trm(Mid$(l$, 2))
              End If
            Wend
          End If
        End If
      Wend
      Close #pf%
      Close #o%
    End If
    bbox.Caption = "update db": DoEvents
    rc = 0
    c$ = "update mailsafe set header='" & rc & "' where id='" & msgid & "' and frm='" & from$ & "'"
    Call form1.sqlqry(c$)
    If sli$ <> "" Then
      sli$ = strrepl(sli$, "'", "")
      sli$ = strrepl(sli$, """", "")
      'c$ = "update mailsafe set volltext='" & sli$ & "' where id='" & msgid & "' and instr(lcase(frm),'" & LCase(from$) & "')>0"
      c$ = "update mailsafe set volltext='" & sli$ & "' where id='" & msgid & "'"
      Call form1.sqlqry(c$)
    End If
    If (rc <> 0 And rc <> 58) Then
      MsgBox ("Error occured while getting the message: Error " & rc)
      Exit Sub
    End If
    bbox.Caption = "client feed": DoEvents
    mlclf$ = strrepl(form1.getusersetting("netscape47inbox"), """", "")
    mlc$ = strrepl(form1.getusersetting("mailclient"), """", "")
    If mymailadr = frome$ Then
      GoTo notthismsg
    End If
    If InStr(LCase(mlc$), "netscape") > 0 Or LCase(form1.getusersetting("Mozillaclient")) = "ja" Then mlcl$ = "NETSCAPE47"
    If mlcl$ = "NETSCAPE47" Then
      If exist(mlclf$) > 0 And form1.getusersetting("feed2inbox", "ja") = "ja" Then
        o% = FreeFile
        Open fn0$ For Input As #o%
        p% = FreeFile
        Open mlclf$ For Append As #p%
        Print #p%, "From " & from$
        While Not EOF(o%)
          Line Input #o%, l$
          If Left(LCase(l$), 9) <> "x-mozilla" Then Print #p%, l$
          DoEvents
        Wend
        Close #o%
        Close #p%
      End If
    End If
    nmget% = 1
    bbox.Caption = "": DoEvents
    MousePointer = 0
    DoEvents
    msgto$ = lvitem.SubItems(LVCTO)
    n% = 0
  End If
notthismsg:
Next i
pgb1.Visible = False
pgb2.Visible = False
bbox.Caption = "": DoEvents
cmdDelete_ask% = 0
Call cmdDelete_Click
cmdDelete_ask% = 1
Call listMessages.SetFocus
If mlcl$ = "NETSCAPE47" And form1.getusersetting("feed2inbox", "ja") = "ja" Then
  mlcl$ = form1.getusersetting("mailclient")
  If exist(word1(mlcl$)) > 0 Then
    X = Shell(mlcl$, 1)
    wait 1
    SendKeys "%dk", 1
  End If
End If

End Sub

Sub mrkbysize()
Dim i As Integer

Dim from$, frome$, c$, lvitem, marksize As Long, rrr
'd2infile = "Frmmain": d2insub = "Command2_Click"

On Error Resume Next
marksize = CLng(form1.getusersetting("markmailsize", 0))
rrr = Err
On Error GoTo 0
If rrr = 0 And marksize > 0 Then

Call listMessages.SetFocus
For i = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i).Selected = True Then Exit Sub
Next i
For i = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i).SubItems(LVCSIZE) > marksize Then
    listMessages.ListItems(i).Selected = True
  Else
    listMessages.ListItems(i).Selected = False
  End If
  DoEvents
Next i

End If
End Sub

