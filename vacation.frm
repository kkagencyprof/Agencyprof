VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form responder 
   Caption         =   "Autoresponder einrichten"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   ScaleHeight     =   6735
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox vscnp 
      Height          =   285
      Left            =   8640
      TabIndex        =   22
      Text            =   "|/usr/local/bin/metmail"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CheckBox vscn 
      Caption         =   "Virenscannerer aktiv (Mailserverkonfiguration erforderlich)"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CheckBox anmich 
      Caption         =   "an mich"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   2640
      Value           =   1  'Aktiviert
      Width           =   1215
   End
   Begin VB.CheckBox o2 
      Caption         =   "Autoresponder aktiv"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      Picture         =   "vacation.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Speichern (Daten übertragen)"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox vacp 
      Height          =   285
      Left            =   8640
      TabIndex        =   16
      Text            =   "|/usr/bin/vacation"
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox ffwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CheckBox o1 
      Caption         =   """Reply-To"" beachten"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "vacation.frx":03A7
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1080
      Picture         =   "vacation.frx":05F7
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "erneut laden"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   3735
   End
   Begin VB.ListBox logwin 
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   1920
      TabIndex        =   6
      Top             =   4320
      Width           =   5895
   End
   Begin VB.ComboBox hid 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "1"
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox subjct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   480
      Width           =   7575
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   7320
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "FTP-Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Weiterleitung(en)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "neu:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "jetzt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tage zw. den Antworten je Empfänger:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Betreff:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "responder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim updhid%
Dim ftphost$, ftpport$, ftpuser$, ftppass$
Dim ftprdir$

Private Sub anmich_Click()
'd2infile = "vacation": d2insub = "anmich_Click"
Call updtext3
End Sub

Private Sub Command2_Click()
'd2infile = "vacation": d2insub = "Command2_Click"
Unload Me

End Sub

Private Sub Command4_Click()
Dim nError As Integer
Dim strLocalFile As String
Dim strRemoteFile As String
Dim bTransfered As Integer

Static bCanceled As Integer
Dim o%, l$, l1$, p%, z$, vac$, vsc$

msg.Text = ""
    If form1.sockCmd.Connected Then
        nError = form1.sockCmd.Disconnect
        If (nError <> 0) Then
            MsgBox "Unable to disconnect.  Error " & nError
        End If
    Else
        Dim strHostName As String
        Dim strDirectory As String
        Dim nRemotePort As Integer

        logwin.Clear

        strHostName = ftphost$
        nRemotePort = Val(ftpport$)

        '
        ' Establish the connection with the FTP server
        '
        If Not FtpConnect(strHostName, nRemotePort) Then
            logwin.AddItem "Verbindung mit " & strHostName & " fehlgeschlagen"
            logwin.ListIndex = logwin.ListCount - 1
            Exit Sub
        End If

        '
        ' Login to the server using the supplied username
        ' and password
        '
        If Not FtpLogin(Trim$(ftpuser$), ftppass$) Then
            form1.sockCmd.Action = SOCKET_DISCONNECT
            logwin.AddItem "Benutzername / Passwort ungültig"
            logwin.ListIndex = logwin.ListCount - 1
            Exit Sub
        End If

        '
        ' Get the current working directory
        '
        If FtpGetDirectory(strDirectory) Then
            ftprdir$ = strDirectory
        Else
            form1.sockCmd.Disconnect
            logwin.AddItem "Remote-Verzeichnis nicht abrufbar"
            logwin.ListIndex = logwin.ListCount - 1
            Exit Sub
        End If
    End If
    logwin.AddItem "Verbunden mit " & ftphost$
    logwin.ListIndex = logwin.ListCount - 1

    '
    ' If the data socket is in use, then a file transfer
    ' is in progress
    '
    If form1.sockData.State <> SOCKET_UNUSED Then
        FtpCancel
        bCanceled = True
        Exit Sub
    End If

    strLocalFile = form1.mydatadir() & "\forward"
    strRemoteFile = ".forward"

    On Error Resume Next
    Kill strLocalFile
    On Error GoTo 0
    bTransfered = FtpGetFile(strLocalFile, strRemoteFile)
    If Not bTransfered Then
        logwin.AddItem "Übertragung abgebrochen"
        logwin.ListIndex = logwin.ListCount - 1
    End If

  Text2.Text = ""
  If exist(strLocalFile) <> 0 Then
  o% = FreeFile
  Open strLocalFile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
    p% = InStr(l$, Chr$(10))
    If p% > 0 Then
      l1$ = trm(Left(l$, p% - 1))
      l$ = trm(Mid$(l$, p% + 1))
    Else
      If Left$(l$, 1) = "\" Then
        l1$ = trm(l$)
        l$ = ""
      Else
        l1$ = ""
        l$ = trm(l$)
      End If
    End If
    Text2.Text = l1$
    o2.value = 0
    p% = InStr(l1$, "-r"): If p% > 1 Then o1.value = 1
    p% = InStr(l1$, "-t"): If p% > 1 Then Text1.Text = Mid$(l1$, p% + 2, 1)
    p% = InStr(l1$, "vacation")
    If p% > 1 Then
      Do
        z$ = Mid$(l1$, p%, 1)
        p% = p% - 1
      Loop Until z$ = """" Or z$ = ";" Or p% < 1
      If p% > 0 Then
        vac$ = trm(word1(Mid$(l1$, p% + 2)))
        vacp.Text = vac$
        o2.value = 1
      End If
    End If
    p% = InStr(l1$, "metmail")
    If p% > 1 Then
      Do
        z$ = Mid$(l1$, p%, 1)
        p% = p% - 1
      Loop Until z$ = """" Or z$ = ";" Or p% < 1
      If p% > 0 Then
        vsc$ = trm(word1(Mid$(l1$, p% + 2)))
        vscnp.Text = vsc$
        vscn.value = 1
      End If
    End If
    While trm(l$) <> ""
      p% = InStr(l$, Chr$(10))
      If p% > 0 Then
          l1$ = trm(Left(l$, p% - 1))
          l$ = trm(Mid$(l$, p% + 1))
      Else
          l1$ = trm(l$)
          l$ = ""
      End If
      If ffwd.Text <> "" Then ffwd.Text = ffwd.Text & vbCrLf
      ffwd.Text = ffwd.Text & l1$
    Wend
    While Not EOF(o%)
      Line Input #o%, l$
      If ffwd.Text <> "" Then ffwd.Text = ffwd.Text & vbCrLf
      ffwd.Text = ffwd.Text & l$
    Wend
  End If
  Close #o%
  End If

    strLocalFile = form1.mydatadir() & "\vacation.msg"
    strRemoteFile = ".vacation.msg"

    On Error Resume Next
    Kill strLocalFile
    On Error GoTo 0
    bTransfered = FtpGetFile(strLocalFile, strRemoteFile)
    If Not bTransfered Then
        logwin.AddItem "Übertragung abgebrochen"
        logwin.ListIndex = logwin.ListCount - 1
    End If
  o% = FreeFile
  If exist(strLocalFile) <> 0 Then
  Open strLocalFile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$ ': l$ = strrepl(l$, Chr$(10), "")
    p% = InStr(l$, Chr$(10))
    If p% > 0 Then
      l1$ = trm(Left(l$, p% - 1))
      l$ = trm(Mid$(l$, p% + 1))
    Else
      l1$ = trm(l$)
      l$ = ""
    End If
    If Left(l1$, 9) = "Subject: " Then l1$ = Mid$(l1$, 10)
    subjct.Text = l1$
    While trm(l$) <> ""
      p% = InStr(l$, Chr$(10))
      If p% > 0 Then
        l1$ = trm(Left(l$, p% - 1))
        l$ = trm(Mid$(l$, p% + 1))
      Else
        l1$ = trm(l$)
        l$ = ""
      End If
      If msg.Text <> "" Then msg.Text = msg.Text & vbCrLf
      msg.Text = msg.Text & l1$
    Wend
    While Not EOF(o%)
      Line Input #o%, l$
      If msg.Text <> "" Then msg.Text = msg.Text & vbCrLf
      msg.Text = msg.Text & l$
    Wend
  End If
  Close #o%
  End If
Call updtext3

End Sub

Private Sub Command5_Click()
Dim bTransfered As Integer
Dim strLocalFile As String, strRemoteFile As String, o%, fn$

'd2infile = "vacation": d2insub = "Command5_Click"
bTransfered = False
o% = FreeFile
fn$ = form1.mydatadir() & "\forward"
strLocalFile = fn$
strRemoteFile = ".forward"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, strrepl(Text3.Text, vbCrLf, Chr$(10));
Close #o%
bTransfered = FtpPutFile(strLocalFile, strRemoteFile)
If Not bTransfered Then
  logwin.AddItem strRemoteFile & ": Übertragung fehlgeschlagen"
Else
  logwin.AddItem strRemoteFile & ": Übertragung ok"
End If
logwin.ListIndex = logwin.ListCount - 1

fn$ = form1.mydatadir() & "\vacation.msg"
strLocalFile = fn$
strRemoteFile = ".vacation.msg"
Open fn$ For Output As #o%
Print #o%, "Subject: " & subjct.Text; Chr$(10); Chr$(10);
Print #o%, strrepl(msg.Text, vbCrLf, Chr$(10));
Close #o%
bTransfered = FtpPutFile(strLocalFile, strRemoteFile)
If Not bTransfered Then
  logwin.AddItem strRemoteFile & ": Übertragung fehlgeschlagen"
Else
  logwin.AddItem strRemoteFile & ": Übertragung ok"
End If
logwin.ListIndex = logwin.ListCount - 1

End Sub

Private Sub ffwd_Change()
'd2infile = "vacation": d2insub = "ffwd_Change"
Call updtext3
End Sub

Private Sub Form_Load()
'd2infile = "vacation": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
updhid% = 1
Call ftpsetlogwin("responder")
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
o1.BackColor = form1.cleancolor()
o2.BackColor = o1.BackColor
vscn.BackColor = o1.BackColor
anmich.BackColor = o1.BackColor
BackColor = o1.BackColor
Call updtext3
Show
End Sub

Private Sub Form_Resize()
'd2infile = "vacation": d2insub = "Form_Resize"
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "vacation": d2insub = "Form_Unload"
If form1.sockCmd.Connected Then form1.sockCmd.Abort
If form1.sockData.Listening Or form1.sockData.Connected Then form1.sockData.Abort
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0


End Sub

Private Sub hid_Change()
Dim c$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "vacation": d2insub = "hid_Change"
If updhid% = 0 Then Exit Sub
updhid% = 0

If form1.sockCmd.Connected Then form1.sockCmd.Abort
If form1.sockData.Listening Or form1.sockData.Connected Then form1.sockData.Abort
Text2.Text = ""
Text3.Text = ""
msg.Text = ""
ffwd.Text = ""
subjct.Text = ""
ftphost$ = "": ftpport$ = "": ftpuser$ = "": ftppass$ = ""
c$ = "SELECT auftritthigru.auftrittsid from auftritthigru " + _
     "WHERE (((auftritthigru.auftrittstyp)='Autoresponder') AND ((auftritthigru.FeldName)='ftp-user') AND ((auftritthigru.FeldDaten)='" & hid.Text & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, dbOpenDynaset, dbReadOnly
If r.EOF Then Exit Sub
ftpuser$ = hid.Text
c$ = "SELECT feldname,felddaten from auftritthigru " + _
     "WHERE (((auftrittsid)='" & r!auftrittsid & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, dbOpenDynaset, dbReadOnly
While Not r.EOF
  Select Case LCase(r!feldname)
    Case "ftp-host": ftphost$ = r!felddaten
    Case "ftp-port": ftpport$ = r!felddaten
    Case "ftp-passwort": ftppass$ = r!felddaten
    Case Else:
  End Select
  r.MoveNext
Wend
Call updtext3
Call Command4_Click
updhid% = 1

End Sub

Private Sub hid_Click()
'd2infile = "vacation": d2insub = "hid_Click"
Call hid_Change
End Sub

Private Sub hid_DropDown()
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "vacation": d2insub = "hid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT * FROM auftritthigru where auftrittstyp='Autoresponder' and feldname='ftp-user'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, dbOpenDynaset, dbReadOnly
hid.Clear
While Not r.EOF
  hid.AddItem r!felddaten
  r.MoveNext
Wend
MousePointer = 0

End Sub
Sub updtext3()
Dim opt$, vac$, vsc$

'd2infile = "vacation": d2insub = "updtext3"
If anmich.value = 1 Then
  Text3.Text = "\" & hid.Text
Else
  Text3.Text = ""
End If
vsc$ = ""
If vscn.value = 1 Then
  vsc$ = "|/usr/local/bin/metmail " + hid.Text
End If
If o2.value = 1 Then
  vac$ = trm(vacp.Text) & " "
  opt$ = "-j -t" & trm(Text1.Text) & " "
  If o1.value <> 0 Then opt$ = opt$ & "-r "
  If trm(Text3.Text) <> "" Then
    Text3.Text = Text3.Text & ", "
  End If
  Text3.Text = Text3.Text & """" & vac$ & " " & opt$ & hid.Text & """"
End If
If trm(ffwd.Text) <> "" Then
  If trm(Text3.Text) <> "" Then Text3.Text = Text3.Text & vbCrLf
  Text3.Text = Text3.Text & ffwd.Text
End If
If vsc$ <> "" Then
  If trm(Text3.Text) <> "" Then Text3.Text = Text3.Text & vbCrLf
  Text3.Text = Text3.Text + """" + vsc$ + """"
End If
End Sub

Private Sub o1_Click()
'd2infile = "vacation": d2insub = "o1_Click"
Call updtext3
End Sub

Private Sub o2_Click()
'd2infile = "vacation": d2insub = "o2_Click"
If o2.value = 0 Then
  Text1.Enabled = False
  o1.Enabled = False
Else
  Text1.Enabled = True
  Text1.Text = 1
  o1.Enabled = True
  o1.value = 1
End If
Call updtext3
End Sub

Private Sub Text1_Change()
'd2infile = "vacation": d2insub = "Text1_Change"
If Len(Text1.Text) > 1 Then Text1.Text = Left(Text1.Text, 1)
Call updtext3

End Sub

Private Sub vacp_Change()
'd2infile = "vacation": d2insub = "vacp_Change"
Call updtext3
End Sub

Private Sub vscn_Click()
'd2infile = "vacation": d2insub = "vscn_Click"
If vscn.value = 1 Then
  If anmich.value <> 0 Then anmich.value = 0
Else
  If anmich.value <> 1 Then anmich.value = 1
End If
Call updtext3
End Sub
