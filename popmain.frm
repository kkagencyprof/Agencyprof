VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSComCtl.ocx"
Begin VB.Form popmain 
   Caption         =   "Maileingang"
   ClientHeight    =   3570
   ClientLeft      =   4140
   ClientTop       =   3105
   ClientWidth     =   2730
   Icon            =   "popmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   2730
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   4920
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.TextBox autoex 
      Height          =   285
      Left            =   960
      TabIndex        =   36
      Text            =   "0200"
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   2400
      Left            =   6600
      TabIndex        =   34
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Width           =   6255
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   5880
      Picture         =   "popmain.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   5880
      Picture         =   "popmain.frx":0884
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5880
      Picture         =   "popmain.frx":0CC6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   1920
      Width           =   495
   End
   Begin VB.CheckBox noti 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2460
      TabIndex        =   28
      ToolTipText     =   "Die Einstellung im Sendefenster von Agencyprof geht vor."
      Top             =   2040
      Width           =   255
   End
   Begin VB.Timer hidetmr 
      Left            =   5880
      Top             =   0
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Picture         =   "popmain.frx":1108
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Emails senden"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hide"
      Default         =   -1  'True
      Height          =   255
      Left            =   1920
      TabIndex        =   25
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   3960
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pgb2 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox darec 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox tout 
      Height          =   285
      Left            =   3960
      TabIndex        =   17
      Text            =   "30"
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox brlle 
      Height          =   615
      Index           =   1
      Left            =   4680
      Picture         =   "popmain.frx":11BA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox brlle 
      Height          =   615
      Index           =   0
      Left            =   4080
      Picture         =   "popmain.frx":2334
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "popmain.frx":34AE
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "vom Server laden"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5400
      Top             =   0
   End
   Begin VB.TextBox pin 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Mailstatus auf allen Servern, Doppelklick zum Aktualisieren"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "popmain.frx":3E74
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "&Mail testen"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Text            =   "110"
      Top             =   2040
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   1080
      Top             =   2280
   End
   Begin VB.Label Label7 
      Caption         =   "Debug Timer"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   37
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Autoexit"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Die Einstellung im Sendefenster von Agencyprof geht vor."
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bestätigung"
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      ToolTipText     =   "Die Einstellung im Sendefenster von Agencyprof geht vor."
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label outbxcount 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "löschen"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2040
      Width           =   735
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
      TabIndex        =   19
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout:"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label cmdViewl 
      Caption         =   "Zeigen"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6000
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
      Left            =   3000
      TabIndex        =   11
      Top             =   600
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
      TabIndex        =   8
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "Pop Server:"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   915
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Wiederherstellen"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "popmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Private Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Const NIM_ADD = &H0
      Const NIM_MODIFY = &H1
      Const NIM_DELETE = &H2
      Const NIF_MESSAGE = &H1
      Const NIF_ICON = &H2
      Const NIF_TIP = &H4
      Const WM_MOUSEMOVE = &H200
      Const WM_LBUTTONDOWN = &H201     'Button down
      Const WM_LBUTTONUP = &H202       'Button up
      Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Const WM_RBUTTONDOWN = &H204     'Button down
      Const WM_RBUTTONUP = &H205       'Button up
      Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Private nid As NOTIFYICONDATA


Dim nomess%, pcode
Dim poplistok%, ismin As Boolean
Dim prevfire%, t2fl As Boolean
Dim tot As Long
Dim cmdDelete_ask%, gcount%
Dim logintimer As Long
Dim boxread As Long
Dim diskwrite As Long

Public Sub cmdCheckMail_Click()
Dim intNum As Integer
Dim msgn%, intMessageNum As Integer 'the number of messages
Dim anyconn As Boolean

If txtServer = "" Then Exit Sub
anyconn = False

MousePointer = 11
DoEvents
    If (PopConnect(txtServer.Text, Val(txtPort.Text), tot) = POP_SOCKET_ERROR) Then
        DoEvents
        '1 Retry
        Call Form1.dbg2f("connect...")
        If (PopConnect(txtServer.Text, Val(txtPort.Text), tot) = POP_SOCKET_ERROR) Then
          DoEvents
          If nomess% = 0 Then
            MsgBox ("Fehler beim Verbinden mit " + trm(txtServer.Text) + ": Error " + trm(Socket1.LastError))
            List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim Verbinden mit " + trm(txtServer.Text) + ": Error " + trm(Socket1.LastError)
            Call Form1.dbg2f("Fehler beim Verbinden mit " + trm(txtServer.Text) + ": Error " + trm(Socket1.LastError))
          End If
          Socket1.Disconnect
          MousePointer = 0
          GoTo mehldann
        End If
    End If
Call Form1.dbg2f("PopLogin...")
    If (PopLogin(txtUserName.Text, txtPassword.Text) = POP_SOCKET_ERROR) Then
'        MsgBox ("Fehler beim  Login: Error " + trm(Socket1.LastError))
        List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim  Login: Error " + trm(Socket1.LastError)
        Call Form1.dbg2f("Fehler beim  Login: Error " + trm(Socket1.LastError))
        Socket1.Disconnect
        MousePointer = 0
        GoTo mehldann
    End If
Call Form1.dbg2f("PopGetMessageCount...")
    If (PopGetMessageCount(intMessageNum) = POP_SOCKET_ERROR) Then
'        MsgBox ("Fehler beim Mailstatus: Error " + trm(Socket1.LastError))
        List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim Mailstatus: Error " + trm(Socket1.LastError)
        Socket1.Disconnect
        MousePointer = 0
        GoTo mehldann
    End If
    lblStatus.Caption = trm(intMessageNum) + " " + transe("Nachrichten")

    If (PopDisconnect() = POP_SOCKET_ERROR) Then
            MousePointer = 0
            Socket1.Disconnect
        GoTo mehldann
    End If
mehldann:
DoEvents
If List1.ListCount = 0 Or List1.List(0) = transe("PIN fehlt") Then Call rlist1
MousePointer = 0

End Sub

Private Sub Command1_Click()

Socket1.Disconnect
Unload Me
End
End Sub

Private Sub Command16_Click()
Dim mlcl$, outbx$, tr, ccnt As Long, rrr
Dim i As Integer, j As Integer

t2fl = True
mlcl$ = Form1.getusersetting("mailserver")
outbx$ = Form1.myoutbox
  
  ccnt = outbxchk()
  If ccnt > 0 Then
    List1.Clear
    tr = Dir(outbx$ + "\*.amf")
    While tr <> ""
      List1.AddItem tr
      On Error GoTo rrrout33
      tr = Dir
      On Error Resume Next
    Wend
    If List1.ListCount > 0 Then
      If ismin Then
        With nid
         .hIcon = Picture2.Picture
        End With
        Shell_NotifyIcon NIM_MODIFY, nid
      End If
      For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        DoEvents
        If Not nexist(outbx$ + "\" + List1.List(i)) Then
          Call Form1.mailresend(outbx$ + "\" + List1.List(i), noti.value)
          rrr = 0
          On Error Resume Next
          Kill outbx$ + "\" + List1.List(i)
          rrr = Err
          On Error GoTo 0
          Call Form1.logerr("killing mailfile " + outbx$ + "\" + List1.List(i) + "=" + trm(rrr))
          If rrr <> 0 Then
            If Not nexist(outbx$ + "\" + List1.List(i)) Then
              MsgBox "Error (" + trm(rrr) + ") when " + outbx$ + "\" + List1.List(i) + " was to be deleted. - Program will stop."
            End If
            End
          End If
        End If
        ccnt = ccnt - 1
        outbxcount.Caption = trm(ccnt)
      Next i
      If ismin Then
        If Not mailthere() Then
           With nid
            .hIcon = Me.Icon
            .szTip = "Agencyprof Mailtransfer V" & App.Major & "." & App.Minor & " Build #" & App.Revision & vbNullChar
           End With
        Else
           With nid
            .hIcon = Picture3.Picture
            .szTip = "Sie haben Mail!" & vbNullChar
           End With
        End If
        Shell_NotifyIcon NIM_MODIFY, nid
      End If
    End If

End If
rrrout33:
List1.Clear
ccnt = outbxchk()
t2fl = False
End Sub

Private Sub Command2_Click()
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Agencyprof Mailtransfer V" & App.Major & "." & App.Minor & " Build #" & App.Revision & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid
ismin = True
Hide
End Sub

Private Sub Command3_Click()
Dim i As Integer, rrr, o%, p%, hd%, f%, idx%, l$, x, cnt%, curr%, n%
Dim rc As Integer, eid$, u$, frm$, up$, c_c$, xx$
Dim sli$, r1c$, j%, w$, p1%, nn$, vn$, intMessageNum As Integer
Dim intResultCode As Integer, z$, lw$, vncv$
Dim intNum As Integer, frome$, hid$, msgto$
Dim strMessage As String, fn$, dn$, sw$, bnd$, ucf$, bag$, inbx$
Dim hd_from$, bdcnt%, nowr%, nmget%, C$, gl As ListItem, kcount
Dim from$, sbj$, dtg$, lvitem As ListItem, mlcl$, mlclf$, mlc$, adl$, tr

If txtServer = "" Then Exit Sub

If ismin Then
       With nid
        .hIcon = Picture1.Picture
       End With
       Shell_NotifyIcon NIM_MODIFY, nid
End If
mlcl$ = Form1.getusersetting("mailserver")
inbx$ = Form1.myinbox
bdcnt% = 0
tr = Dir(inbx$ + "\*.amf")
While tr <> "" And bdcnt < 10000
  bdcnt% = bdcnt% + 1
  tr = Dir()
Wend
If bdcnt > 9000 Then
  List2.AddItem trm(Date) + " " + trm(Time) + " (no error) Maildir too full, not polling"
  Call Form1.dbg2f(" (no error) Maildir too full, not polling")
  GoTo c3xout
End If
Call l7set
Call tm_start(0)
Call Form1.dbg2f("PopConnect...")
If (PopConnect(txtServer.Text, Val(txtPort.Text), tot) = POP_SOCKET_ERROR) Then
  DoEvents
  If (PopConnect(txtServer.Text, Val(txtPort.Text), tot) = POP_SOCKET_ERROR) Then
'    MsgBox ("Fehler beim Verbinden: Error " + trm(Socket1.LastError))
    List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim Verbinden: Error " + trm(Socket1.LastError)
    GoTo c3xout
  End If
End If
Call Form1.dbg2f("PopLogin...")
If (PopLogin(txtUserName.Text, txtPassword.Text) = POP_SOCKET_ERROR) Then
'            MsgBox ("Fehler beim  Login: Error " + trm(Socket1.LastError))
            Call Form1.dbg2f("Fehler beim  Login: Error " + trm(Socket1.LastError))
            List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim  Login: Error " + trm(Socket1.LastError)
            GoTo c3xout
End If
Call Form1.dbg2f("PopGetMessageCount...")
If (PopGetMessageCount(intMessageNum) = POP_SOCKET_ERROR) Then
'        MsgBox ("Fehler beim Mailstatus: Error " + trm(Socket1.LastError))
        Call Form1.dbg2f("Fehler beim Mailstatus: Error " + trm(Socket1.LastError))
        List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim Mailstatus: Error " + trm(Socket1.LastError)
        GoTo c3xout
End If
cnt% = 0
cnt% = intMessageNum

If cnt% > 0 Then
logintimer = logintimer + tm_stop(0)

curr% = 0
pgb1.Visible = True
pgb1.Max = intMessageNum
If rrr <> 0 Then Exit Sub
pgb2.Visible = True
pgb2.Max = 100
pgb2.value = 0
MousePointer = 11
DoEvents
For i = 1 To intMessageNum
    pgb1.value = i: DoEvents
    curr% = curr% + 1
    lblStatus.Caption = "lade " + trm(curr%) + transe(" von ") + trm(cnt%)
    Call Form1.dbg2f("lade " + trm(curr%) + transe(" von ") + trm(cnt%))
    Call tm_start(0)
    DoEvents
    bbox.Caption = "download": DoEvents
    dtg$ = datum2sql(Date)
    p% = 0
    Do
      fn$ = inbx$ + "\" + dtg$ + "-" + strrepl(trm(Time), ":", "") + "-" + trm(p%) + ".amf"
      p% = p% + 1
    Loop Until nexist(fn$)
    Call Form1.dbg2f("Erstelle Datei " + fn$)
    n% = FreeFile
    Open fn$ + ".lck" For Output As #n%: Close #n%
    Call Form1.dbg2f("calling PopGetLongMessage " + trm(i))
    diskwrite = diskwrite + tm_stop(0)
    rc = PopGetLongMessage(i, fn$)
    Call Form1.dbg2f("returned fromcalling PopGetLongMessage")
    On Error Resume Next
    Kill fn$ + ".lck"
    On Error GoTo 0
    'set up volltextsuche
    If (rc = POP_SOCKET_ERROR) Then
'      MsgBox ("Error occured while getting the message: Error " & Socket1.LastError)
      List2.AddItem trm(Date) + " " + trm(Time) + " Error occured while getting the message: Error " & Socket1.LastError
      GoTo c3xout
    End If
    If darec.value <> 0 Then
      Call Form1.dbg2f("Deleting Message " + trm(i))
      If (PopDeleteMessage(i) = POP_SOCKET_ERROR) Then
'          MsgBox ("Fehler beim Löschen: Error " + trm(Socket1.LastError))
          List2.AddItem trm(Date) + " " + trm(Time) + " Fehler beim Löschen: Error " + trm(Socket1.LastError)
          GoTo c3xout
      End If
    End If
    bbox.Caption = ""
    Call Form1.dbg2f("Datei fertig")
    DoEvents
notthismsg:
Next i


End If
DoEvents
If (PopDisconnect() = POP_SOCKET_ERROR) Then
'        MsgBox ("Error occured during disconnect: Error " & Socket1.LastError)
        Call Form1.dbg2f("disconnected with error " & Socket1.LastError)
        GoTo c3xout
End If
c3xout:
Call l7set
Call Form1.dbg2f("Mail retrieve done")
If ismin Then
    If Not mailthere() Then
       With nid
        .hIcon = Me.Icon
        .szTip = "Agencyprof Mailtransfer V" & App.Major & "." & App.Minor & " Build #" & App.Revision & vbNullChar
       End With
    Else
       With nid
        .hIcon = Picture3.Picture
        .szTip = "Sie haben Mail!" & vbNullChar
       End With
    End If
       Shell_NotifyIcon NIM_MODIFY, nid
End If
Socket1.Disconnect
pgb1.Visible = False
pgb2.Visible = False
bbox.Caption = "": DoEvents
lblStatus.Caption = ""
MousePointer = 0

End Sub

Private Sub Form_Load()
Dim colHeader As ColumnHeader
Dim rrr, stst$, i%, s%, klrv%, C$

gcount% = 10
t2fl = False
Socket1.AutoResolve = False
Socket1.Blocking = True
Socket1.Binary = False   'Read a line at a time
Socket1.Protocol = IPPROTO_IP
prevfire% = 0
poplistok% = 0
pin.Visible = False
Label2.Visible = False
List1.Visible = False
ismin = False
cmdDelete_ask% = 1
logintimer = 0
boxread = 0
diskwrite = 0

If Form1.popentries > 0 Then
  poplistok% = 1
  pin.Visible = True
  Label2.Visible = True
  List1.Visible = True
End If

nomess% = 0

tout.Text = "300"
tot = Val(tout.Text)
lblStatus.Caption = "ok"
klrv% = Val(Form1.delrecvd)
If klrv% <> 0 Then
  klrv% = 1
  darec.value = klrv%
End If

popmain.Caption = transe("Maileingang")
darec.Caption = transe("Nach Empfang löschen")
Command3.ToolTipText = transe("vom Server laden")
List1.ToolTipText = transe("Mailstatus auf allen Servern, Doppelklick zum Aktualisieren")
Command1.ToolTipText = transe("Dieses Formular schliessen")
cmdCheckMail.Caption = transe("&Mail testen")
Label4.Caption = transe("löschen")
Label3.Caption = transe("TOut:")
Label2.Caption = transe("PIN:")
Label1.Caption = transe("Password:")
lblUserName.Caption = transe("User")
lblPort.Caption = transe("Port:")
lblServer.Caption = transe("Pop Server:")
klrv% = Val(Form1.mylastFormVar(Me.Name, "delrecvd", "0"))
If klrv% <> 0 Then
  klrv% = 1
  darec.value = klrv%
End If
klrv% = Val(Form1.mylastFormVar(Me.Name, "smtpnotify", "0"))
If klrv% <> 0 Then
  klrv% = 1
  noti.value = klrv%
End If
Me.Top = Form1.mylasttop(Me.Name)
Me.Left = Form1.mylastleft(Me.Name)
Show
Call Form1.dbg2f("POPClient started")
hidetmr.Interval = 2000
hidetmr.Enabled = True
End Sub


Public Sub chkm()
DoEvents
Call cmdCheckMail_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load closing
DoEvents
Call Form1.setmylasttop(Me.Name, Me.Top)
Call Form1.setmylastleft(Me.Name, Me.Left)
Hide
On Error Resume Next
Kill Form1.mylocaldatadir() & "\debug2file_" & Form1.getuserid() & "_popmain_Socket1.txt"
On Error GoTo 0
On Error GoTo exuld
Call Form1.setmylasttop(Me.Name, Me.Top)
Call Form1.setmylastleft(Me.Name, Me.Left)
exuld:
On Error GoTo 0
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub hidetmr_Timer()
hidetmr.Enabled = False
If Form1.getusersetting("popclientverstecken", "ja") = "ja" Then Call Command2_Click

End Sub

Private Sub List1_Click()
Dim id$, i%, u$, l$, o%
Dim aKey() As Byte, rc$, rrr
Dim rid$, pw$
Dim rserver$
Dim ruser$
Dim rpsswd$
Dim rport$, ap$

If List1.ListIndex < 0 Then Exit Sub
id$ = List1.List(List1.ListIndex)
id$ = Mid$(id$, InStr(id$, "auf ") + 4)
u$ = Form1.getuserid()
ap$ = trm(pin.Text)
If ap$ <> "" Then
  o% = FreeFile
  If Not nexist(Form1.hpth$ + "\poplist.agp") Then
  Open Form1.hpth$ + "\poplist.agp" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    rid$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    If rid$ = u$ + "_" + id$ Then
      rserver$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      ruser$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      rpsswd$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      rport$ = l$
      pw$ = decrypt(rpsswd$, ap$)
      txtServer.Text = rserver$
      txtUserName.Text = ruser$
      txtPassword.Text = pw$
      txtPort.Text = rport$
      DoEvents
      Call cmdCheckMail_Click
    End If
  Wend
  End If
End If

DoEvents
End Sub

Private Sub List1_DblClick()
'd2infile = "popmain": d2insub = "List1_DblClick"
List1.Clear
End Sub

Private Sub noti_Click()
Call Form1.setmylastFormVar(Me.Name, "smtpnotify", trm(noti.value))

End Sub

Private Sub Timer1_Timer()
Dim u$, ap$, C$, o%, l$, pw$
Dim aKey() As Byte, rc$, rrr
Dim rid$
Dim rserver$
Dim ruser$
Dim rpsswd$
Dim rport$

Timer1.Enabled = False
Timer1.Interval = 0
u$ = Form1.getuserid()
ap$ = trm(pin.Text)
If ap$ <> "" Then
  o% = FreeFile
  If Not nexist(Form1.hpth$ + "\poplist.agp") Then
    Open Form1.hpth$ + "\poplist.agp" For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      rid$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      rserver$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      ruser$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      rpsswd$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      rport$ = l$
      If rid$ <> "PDFServer" Then
        pw$ = decrypt(rpsswd$, ap$)
Call Form1.dbg2f("teste server: " + rserver$ + ", user=" + ruser$)
        txtServer.Text = rserver$
        txtUserName.Text = ruser$
        txtPassword.Text = pw$
        txtPort.Text = rport$
        DoEvents
        Call cmdCheckMail_Click
      End If
    Wend
  End If
End If

End Sub

Function outbxchk() As Long
Dim outbx$, tr, ccnt As Long

outbxchk = 0
outbx$ = Form1.myoutbox
ccnt = 0
tr = Dir(outbx$ + "\*.amf")
While tr <> ""
  ccnt = ccnt + 1
  tr = Dir()
Wend
If ccnt > 0 Then
  Command16.Enabled = True
  outbxcount.Caption = trm(ccnt)
Else
  Command16.Enabled = False
  outbxcount.Caption = "0"
End If
outbxchk = ccnt

End Function

Private Sub Timer2_Timer()
Dim i As Integer, ccnt As Long, outbx$, tr, tn As String, tb As String

If t2fl Then Exit Sub
Timer2.Enabled = False
DoEvents
tb = strrepl(trm(autoex.Text), ":", "")
If tb <> "" Then
  tn = strrepl(trm(Time), ":", "")
  If Left(tn, Len(tb)) = tb Then
    End
  End If
End If
If nexist(Form1.setfn$) Then
  Unload Me
  Exit Sub
End If
gcount% = gcount% + 1
If Not Form1.iamserver Then
  outbx$ = Form1.myoutbox
  tr = Dir(outbx$ + "\lock.lck")
  If tr = "" Then
    ccnt = outbxchk()
    If ccnt > 0 Then
      Command1.Enabled = False
      DoEvents
      Call Command16_Click
      Command1.Enabled = True
      DoEvents
    End If
  End If
Else
  tr = Dir(Form1.allusershome + "*.*", vbDirectory)
  List3.Clear
  While tr <> ""
    If tr <> "." And tr <> ".." Then
      Form1.myoutbox = Form1.allusershome + tr + "\mail\outbox"
      If nexist(Form1.myoutbox + "\lock.lck") Then
        List3.AddItem Form1.myoutbox
      Else
        Call Form1.dbg2f(Form1.myoutbox + " is locked")
      End If
    End If
    tr = Dir()
  Wend
  Command1.Enabled = False
  DoEvents
  While List3.ListCount > 0
    Form1.myoutbox = List3.List(0)
    Call Command16_Click
    List3.RemoveItem 0
  Wend
  Command1.Enabled = True
  DoEvents
End If
If gcount% = 12 Then
  gcount% = 0
  Command1.Enabled = False
  DoEvents
  Call Command3_Click
  Command1.Enabled = True
End If
Timer2.Enabled = True

End Sub

Private Sub tout_Change()
'd2infile = "popmain": d2insub = "tout_Change"
tot = Val(tout.Text)
End Sub

Sub rlist1()
Dim intNum As Integer, rrr
Dim intMessageNum As Integer 'the number of messages
Dim ap$, u$, pw$, aKey() As Byte, C$, o%, l$
Dim rid$
Dim rserver$
Dim ruser$
Dim rpsswd$
Dim rport$

If pin.Visible = True And trm(pin.Text) = "" Then
  If Form1.pin.Visible = True And trm(Form1.pin.Text) <> "" Then
    pin.Text = Form1.pin.Text
  End If
End If
MousePointer = 11
DoEvents
List1.Clear
u$ = Form1.getuserid()
ap$ = trm(pin.Text)
If List1.Visible = True And ap$ = "" Then List1.AddItem transe("PIN fehlt")
If poplistok% = 1 And List1.Visible = True And ap$ <> "" Then
  ap$ = trm(pin.Text)
  o% = FreeFile
  If Not nexist(Form1.hpth$ + "\poplist.agp") Then
  Open Form1.hpth$ + "\poplist.agp" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    rid$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    rserver$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    ruser$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    rpsswd$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    rport$ = l$
    If rid$ <> "PDFServer" Then
    pw$ = decrypt(rpsswd$, ap$)
    If (PopConnect(rserver$, Val(rport$), tot) = POP_SOCKET_ERROR) Then
        DoEvents
        List1.AddItem Mid(rid$, Len(u$) + 2) & ": Fehler"
        GoTo rrrout
    End If
    If (PopLogin(ruser$, pw$) = POP_SOCKET_ERROR) Then
        DoEvents
        List1.AddItem Mid(rid$, Len(u$) + 2) & ": Fehler"
        GoTo rrrout
    End If

    If (PopGetMessageCount(intMessageNum) = POP_SOCKET_ERROR) Then
        DoEvents
        List1.AddItem Mid(rid$, Len(u$) + 2) & ": Fehler"
        GoTo rrrout
    End If
    List1.AddItem intMessageNum & " auf " & Mid(rid$, Len(u$) + 2)
    End If
rrrout:
    Socket1.Disconnect
    DoEvents
  Wend
  Close #o%
  End If
End If
MousePointer = 0
DoEvents

End Sub

Public Sub popfire(r As Long, t As Long)
Dim l$, p%, dr As Double

'd2infile = "popmain": d2insub = "popfire"
dr = r
l$ = lblStatus.Caption
p% = InStr(l$, vbCrLf)
If p% > 0 Then l$ = trm(Left(l$, p% - 1))
If t = 0 Then Exit Sub
If Int(dr * 20 / t) <> prevfire% Then
  prevfire% = Int(dr * 20 / t)
  lblStatus.Caption = l$ & vbCrLf & trm(dr) & "/" & trm(t) & " (" & trm(Int(dr * 100 / t)) & "%)"
  pgb2.value = imin(Int(dr * 100 / t), 100)
  DoEvents
End If
Call l7set
End Sub

Private Sub darec_Click()

Call Form1.setmylastFormVar(Me.Name, "delrecvd", trm(darec.value))

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = x
       Else
        msg = x / Screen.TwipsPerPixelX
       End If
       Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopupSys
       End Select
End Sub

Private Sub Form_Resize()
       'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub mPopExit_Click()
       'called when user clicks the popup menu Exit command
       Unload Me
End Sub

Private Sub mPopRestore_Click()
       'called when the user clicks the popup menu Restore command
       Dim Result As Long
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
End Sub

Private Function mailthere() As Boolean
Dim tr

mailthere = False
tr = Dir(Form1.myinbox + "\*.amf")
If tr <> "" Then mailthere = True

End Function

Private Sub txtPassword_DblClick()
 If pin.Text = "0815" Then MsgBox txtPassword.Text
End Sub

Public Sub add_boxread(diff As Long)
boxread = boxread + diff
End Sub

Public Sub add_diskwrite(diff As Long)
diskwrite = diskwrite + diff
End Sub

Sub l7set()
Dim s As Long, t As Long, rc As String
s = logintimer + boxread + diskwrite
If s > 0 Then
  t = 100 * logintimer / s: rc = cut_d1(cut_d1(trm(t), ","), ".") + "/"
  t = 100 * boxread / s: rc = rc + cut_d1(cut_d1(trm(t), ","), ".") + "/"
  t = 100 * diskwrite / s: rc = rc + cut_d1(cut_d1(trm(t), ","), ".")
  Label7.Caption = "connect/box/file" + vbCrLf + rc
End If
End Sub
