VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Begin VB.Form Form1 
   Caption         =   "Agencyprof - POPClient"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3750
   Icon            =   "popForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows-Standard
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   1920
      Top             =   120
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
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      Picture         =   "popForm1.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Ihr Dokumentenverzeichnis öffnen"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox pin 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "popForm1.frx":0A6C
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Auf Wiedersehen!"
      Top             =   480
      Width           =   375
   End
   Begin VB.Label pinlbl 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Aktuelles Datum mit Uhrzeit"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public smtpUser As String, smtpPassword As String, iamserver As Boolean
Private Declare Function GetWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Dim userid$, s0d$, doc0dir$, dbname$, uId$, enc$, dbg2file%, localdir
Public hpth$, popentries As Integer, delrecvd As String, poplock
Public myinbox As String, myoutbox As String, setfn$, mymailserver As String
Public allusershome$

Public Sub Command1_Click()
Unload Me
End Sub

Private Sub Command19_Click()
Dim x

x = Shell("explorer.exe " & myinbox, vbNormalFocus)

End Sub

Private Sub Form_Load()
Dim o%, l$, rrr, s1d$

hpth$ = Environ$("HOMEDRIVE") + Environ$("HOMEPATH")
setfn$ = hpth$ + "\settings.agp"
popentries = 1
dbg2file% = 1
iamserver = False
enc$ = "hihallohuhu4716"
If TitleCounter(Me.Caption) = 2 Then
  End
End If
o% = FreeFile
On Error Resume Next
If nexist(setfn$) Then
  setfn$ = hpth$ + "\apserversettings.agp"
  If nexist(setfn$) Then setfn$ = "c:\Agencyprof\apserversettings.agp"
  iamserver = True
End If
Open setfn$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Benutzerdaten können nicht gelesen werden." + vbCrLf + "Datei: " + setfn$ + " (und settings.agp) nicht gefunden." + vbCrLf + "Starten Sie Agencyprof.", vbCritical, "AgencyprofPOPClient"
  End
Else
  Close #o%
End If
Command19.ToolTipText = "Inbox öffnen"
uId$ = Form1.getusersetting("userid", "")
If uId$ = "" Then
  MsgBox "No user-id given.", vbCritical, "AgencyprofPOPClient"
  End
End If
s0d$ = Form1.getusersetting("agencyprof", CurDir)
localdir = s0d$
localdir = getusersetting("localdir", trm(localdir))
allusershome$ = localdir + "\" + Form1.getusersetting("userdocdir", "") + "\"
doc0dir$ = localdir + "\" + Form1.getusersetting("userdocdir", "") + "\" + uId$
On Error Resume Next
Kill s0d$ & "\debug2file_" & uId$ & "_pop.txt"
Kill doc0dir + "\_mailausgang.log"
MkDir doc0dir + "\mail"
MkDir doc0dir + "\mail\inbox"
MkDir doc0dir + "\mail\outbox"
MkDir mylocaldatadir() + "\mail"
MkDir mylocaldatadir() + "\mail\inbox"
MkDir mylocaldatadir() + "\mail\outbox"
On Error GoTo 0

myinbox = doc0dir + "\mail\inbox"
myoutbox = doc0dir + "\mail\outbox"
Me.Top = Form1.mylasttop(Me.Name)
Me.Left = Form1.mylastleft(Me.Name)
Load popmain
popmain.txtUserName = getusersetting("popuser", uId$)
popmain.txtServer = getusersetting("popserver", "")
mymailserver = getusersetting("Mailserver", "")
'If popmain.txtServer = "" Then
'  End
'End If
smtpUser = getusersetting("smtpauth_user", "")
smtpPassword = getusersetting("smtpauth_password", "")
If smtpPassword <> "" Then smtpPassword = decrypt(Mid(smtpPassword, 9), "kzJfuz5vFRiuZ9oui974kJHbkGf")
popmain.txtPassword = decrypt(getusersetting("poppsswd", ""), enc$)
'popmain.txtPassword = decrypt(getusersetting("poppsswd", ""), "kzJfuz5vFRiuZ9oui974kJHbkGf")
popmain.txtPort = getusersetting("popport", "110")
Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form1.setmylasttop(Me.Name, Me.Top)
Call Form1.setmylastleft(Me.Name, Me.Left)
Unload popmain

End Sub

Public Function getuserid() As String
getuserid = uId$
End Function

Public Function getusersetting(fldn$, Optional vifnull As String) As String
Dim o%, l$, vin As String, fld$, rrr

If nexist(setfn$) Then
  End
End If
If vifnull <> "" Then vin = vifnull
fld$ = LCase(fldn$)
getusersetting = vin
o% = FreeFile
On Error Resume Next
Open setfn$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Call Form1.dbg2f("getusersetting: cannot open '" + setfn$ + "'(" + trm(rrr) + ")")
End If
While Not EOF(o%)
  Line Input #o%, l$
  If LCase(cut_d1(l$, "=")) = fld$ Then
    getusersetting = cut_d2bis(l$, "=")
    GoTo xfst1
  End If
Wend
xfst1:
Close #o%
End Function

Public Function getmyoutlook()
Dim uoutlk$

uoutlk$ = Form1.getusersetting("outlook", "")
If exist(uoutlk$) = 0 Then
  uoutlk$ = ""
End If
getmyoutlook = uoutlk$

End Function

Public Function inmylanguage(txt$) As String
Dim ltxt$, i%

inmylanguage = txt$
'ltxt$ = LCase(txt$)
'For i% = 0 To ttabptr% - 1
'  If LCase(transtab(0, i%)) = ltxt$ Then
'    inmylanguage = transtab(1, i%)
'    Exit For
'  End If
'Next i%

End Function

Public Function outmylanguage(txt$) As String
Dim ltxt$, i%

outmylanguage = txt$
'ltxt$ = LCase(txt$)
'For i% = 0 To ttabptr% - 1
'  If LCase(transtab(1, i%)) = ltxt$ Then
'    outmylanguage = transtab(0, i%)
'    Exit For
'  End If
'Next i%

End Function

Public Function s0dir() As String

s0dir = s0d$

End Function

Public Function mylocaldatadir() As String

'd2infile = "Form1": d2insub = "mydatadir"
mylocaldatadir = docs()

On Error Resume Next
MkDir docs()
MkDir mylocaldatadir
On Error GoTo 0

End Function

Public Function mylasttop(f$)
Dim inifile As String, l$, o%, rrr

l$ = "20"
inifile = Form1.mydatadir() + "\positions\" + f$ + ".top"
If exist(inifile) = 1 Then
  o% = FreeFile
  On Error Resume Next
  Open inifile For Input As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Function
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylasttop = CInt(l$)

End Function
Public Function mylastleft(f$)
Dim inifile As String, l$, o%, rrr

l$ = "20"
inifile = Form1.mydatadir() + "\positions\" + f$ + ".lft"
If exist(inifile) = 1 Then
  o% = FreeFile
  On Error Resume Next
  Open inifile For Input As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Function
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylastleft = CInt(l$)

End Function
Public Sub setmylastleft(f$, wert%)
Dim inifile As String, o%, rrr

inifile = Form1.mydatadir() + "\positions\" + f$ + ".lft"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

Print #o%, wert%
Close #o%

End Sub

Public Sub setmylasttop(f$, wert%)
Dim inifile As String, o%, rrr

'd2infile = "Form1": d2insub = "setmylasttop"
inifile = Form1.mydatadir() + "\positions\" + f$ + ".top"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, wert%
  Close #o%
End If
End Sub

Public Function mydatadir() As String

mydatadir = mylocaldatadir()

On Error Resume Next
MkDir mydatadir
On Error GoTo 0

End Function

Public Function docs() As String

'd2infile = "Form1": d2insub = "docs"
docs = doc0dir$

End Function

Private Sub pin_Change()
popmain.pin.Text = pin.Text
End Sub

Public Function mylastFormVar(frm$, var$, def$) As String
Dim inifile As String, l$, o%, rrr

l$ = def$
inifile = Form1.mydatadir() + "\positions\" + frm$ + "." & var$
If exist(inifile) = 1 Then
  o% = FreeFile
  On Error Resume Next
  Open inifile For Input As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Function
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
On Error Resume Next
mylastFormVar = l$
On Error GoTo 0

End Function

Public Sub setmylastFormVar(f$, v$, wert$)
Dim inifile As String, o%, rrr


'd2infile = "Form1": d2insub = "setmylastFormVar"
inifile = Form1.mydatadir() + "\positions\" + f$ + "." & v$
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, wert$
  Close #o%
End If
End Sub

Private Function TitleCounter(ttext As String) As Integer
Dim Length As Long
Dim sTitel As String
Dim CurHwnd As Long

TitleCounter = 0
CurHwnd = GetWindow(hwnd, GW_HWNDFIRST)
Do While CurHwnd <> 0
  ' Fenstertitel ermitteln
  sTitel = Space$(255)
  Length = GetWindowText(CurHwnd, sTitel, Len(sTitel))
  sTitel = Left$(sTitel, Length)

  ' Fenstertitel prüfen
  If InStr(sTitel, ttext) > 0 Then
    TitleCounter = TitleCounter + 1
  End If

  ' Handle des nächsten Fensters
  ' 0, wenn kein weiteres Fenster vorhanden
  CurHwnd = GetWindow(CurHwnd, GW_HWNDNEXT)
Loop
End Function

Private Function SmtpGetResultCodeSocket1(strResultString As String) As Integer

    Dim n As Integer
    Dim strResultCode As String
    Dim intBuffer As Integer
    Dim strBuffer As String
    Dim intBufferLength As Integer
    Dim bMultiLine As Boolean

'd2infile = "smtp": d2insub = "SmtpGetResultCodeSocket1"
    ' Read the result string sent back to the client after a
    ' command has been issued. The string has two parts, a numeric
    ' code which indicates success or failure, and a description
    ' of the result. In some cases, this may contain requested data
    ' or it may contain additional information about the result code
    ' (such as a description of why an error was returned).

    Do
        'Read until the end of the line

        strBuffer = String(1, 0)
        strResultCode = ""
        For n = 0 To 3
        'Read the first four digits which is the result code
            intBuffer = Socket1.Read(strBuffer, 1)
            strResultCode = strResultCode + strBuffer
        Next
Debug.Print strResultCode
        'If the last digit is a - then it is a multiline response
        If (Mid(strResultCode, 4, 1) <> "-") Then
            bMultiLine = False
        Else
            bMultiLine = True
        End If

        strResultCode = Left(strResultCode, 3)
        strBuffer = String(1, 0)

        'The read method will read only one line at a time because the binary property is false
        intBuffer = Socket1.Read(strBuffer, BUFFERSIZE)
        strResultString = strResultString + strBuffer
Debug.Print strResultString
Call Form1.logerr(trm(strResultString))
       If (intBuffer < 1) Then
            SmtpGetResultCodeSocket1 = SMTP_ERROR
            Exit Function
        End If

        If (bMultiLine = False) Then
            'This is the end of the last line of the response
            SmtpGetResultCodeSocket1 = Val(strResultCode)
            Exit Function
        End If
    Loop
    SmtpGetResultCodeSocket1 = SMTP_ERROR 'Error if the above Do Loop is broken
End Function
Private Function SmtpConnectSocket1(strHostName As String, lPort As Long, lTimeout As Long) As Integer
'd2infile = "smtp": d2insub = "SmtpConnectSocket1"
    'Connect to the server and recieve the result code
    Dim intError As Integer
    Dim intResult As Integer
    Dim strResultString As String

    Socket1.HostName = strHostName
    Socket1.RemotePort = lPort
    Socket1.Timeout = lTimeout * 1000 'Turn milliseconds into seconds
    Socket1.BUFFERSIZE = 1024
    intError = Socket1.Connect

    If (intError <> 0) Then
        SmtpConnectSocket1 = SMTP_ERROR
        Exit Function
    End If

    intResult = SmtpGetResultCodeSocket1(strResultString)

    SmtpConnectSocket1 = intResult
End Function
Private Function SmtpHelloExSocket1(strDomain As String, bExtended As Boolean) As Integer
    Dim intResultCode As Integer
    Dim intPos As Integer

'd2infile = "smtp": d2insub = "SmtpHelloExSocket1"
    'StrDomain is your domain name.  If a null string was passed in
    'Socket Wrench will determine the domain name.
    If (strDomain = "") Then
        strDomain = Socket1.LocalName
        intPos = InStr(strDomain, ".")
        If intPos > 1 Then strDomain = Left(strDomain, intPos - 1)
    End If

    'If bExtended is true, then the server's extended options will be sent back in the Result String
    If (bExtended) Then
        intResultCode = SmtpCommandSocket1("EHLO", strDomain)
    Else
        intResultCode = SmtpCommandSocket1("HELO", strDomain)
    End If

    SmtpHelloExSocket1 = intResultCode
End Function
Private Function SmtpCommandSocket1(strCmd As String, strParam As String) As Integer
    Dim strCommand As String
    Dim intBuffer As Integer
    Dim strResultString As String

'd2infile = "smtp": d2insub = "SmtpCommandSocket1"
    ' Send a command to the server and return the result code
    ' along with the string which describes the result (this
    ' can be particularly useful with errors)

    ' All commands must be terminated with a carriage-return
    ' linfeed sequence
    strCommand = strCmd & " " & strParam & Chr(13) & Chr(10)
    intBuffer = Socket1.Write(strCommand, Len(strCommand))
    If (intBuffer < 1) Then
        SmtpCommandSocket1 = SMTP_ERROR
        Exit Function
    End If

    ' Get the result code back from the server that
    ' indicates if the command was successful or not
    SmtpCommandSocket1 = SmtpGetResultCodeSocket1(strResultString)

End Function

Private Function SmtpBeginMessageSocket1(strFrom As String) As Integer
    Dim strParam As String
    Dim intResult As Integer

'd2infile = "smtp": d2insub = "SmtpBeginMessageSocket1"
    'Reset the server to start a new message
    intResult = SmtpCommandSocket1("RSET", "")

    If (intResult = SMTP_ERROR) Then
        SmtpBeginMessageSocket1 = SMTP_ERROR
        Exit Function
    End If

    'Tell the server who the message is from
    strParam = "FROM: <" & strFrom & ">"
    intResult = SmtpCommandSocket1("MAIL", strParam)
    SmtpBeginMessageSocket1 = intResult

End Function
Private Function SmtpAddRecipientSocket1(strAddress As String) As Integer
    Dim strParam As String
    Dim intResult As Integer

'd2infile = "smtp": d2insub = "SmtpAddRecipientSocket1"
    'Tell the server who to send the messaage to
    strParam = "TO:" & strAddress
    intResult = SmtpCommandSocket1("RCPT", strParam)
    SmtpAddRecipientSocket1 = intResult

End Function
Private Function SmtpAddCCRecipientSocket1(strAddress As String) As Integer
    Dim strParam As String
    Dim intResult As Integer

'd2infile = "smtp": d2insub = "SmtpAddCCRecipientSocket1"
    'Tell the server who to send the messaage to
    strParam = "cc:" & strAddress
    intResult = SmtpCommandSocket1("RCPT", strParam)
    SmtpAddCCRecipientSocket1 = intResult

End Function

Private Function SmtpEndMessageSocket1() As Integer
    Dim intBuffer As Integer
    Dim strResultString As String

'd2infile = "smtp": d2insub = "SmtpEndMessageSocket1"
   'Use the <crlf>.<crlf> format to tell the server this is the end of your message
    intBuffer = SmtpWriteSocket1(Chr(13) & Chr(10) & "." & Chr(13) & Chr(10), 5)
    bFirstWrite = True
    If (intBuffer < 1) Then
        SmtpEndMessageSocket1 = SMTP_ERROR
        Exit Function
    End If

    intBuffer = SmtpGetResultCodeSocket1(strResultString)

    SmtpEndMessageSocket1 = intBuffer
End Function

Private Function SmtpWriteSocket1(strBuffer As String, intLength As String) As Integer

    Dim intBuffer As Integer

'd2infile = "smtp": d2insub = "SmtpWriteSocket1"
    'If this is the first part of the message, use the DATA command to let the server know
    If (bFirstWrite) Then
        intBuffer = SmtpCommandSocket1("DATA", "")
        bFirstWrite = False
    End If

    'Write strBuffer to the server
    intBuffer = Socket1.Write(strBuffer, intLength)
    SmtpWriteSocket1 = intBuffer
End Function

Private Function SmtpDisconnectSocket1()
    Dim intBuffer As Integer, rc$

'd2infile = "smtp": d2insub = "SmtpDisconnectSocket1"
    'Tell the server you are quiting and disconnect
    intBuffer = SmtpCommandSocket1("QUIT", "")
    bFirstWrite = True
    Socket1.Disconnect
    SmtpDisconnectSocket1 = intBuffer
Call logerr("Quit: " + trm(SmtpGetResultCodeSocket1(rc$)))
End Function

Public Function mailresend(mailfile$, noti As Integer) As Boolean
Dim server$, from$, infn%, fromdom$, n%, port$, an$, l$
Dim strBuffer As String, cc$
Dim intPos As Integer, ifn$, a1n$, p%, a2n$, hdrend As Boolean
Dim strMess As String, i%, bndry$, o%
Dim wrb As Long, rrr
'd2infile = "Form1": d2insub = "mailresend"
mailresend = False
server$ = getusersetting("mailserver")
port$ = "25"
from$ = ""
cc$ = ""
fromdom$ = ""
Debug.Print "sending " + mailfile$ + " via " + server$
Call logerr("start sending " + mailfile$ + " via " + server$)
infn% = FreeFile
Open mailfile$ For Input As #infn%
hdrend = False
While Not EOF(infn%) And Not hdrend
  Line Input #infn%, l$
Debug.Print l$
  If l$ = "" Or InStr(LCase(l$), "content-type: ") > 0 Then
    hdrend = True
  Else
    If from$ = "" Then
      n% = InStr(LCase(l$), "from: ")
      If n% = 1 Then
        from$ = emailonly(trm(Mid$(l$, n% + 6)))
        fromdom$ = domainofemail(emailonly(strrepl(from, """", "")))
      End If
    End If
    n% = InStr(LCase(l$), "to: ")
    If n% = 1 Then
      an$ = emailonly(trm(Mid$(l$, n% + 4)))
    End If
    n% = InStr(LCase(l$), "cc: ")
    If n% = 1 Then
      cc$ = trm(Mid$(l$, n% + 4))
    End If
  End If
Wend
Close #infn%
If server$ <> "" And fromdom$ <> "" And an$ <> "" And from$ <> "" And mailfile$ <> "" Then
Call logerr("sending " + mailfile$)
  mailresend = plainsend(server$, fromdom$, an$, from$, mailfile$, noti)
Call logerr("testing CC")
  While cc$ <> ""
    an$ = cut_d1(cc$, ",")
    cc$ = cut_d2bis(cc$, ",")
    an$ = emailonly(trm(an$))
    mailresend = plainsend(server$, fromdom$, an$, from$, mailfile$, noti)
  Wend
Call logerr("testing BCC: " + mailfile$ + ".bcc")
  If Not nexist(mailfile$ + ".bcc") Then
    o% = FreeFile
    Open mailfile$ + ".bcc" For Input As #o%
    While Not EOF(o%)
      Line Input #o%, cc$
      While cc$ <> ""
        an$ = cut_d1(cc$, ",")
        cc$ = cut_d2bis(cc$, ",")
        an$ = emailonly(trm(an$))
Call logerr("sending BCC to " + an$)
        mailresend = plainsend(server$, fromdom$, an$, from$, mailfile$, noti)
      Wend
    Wend
    Close #o%
    rrr = 0
    On Error Resume Next
    Kill mailfile$ + ".bcc"
    rrr = Err
    On Error GoTo 0
Call logerr("killing BCC-file " + mailfile$ + ".bcc=" + trm(rrr))
  Else
    Call logerr("no BCC found")
  End If
End If
Call logerr("ended sending " + mailfile$ + " via " + server$)
End Function

Public Function plainsend(server$, helo$, an$, from$, file$, notify As Integer) As Boolean
Dim strBuffer As String, erg As Integer, quittung As String, sbuff As String
Dim intPos As Integer, ifn$, a1n$, p%, a2n$, optfile As String
Dim strMess As String, i%, bndry$, o%, s2send, s2read As Long, l$
Dim rc$, wrseq As Long, rccode$

'd2infile = "smtp": d2insub = "plainsend"
    
Socket1.Disconnect
Socket1.AutoResolve = False
Socket1.Blocking = True
Socket1.Binary = False   'Read a line at a time
Socket1.Protocol = IPPROTO_IP
bFirstWrite = True
If notify = 0 Then
  quittung = "0"
Else
  quittung = "1"
End If
quittung = "0"
optfile = Left(file$, Len(file$) - 3) + "aof"
If Not nexist(optfile) Then
  o% = FreeFile
  Open optfile For Input As #o%
  Line Input #o%, l$
  If InStr(LCase(l$), "quittung") = 1 Then quittung = cut_d2bis(l$, "=")
  Close #o%
  On Error Resume Next
  Kill optfile
  On Error GoTo 0
End If
  plainsend = False
    'Connect to the SMTP server
    If (SmtpConnectSocket1(server$, 25, 30) = SMTP_ERROR) Then
        Call logerr(an$ + " Fehler beim Verbindungsaufbau mit: " & server$ & vbCrLf & "Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If

    'Use Hello command
    If (SmtpHelloExSocket1(helo$, True) = SMTP_ERROR) Then
        Call logerr(an$ + "Fehler bei ""Helo"": Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If

'login wenn nötig:
    If smtpUser <> "" Then
      strBuffer = "AUTH LOGIN" + vbCrLf
      Call logerr(strBuffer)
      erg = Socket1.Write(strBuffer, Len(strBuffer))
      If (erg = SMTP_ERROR) Then
        Call logerr(an$ + " Schreibfehler: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
      End If
      rccode$ = trm(SmtpGetResultCodeSocket1(rc$))
      Call logerr("Resultcode=" + rccode$)
      strBuffer = EncodeStr64(smtpUser) + vbCrLf
      Call logerr(strBuffer)
      erg = Socket1.Write(strBuffer, Len(strBuffer))
      If (erg = SMTP_ERROR) Then
        Call logerr(an$ + " Schreibfehler: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
      End If
      rccode$ = trm(SmtpGetResultCodeSocket1(rc$))
      Call logerr("Resultcode=" + rccode$)
      strBuffer = EncodeStr64(smtpPassword) + vbCrLf
      Call logerr("statt passwort ...")
      erg = Socket1.Write(strBuffer, Len(strBuffer))
      If (erg = SMTP_ERROR) Then
        Call logerr(an$ + " Schreibfehler: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
      End If
      rccode$ = trm(SmtpGetResultCodeSocket1(rc$))
      Call logerr("Resultcode=" + rccode$)
    End If
    If Left(rccode$, 1) = "5" Then
      Call logerr("Error during login")
      Call logerr("Resultcode=" + rccode$)
      Exit Function
    End If
'nun sollten wir (besser) engeloggt sein....
    wrseq = 0
    'Tell the server who you are
Call logerr("MAIL FROM: <" + from$ + ">")
    strBuffer = "MAIL FROM: <" + from$ + ">" + vbCrLf
    erg = Socket1.Write(strBuffer, Len(strBuffer))
    If (erg = SMTP_ERROR) Then
        Call logerr(an$ + " Schreibfehler: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If
    'Tell the server who you will be sending the message to
    an$ = "RCPT TO: <" + an$ + ">"
Call logerr(an$ + " delivery notification=" + trm(quittung))
    If quittung = "1" Then an$ = an$ + " NOTIFY=SUCCESS,FAILURE"
    an$ = an$ + vbCrLf
    erg = Socket1.Write(an$, Len(an$))
    If (erg = SMTP_ERROR) Then
        Call logerr(an$ + " Fehler beim Hinzufügen des Empfängers: " & an$ & vbCrLf & "Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If
Call logerr("sending data")
    'Write the message
    s2send = FileLen(file$)
Call logerr("sending " + trm(s2send) + " Bytes")
    o% = FreeFile
    Open file$ For Input As #o%
    popmain.pgb1.Max = s2send
    s2read = 0
    popmain.pgb1.Visible = True
    While Not EOF(o%)

    Line Input #o%, strMess
    s2read = s2read + Len(strMess)
    popmain.pgb1.value = s2read
    DoEvents
    Debug.Print strMess
    Do
        'Find each line and write it to the server with crlf at the end
        wrseq = wrseq + 1
        intPos = InStr(strMess, Chr(10))
        If intPos > 512 Or Len(strMess) > 512 Then
            intPos = 500
            strMess = Left(strMess, intPos) & Chr(13) & Chr(10) & Right(strMess, Len(strMess) - intPos)
        End If
        intPos = InStr(strMess, Chr(10))
        If (intPos <> 0) Then
            strBuffer = Left(strMess, intPos - 2) & Chr(13) & Chr(10)
            strMess = Right(strMess, Len(strMess) - intPos)
        Else
            strBuffer = strMess & Chr(13) & Chr(10)
        End If
        If (SmtpWriteSocket1(strBuffer, Len(strBuffer)) = SMTP_ERROR) Then
            Call logerr(an$ + " Schreibfehler in Sequenz " + trm(wrseq) + ": Error " & Socket1.LastError)
            Socket1.Disconnect
            Exit Function
        End If
        DoEvents
    Loop While (intPos <> 0)

    Wend
    Close #o%
Call logerr("end of data")
    popmain.pgb1.Visible = False
    'Tell the server it is the end of the message and disconnect
    If (SmtpEndMessageSocket1() = SMTP_ERROR) Then
        Call logerr("Schreibfehler am Nachrichtsende: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If
Call logerr("Ended Message: " + trm(SmtpGetResultCodeSocket1(rc$)))
    'Let the user know that the message has been sent
    plainsend = True

    If (SmtpDisconnectSocket1() = SMTP_ERROR) Then
        Call logerr(an$ + " Fehler beim Verbindungsabbau: Error " & Socket1.LastError)
        Socket1.Disconnect
        Exit Function
    End If
Socket1.Disconnect

End Function

Sub logerr(txt As String)
Dim o%, dtg As String, fn$, out$

dtg = datum2sql(Date) + " " + trm(Time)
fn$ = doc0dir + "\_mailausgang.log"
o% = FreeFile
Open fn$ For Append As #o%
out$ = strrepl(txt, vbCrLf, " ")
out$ = strrepl(out$, Chr$(13), " ")
out$ = strrepl(out$, Chr$(10), " ")
out$ = strrepl(out$, "  ", " ")
Debug.Print dtg + " " + out$
Print #o%, dtg + " " + out$
Close #o%
End Sub

Public Sub dbg2f(l$)
Dim o%, diffd As Double, rrr

o% = FreeFile
On Error Resume Next
Open s0d$ & "\debug2file_" & uId$ & "_pop.txt" For Append As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "cannot write to log:" + vbCrLf + l$
  Exit Sub
End If
Print #o%, trm(Date) + " " + trm(Time) & ": " & l$
Close #o%

End Sub

