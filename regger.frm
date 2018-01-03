VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form regger 
   Caption         =   "Create RegistrationCode"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows-Standard
   Begin SocketWrenchCtrl.Socket sockServer 
      Index           =   0
      Left            =   7440
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
   Begin VB.TextBox cookie 
      Height          =   285
      Left            =   4680
      TabIndex        =   13
      Text            =   "cookie"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   0
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   315
      Left            =   7200
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   8295
   End
   Begin VB.TextBox editMaxClients 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox editLocalPort 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text9"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox editLocalAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text9"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&schliessen"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Height          =   1245
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "regger.frx":0000
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "regger.frx":000B
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "to save in"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Got these keys"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Send back this Key"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "regger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g_nMaxClients As Integer
Dim g_nLastClient As Integer, crlf$
Dim g_nActiveClients As Integer, x4$
Dim user$(999), wrtcnt%, birth As Variant, nowr As Integer, curri%, dbname$, dbkkname$
Dim wrkJet As Workspace
Dim sqla As Database, dbpara$
Dim sqlkk As Database, dbkkpara$
Dim aKey() As Byte, licfile$
Dim strCipher As String, udat$


Public Function strrepl(Text$, such$, ersetz$) As String
Dim t$, n$

t$ = Text$
n$ = ""
While InStr(t$, such$) > 0
  n$ = n$ + Left$(t$, InStr(t$, such$) - 1) + ersetz$
  t$ = Mid$(t$, InStr(t$, such$) + Len(such$))
Wend
If Len(t$) > 0 Then n$ = n$ + t$
strrepl = n$

End Function

Public Function mkkey(l%) As String
Dim i%, rc$, v%, K%

i% = l%
While i% > 0
  z$ = Chr$(Int(Rnd * 25 + 65))
  rc$ = rc$ + z$
  i% = i% - 1
Wend

mkkey = rc$

End Function

Public Function newid(t$, key$, l)
Dim id$, stmp As Recordset, le%
le% = l
Do
  id$ = mkkey(le%)
  cmd$ = "SELECT " + key$ + " FROM " + t$ + " where [" + key$ + "]='" + id$ + "'"
  rcnt% = 0
  Do
    On Error Resume Next
    Set stmp = sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
    rrr = Err
    On Error GoTo 0
  Loop Until rrr = 0 Or rcnt% > 2
  stw = stmp.EOF
  stmp.Close
Loop Until stw
newid = id$

End Function

Public Sub ExQDef(qdfTemp As QueryDef)

    Dim errLoop As Error, didr%
    
    List1.AddItem qdfTemp.SQL
    On Error GoTo Err_Execute
    qdfTemp.Execute dbFailOnError
    On Error GoTo 0

    Exit Sub

Err_Execute:

    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            List1.AddItem "Fehlernummer: " & errLoop.Number & vbCr & errLoop.Description
        Next errLoop
    End If
    
    Resume Next

End Sub

Public Sub sqlqry(stmt$)
Dim rtmp As QueryDef

On Error Resume Next
Set rtmp = sqla.CreateQueryDef("", stmt$)
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Call ExQDef(rtmp)
  rtmp.Close
Else
  List1.AddItem "Error in SQL-Satement:" + stmt$
'  MsgBox stmt$
End If

End Sub


Public Function datfromsql(dtg) As String

datfromsql = Right(dtg, 2) + "." + Mid(dtg, 6, 2) + "." + Left(dtg, 4)

End Function
Public Function datum2sql(dtg) As String


datum2sql = ""
If Len(dtg) > 0 Then
y$ = Year(dtg)
M$ = Format$(Month(dtg), "00")
d$ = Format$(Day(dtg), "00")
datum2sql = y$ + "-" + M$ + "-" + d$
End If

End Function

Public Function exist(fn$)
Dim o%, rrr

o% = FreeFile
On Error Resume Next
Open fn$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Close #o%
  exist = 1
Else
  exist = 0
End If

End Function

Sub sswr(i%, l$)
    Dim strBuffer As String
    Dim cchBuffer As Integer
    Dim nBytesWritten As Integer

        Call smsg("write(" & i% & "):" & l$)
        strBuffer = l$ + Chr$(13) + Chr$(10)
        cchBuffer = Len(strBuffer)
        nBytesWritten = sockServer(i%).Write(strBuffer, cchBuffer)
        If (nBytesWritten = -1) Then
            smsg "error during write.  Error " & sockServer(i%).LastError
            sockServer(i%).Disconnect
            g_nActiveClients = g_nActiveClients - 1
            
        End If

End Sub
Private Sub cmdDisconnect_Click()
    Dim nClient As Integer
    Dim nError As Integer


    For nClient = 1 To g_nLastClient
        If sockServer(nClient).Connected Then
            nError = sockServer(nClient).Disconnect
            If (nError <> 0) Then
                Call smsg("Unable to disconnect client " & nClient & ".  Error " & nError)
            Else
              Call smsg("disconnected client " & nClient)
            End If
            g_nActiveClients = g_nActiveClients - 1
        End If
    Next
Call smsg("disconnected")
End Sub
Sub smsg(l$)
List1.AddItem l$
List1.ListIndex = List1.ListCount - 1
DoEvents
List1.ListIndex = -1
While List1.ListCount > 1000
  List1.RemoveItem 0
Wend
End Sub
Private Sub cmdListen_Click()
    Dim nLocalPort As Long
    Dim nValue As Integer
    Dim nError As Integer

    '
    ' If the server is listening for connections, then
    ' effectively pause the server by closing the listening
    ' socket; no further connections will be accepted on
    ' the specified port, although existing client connections
    ' will be unaffected.
    '
    If sockServer(0).Listening Then
        nError = sockServer(0).Disconnect
        If (nError <> 0) Then
            MsgBox "Unable to disconnect.  Error " & nError
        End If
        Exit Sub
    End If

    '
    ' Update the value for the maximum number of client connections
    ' that will be accepted by the server
    '
    nValue = Val(Trim(editMaxClients.Text))
    If nValue < 0 Or nValue > 999 Then
        MsgBox "The specified maximum number of clients is invalid", vbExclamation, App.Title
        editMaxClients.SetFocus
        Exit Sub
    End If

    '
    ' Make sure that the maximum number of connections is not
    ' less than the number of clients already connected
    '
    If nValue >= g_nActiveClients Then
        g_nMaxClients = nValue
    Else
        g_nMaxClients = g_nActiveClients
        editMaxClients.Text = CStr(g_nMaxClients)
    End If

    '
    ' Valid port numbers are 1 through 65535, however, since
    ' the control's LocalPort property is an signed integer,
    ' we need to convert ports above 32767 to a negative value
    ' so that it doesn't overflow
    '
    nLocalPort = CLng(Val(Trim(editLocalPort.Text)))
    If nLocalPort > 0 And nLocalPort < 65536 Then
        If nLocalPort > 32767 Then nLocalPort = nLocalPort - 65536
    Else
        MsgBox "The specified port number is invalid", vbExclamation, App.Title
        editLocalPort.SetFocus
        Exit Sub
    End If
    
    sockServer(0).AutoResolve = False
    sockServer(0).BindAddress = Trim(editLocalAddress.Text)
    sockServer(0).Blocking = False
    sockServer(0).BUFFERSIZE = 4096
    sockServer(0).LocalPort = nLocalPort
    
    If sockServer(0).Listen() > 0 Then
        Select Case sockServer(0).LastError
        Case WSAEADDRINUSE:
            '
            ' This error occurs when another application is listening on
            ' the same local IP address and port number. If this error
            ' occurs on port 7 under Windows NT or Windows 2000, this
            ' typically indicates that "Simple TCP/IP Services" have been
            ' installed and are currently running on the local host.
            '
            MsgBox "Another application is already listening on this port", vbExclamation, App.Title

        Case Else
            MsgBox "An unexpected error occurred (Error " & sockServer(0).LastError & ")"
        End Select
    Else
      smsg "Listening ..."
    End If
    
End Sub

Private Sub Command1_Click()
Call cmdDisconnect_Click
Unload regger
End Sub

Private Sub editLocalAddress_Change()
If nowr = 1 Then Exit Sub

      o% = FreeFile
      Open "c:\myip.txt" For Output As #o%
      Print #o%, editLocalAddress.Text
      Close #o%
    

End Sub

Private Sub Form_Load()

curri% = -1
birth = Now
crlf$ = Chr$(13) + Chr$(10)
Randomize
    g_nMaxClients = 100
    g_nLastClient = 0
    g_nActiveClients = 0

    nowr = 1
    editLocalAddress.Text = "192.168.10.201"
    nowr = 0
    editLocalPort.Text = CStr(7)
    editMaxClients.Text = CStr(g_nMaxClients)
    If exist("c:\myip.txt") > 0 Then
      o% = FreeFile
      Open "c:\myip.txt" For Input As #o%
      Line Input #o%, l$
      Close #o%
      editLocalAddress.Text = l$
    End If
Show
x4$ = "Suak"

wrtcnt% = 5
Timer1.Interval = 1000
Timer1.Enabled = True

smsg "checking database"
Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
dbname$ = "ap-users"
dbpara$ = "ODBC;DATABASE=" & Trim(dbname$) & _
          ";UID=root" & _
          ";PWD=" & _
          ";DSN=" & dbname$
dbkkname$ = "sqlagent-kk"
dbkkpara$ = "ODBC;DATABASE=" & Trim(dbkkname$) & _
          ";UID=root" & _
          ";PWD=" & _
          ";DSN=" & dbkkname$
          
Set sqla = wrkJet.OpenDatabase(dbname$, dbDriverNoPrompt, False, dbpara$)

End Sub


Function dval(a$)
dval = 0
a$ = LCase(a$)
If a$ >= "0" And a$ <= "9" Then
  dval = Val(a$)
Else
  If a$ >= "a" And a$ <= "f" Then
    dval = (Asc(a$) - Asc("a")) + 10
  Else
    dval = 0
  End If
End If
Debug.Print a$, dval
End Function


Private Sub Text6_Change()
Dim l1$, l2$

l$ = Text6.Text
If Len(l$) = 0 Then Exit Sub

l3$ = ""
p% = InStr(l$, crlf$)
If p% > 0 Then
  l1$ = Left$(l$, p% - 1)
  l2$ = Mid$(l$, p% + 2)
  p% = InStr(l2$, crlf$)
  If p% > 0 Then
    l3$ = Mid$(l2$, p% + 2)
    l2$ = Left$(l2$, p% - 1)
  End If
Else
  l1$ = l$
  l2$ = ""
End If
smsg "decrypt:" + l1$
K$ = datum2sql(Date)
e$ = ""
For i% = Len(K$) To 1 Step -1
  e$ = e$ + Mid$(K$, i%, 1)
Next i%
smsg "key    :" + "Netsrak" + e$ + x4$
e$ = cv_HexFromBytes("Netsrak" + e$ + x4$)
aKey() = cv_BytesFromHex(e$)
Call blf_KeyInit(aKey)
g$ = blf_StringDec(cv_StringFromHex(l1$))
smsg "result :" + strrepl(g$, "  ", " ")
While isdigit(Left$(g$, 1)) = 0: g$ = Mid$(g$, 2): Wend
e$ = cv_HexFromString(g$)
aKey() = g$
'Debug.Print g$, aKey
Call blf_KeyInit(aKey)
smsg "decrypt:" + l2$
g$ = blf_StringDec(cv_StringFromHex(l2$))
smsg "result :" + strrepl(g$, "  ", " ")
licfile$ = g$
Text8.Text = g$
DoEvents
smsg "decrypt:" + l3$
g$ = blf_StringDec(cv_StringFromHex(l3$))
smsg "result :" + strrepl(g$, "  ", " ")
udat$ = g$
DoEvents

End Sub
Public Function isdigit(C$)

isdigit = InStr("1234567890", C$)

End Function

Private Sub sockServer_Accept(Index As Integer, SocketId As Integer)
    Dim nClient As Integer

    For nClient = 1 To g_nLastClient
        If sockServer(nClient).Connected = False Then
            sockServer(nClient).Accept = SocketId
            Exit Sub
        End If
    Next

    g_nLastClient = g_nLastClient + 1
    
    Load sockServer(g_nLastClient)
    sockServer(g_nLastClient).AutoResolve = False
    sockServer(g_nLastClient).Blocking = False
    sockServer(g_nLastClient).BUFFERSIZE = 4096
    sockServer(g_nLastClient).Accept = SocketId
    Exit Sub
End Sub
Private Sub sockServer_Connect(Index As Integer)
    '
    ' Check the number of active clients against the maximum number
    ' of clients specified by the user
    '
    If g_nActiveClients < g_nMaxClients Then
        g_nActiveClients = g_nActiveClients + 1
    Else
        smsg "Rejected connection from client at " & sockServer(Index).PeerAddress
        sockServer(Index).Disconnect
        Exit Sub
    End If
    
    smsg "Accepted connection from client " & Index & " at " & sockServer(Index).PeerAddress
    Call sswr(Index, "Hi there " & " at " & sockServer(Index).PeerAddress)
    
End Sub

Private Sub sockServer_Disconnect(Index As Integer)
    Dim nError As Integer
    
    nError = sockServer(Index).Disconnect
    If (nError <> 0) Then
        smsg "Unable to disconnect.  Error " & nError
    Else
        g_nActiveClients = g_nActiveClients - 1
    End If
    
    smsg "Client " & Index & " disconnected"
    
End Sub

Private Sub sockServer_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
    Dim strBuffer As String
    Dim cchBuffer As Integer
    Dim nBytesWritten As Integer

    cchBuffer = sockServer(Index).Read(strBuffer, 1024)
    If cchBuffer > 0 Then
      l$ = Left$(strBuffer, Len(strBuffer) - 2)
        Call smsg("got(" & Index & "):" & l$)
        nBytesWritten = sockServer(Index).Write(strBuffer, cchBuffer)
        If (nBytesWritten = -1) Then
            smsg "error during write.  Error " & sockServer(Index).LastError
            sockServer(Index).Disconnect
            g_nActiveClients = g_nActiveClients - 1
            
        End If
        Call interpreter(Index, l$)
    End If
    
    If cchBuffer = -1 Then
        smsg "An error occurred during read.  Error " & sockServer(Index).LastError
        sockServer(Index).Disconnect
        g_nActiveClients = g_nActiveClients - 1
        
    End If
    
End Sub



Private Sub Text8_Change()
Dim d0 As Variant

Text7.Text = ""
On Error GoTo rout
d0 = datfromsql(Left$(Text8.Text, 10))
d0 = CDate(d0) + 7
g$ = datum2sql(d0) + cookie.Text
strCipher = blf_StringEnc(g$)
g2$ = cv_HexFromString(strCipher)
Text7.Text = g2$

rout:
On Error GoTo 0

End Sub

Private Sub Timer1_Timer()
If wrtcnt > 0 Then
  wrtcnt% = wrtcnt% - 1
  smsg "" & wrtcnt%
Else
  Timer1.Enabled = False
  Call cmdListen_Click

End If
End Sub

Sub interpreter(i%, l$)
Dim r1 As Recordset, nid$, w As Recordset

p% = InStr(l$, " ")
If p% > 0 Then
  w1$ = Trim(Left$(l$, p% - 1))
  r$ = Mid$(l$, p% + 1)
Else
  w1$ = Trim(l$)
  r$ = ""
End If
Select Case w1$
 Case "date": Call sswr(i%, Date)
 Case "time": Call sswr(i%, Time)
 Case "cyc": o% = FreeFile: Open "lock.lck" For Output As #o%: Close #o%
              Call sswr(i%, "oki")
              Call cmdDisconnect_Click
              x = Shell("zlauncher.exe register.exe", 1)
              End
 Case "huhu": Call sswr(i%, "selba huhu")
 Case "status": Call sswr(i%, App.EXEName & " " & App.Major & " " & App.Minor & " " & App.Revision & " " & App.LegalCopyright)
                Call sswr(i%, "" & Now - birth)
 Case "lock": o% = FreeFile: Open "lock.lck" For Output As #o%: Close #o%
              Call sswr(i%, "oki")
 Case "schuess": Call cmdDisconnect_Click
                 x = Shell("zlauncher.exe register.exe", 1)
                 End
 Case "off": End
 Case "dumpreglog": cmd$ = "select * from reglog order by dtg"
                    Set r1 = sqla.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
                    While Not r1.EOF
                      Call sswr(i%, "" & r1!dtg & ":" & r1!remoteip & ":" & r1!licfile)
                      r1.MoveNext
                    Wend
 Case "req":
              If curri% = -1 Then
                curri% = i%
                nid$ = newid("reglog", "id", 28)
                cookie.Text = nid$
                Text6.Text = strrepl(r$, "|", crlf$)
                DoEvents
                If Text7.Text <> "" Then
                  Call sswr(curri%, "Welcome " + Text7.Text)
                  cmd$ = "insert into reglog (id,remoteip,licfile,incode1,incode2,dtg,udata) values('" & _
                    nid$ & "','" & _
                    sockServer(i%).PeerAddress & "','" & _
                    licfile$ & "','" & _
                    Left$(r$, InStr(r$, "|") - 1) & "','" & _
                    Mid$(r$, InStr(r$, "|") + 1) & "','" & _
                    datum2sql(Date) & "-" & Time & "','" & _
                    udat$ & "')"
                  Call sqlqry(cmd$)
                End If
                Text6.Text = ""
                curri% = -1
              Else
                Call sswr(i%, "STANDBY")
              End If
 Case "kurse":
            hisbis$ = r$
            Set sqlkk = wrkJet.OpenDatabase(dbkkname$, dbDriverNoPrompt, False, dbkkpara$)
            cmd$ = "select top 1 id from kurse order by id desc"
            Set w = sqlkk.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
            If Not w.EOF Then
                Call sswr(i%, "Kurse " & Left(w!id, 10))
            End If
            If hisbis$ <> "" Then
              cmd$ = "select * from kurse where [id]>'" + hisbis$ + "'"
              Set w = sqlkk.OpenRecordset(cmd$, dbOpenDynaset, dbReadOnly)
              Call sswr(i%, "sende kurse")
              While Not w.EOF
                Call sswr(i%, "'" & w!id & "','" & w!wid & "','" & w!kurs & "','" & w!einheit & "'")
                w.MoveNext
              Wend
              ' no error, just fertig
              Call sswr(i%, "byebye")
              d0 = Time: While d0 = Time: DoEvents: Wend
              Call sswr(i%, "byebye")
              d0 = Time: While d0 = Time: DoEvents: Wend
              Call sswr(i%, "byebye")
              d0 = Time: While d0 = Time: DoEvents: Wend
              Call sswr(i%, "byebye")
              d0 = Time: While d0 = Time: DoEvents: Wend
              sockServer(i%).Disconnect
              Exit Sub
            End If
            Call sswr(i%, "ende kurse")
            sqlkk.Close
 Case Default:
End Select
End Sub
