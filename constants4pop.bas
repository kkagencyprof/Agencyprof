Attribute VB_Name = "Constants4pop"
Global wrkJet As Workspace
'Global sqla As Database, dbname$, dbpara$
'----------------------------------------------------------------------------
'Function to play a wav file
Public Const SND_SYNC = &H0
Declare Function sndPlaySound Lib "winmm.dll" _
          Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                 ByVal uFlags As Long) As Long
'----------------------------------------------------------------------------

Declare Function DecodeFileEx Lib "UUCODE32.DLL" Alias "DecodeFileExA" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Declare Function EncodeFileEx Lib "UUCODE32.DLL" Alias "EncodeFileExA" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Declare Function CompressFile Lib "UUCODE32.DLL" Alias "CompressFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
Declare Function ExpandFile Lib "UUCODE32.DLL" Alias "ExpandFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
'
' The following functions have been deprecated and should no longer
' be used in new applications; use the DecodeFileEx and EncodeFileEx
' functions instead.
'
Declare Function DecodeFile Lib "UUCODE32.DLL" Alias "DecodeFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
Declare Function EncodeFile Lib "UUCODE32.DLL" Alias "EncodeFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
Declare Function DecodeBase64File Lib "UUCODE32.DLL" Alias "DecodeBase64FileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
Declare Function EncodeBase64File Lib "UUCODE32.DLL" Alias "EncodeBase64FileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long

Global Const POP_SOCKET_ERROR = -1 'An error occurred with the socket
Global Const POP_RESULT_ERROR = 0 'The server returned an error as a result of a command
Global Const POP_SUCCESS = 1 'The server returned a success as a result of a command
Global Const BUFFERSIZE = 4096
Global bFirstWrite As Boolean
Global Const SMTP_ERROR = -1
'
' General constants used with most of the controls
'
Global Const INVALID_HANDLE = -1
Global Const CONTROL_ERRIGNORE = 0
Global Const CONTROL_ERRDISPLAY = 1

'
' SocketWrench Control Actions
'
Global Const SOCKET_OPEN = 1
Global Const SOCKET_CONNECT = 2
Global Const SOCKET_LISTEN = 3
Global Const SOCKET_ACCEPT = 4
Global Const SOCKET_CANCEL = 5
Global Const SOCKET_FLUSH = 6
Global Const SOCKET_CLOSE = 7
Global Const SOCKET_DISCONNECT = 7
Global Const SOCKET_ABORT = 8
Global Const SOCKET_STARTUP = 9
Global Const SOCKET_CLEANUP = 10

'
' SocketWrench Control States
'
Global Const SOCKET_UNUSED = 0
Global Const SOCKET_IDLE = 1
Global Const SOCKET_LISTENING = 2
Global Const SOCKET_CONNECTING = 3
Global Const SOCKET_ACCEPTING = 4
Global Const SOCKET_RECEIVING = 5
Global Const SOCKET_SENDING = 6
Global Const SOCKET_CLOSING = 7

'
' Socket Address Families
'
Global Const AF_UNSPEC = 0
Global Const AF_UNIX = 1
Global Const AF_INET = 2

'
' Socket Types
'
Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2
Global Const SOCK_RAW = 3
Global Const SOCK_RDM = 4
Global Const SOCK_SEQPACKET = 5

'
' Protocol Types
'
Global Const IPPROTO_IP = 0
Global Const IPPROTO_ICMP = 1
Global Const IPPROTO_GGP = 2
Global Const IPPROTO_TCP = 6
Global Const IPPROTO_PUP = 12
Global Const IPPROTO_UDP = 17
Global Const IPPROTO_IDP = 22
Global Const IPPROTO_ND = 77
Global Const IPPROTO_RAW = 255
Global Const IPPROTO_MAX = 256

'
' Well-Known Port Numbers
'
Global Const IPPORT_ANY = 0
Global Const IPPORT_ECHO = 7
Global Const IPPORT_DISCARD = 9
Global Const IPPORT_SYSTAT = 11
Global Const IPPORT_DAYTIME = 13
Global Const IPPORT_NETSTAT = 15
Global Const IPPORT_CHARGEN = 19
Global Const IPPORT_FTP = 21
Global Const IPPORT_TELNET = 23
Global Const IPPORT_SMTP = 25
Global Const IPPORT_TIMESERVER = 37
Global Const IPPORT_NAMESERVER = 42
Global Const IPPORT_WHOIS = 43
Global Const IPPORT_MTP = 57
Global Const IPPORT_TFTP = 69
Global Const IPPORT_FINGER = 79
Global Const IPPORT_HTTP = 80
Global Const IPPORT_POP3 = 110
Global Const IPPORT_NNTP = 119
Global Const IPPORT_SNMP = 161
Global Const IPPORT_EXEC = 512
Global Const IPPORT_LOGIN = 513
Global Const IPPORT_SHELL = 514
Global Const IPPORT_RESERVED = 1024
Global Const IPPORT_USERRESERVED = 5000

'
' Network Addresses
'
Global Const INADDR_ANY = "0.0.0.0"
Global Const INADDR_LOOPBACK = "127.0.0.1"
Global Const INADDR_NONE = "255.255.255.255"

'
' Shutdown Values
'
Global Const SOCKET_READ = 0
Global Const SOCKET_WRITE = 1
Global Const SOCKET_READWRITE = 2

'
' Byte Order
'
Global Const LOCAL_BYTE_ORDER = 0
Global Const NETWORK_BYTE_ORDER = 1

'
' SocketWrench Error Response
'
Global Const SOCKET_ERRIGNORE = 0
Global Const SOCKET_ERRDISPLAY = 1

'
' SocketWrench Error Codes
'
Global Const WSABASEERR = 24000
Global Const WSAEINTR = 24004
Global Const WSAEBADF = 24009
Global Const WSAEACCES = 24013
Global Const WSAEFAULT = 24014
Global Const WSAEINVAL = 24022
Global Const WSAEMFILE = 24024
Global Const WSAEWOULDBLOCK = 24035
Global Const WSAEINPROGRESS = 24036
Global Const WSAEALREADY = 24037
Global Const WSAENOTSOCK = 24038
Global Const WSAEDESTADDRREQ = 24039
Global Const WSAEMSGSIZE = 24040
Global Const WSAEPROTOTYPE = 24041
Global Const WSAENOPROTOOPT = 24042
Global Const WSAEPROTONOSUPPORT = 24043
Global Const WSAESOCKTNOSUPPORT = 24044
Global Const WSAEOPNOTSUPP = 24045
Global Const WSAEPFNOSUPPORT = 24046
Global Const WSAEAFNOSUPPORT = 24047
Global Const WSAEADDRINUSE = 24048
Global Const WSAEADDRNOTAVAIL = 24049
Global Const WSAENETDOWN = 24050
Global Const WSAENETUNREACH = 24051
Global Const WSAENETRESET = 24052
Global Const WSAECONNABORTED = 24053
Global Const WSAECONNRESET = 24054
Global Const WSAENOBUFS = 24055
Global Const WSAEISCONN = 24056
Global Const WSAENOTCONN = 24057
Global Const WSAESHUTDOWN = 24058
Global Const WSAETOOMANYREFS = 24059
Global Const WSAETIMEDOUT = 24060
Global Const WSAECONNREFUSED = 24061
Global Const WSAELOOP = 24062
Global Const WSAENAMETOOLONG = 24063
Global Const WSAEHOSTDOWN = 24064
Global Const WSAEHOSTUNREACH = 24065
Global Const WSAENOTEMPTY = 24066
Global Const WSAEPROCLIM = 24067
Global Const WSAEUSERS = 24068
Global Const WSAEDQUOT = 24069
Global Const WSAESTALE = 24070
Global Const WSAEREMOTE = 24071
Global Const WSASYSNOTREADY = 24091
Global Const WSAVERNOTSUPPORTED = 24092
Global Const WSANOTINITIALISED = 24093
Global Const WSAHOST_NOT_FOUND = 25001
Global Const WSATRY_AGAIN = 25002
Global Const WSANO_RECOVERY = 25003
Global Const WSANO_DATA = 25004
Global Const WSANO_ADDRESS = 25004

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Sub GetLocalTime Lib "kernel32" (lpSystem As SYSTEMTIME)


Dim wheretolog$

Sub ShowError(ByVal strError As String)
    MsgBox strError, vbExclamation, "FTP"
End Sub

'POP**************************************************************

Public Function PopCommand(strCmd As String, strParam As String, strResultString) As Integer
    Dim strCommand As String
    Dim intBuffer As Integer
    Dim strResult As String, strp As String
    
    ' Send a command to the server and return the result code
    ' along with the string which describes the result (this
    ' can be particularly useful with errors)
    
    ' All commands must be terminated with a carriage-return
    ' linfeed sequence
    strp = strParam
    If strCmd = "PASS" Then strp = "xxxxxxxx"
    strCommand = strCmd & " " & strParam & Chr(13) & Chr(10)

    
    intBuffer = popmain.Socket1.Write(strCommand, Len(strCommand))
    If (intBuffer < 1) Then
        Call Form1.dbg2f("PopCommand Error: cmd=" + strCmd + " params=" + strp + " err=" + trm(Socket1.LastError))
        PopCommand = POP_SOCKET_ERROR
        Exit Function
    End If
    
    ' Get the result code back from the server that
    ' indicates if the command was successful or not
    Call Form1.dbg2f("PopCommand ok: cmd=" + strCmd + " params=" + strp)
    PopCommand = PopGetResultCode(strResult)
    'Sets strResultString to the string returned by the server
    strResultString = strResult
Call Form1.dbg2f(trm(strResultString))
End Function
Public Function PopGetResultCode(strResultString As String) As Integer

    Dim n As Integer
    Dim strResultCode As String
    Dim intBuffer As Integer
    Dim strBuffer As String
    Dim intBufferLength As Integer
    Dim bMultiLine As Boolean
    
    ' Read the result string sent back to the client after a
    ' command has been issued. The string has two parts, a
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
            intBuffer = popmain.Socket1.Read(strBuffer, 1)
            strResultCode = strResultCode + strBuffer
        Next
        
        
        'If the result code is +OK, then the command was a success and POP_SUCCESS should be returned
        If (strResultCode = "+OK ") Then
            PopGetResultCode = POP_SUCCESS
        Else
            'If the result code is -ERR, then there was an error and POP_RESULT_ERROR should be returned
            PopGetResultCode = POP_RESULT_ERROR
        End If
        
        
        Debug.Print "Result Code: " & strResultCode
        
        strBuffer = String(BUFFERSIZE, 0)
            
        'The read method will read only one line at a time because the binary property is false
        intBuffer = popmain.Socket1.Read(strBuffer, BUFFERSIZE)
        strResultString = strResultString + strBuffer
        If (intBuffer < 1) Then
            PopGetResultCode = POP_SOCKET_ERROR
            Exit Function
        End If
            
        If (bMultiLine = False) Then
            'This is the end of the last line of the response
            Debug.Print "Result String: " & strResultString
            Exit Function
        End If
    Loop
    PopGetResultCode = POP_SOCKET_ERROR 'Error if the above Do Loop is broken
End Function

Public Function PopConnect(strHostName As String, lPort As Long, lTimeout As Long) As Integer
    'Connect to the server and recieve the result code
    Dim intError As Integer
    Dim intResult As Integer
    Dim strResultString As String

    popmain.Socket1.HostName = strHostName
    popmain.Socket1.RemotePort = lPort
    popmain.Socket1.Timeout = lTimeout * 1000 'Change from milliseconds to seconds
    intError = popmain.Socket1.Connect
    
    If (intError <> 0) Then
        PopConnect = POP_SOCKET_ERROR
        Exit Function
    End If
    
    intResult = PopGetResultCode(strResultString)
    
    PopConnect = intResult
End Function
Public Function PopLogin(strUserName As String, strPassword As String) As Integer
    Dim intResultCode As Integer
    Dim strResultString As String
    'Logs into the server and returns the result code or POP_SOCKET_ERROR if there was an error
    intResultCode = PopCommand("USER", strUserName, strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopLogin = POP_SOCKET_ERROR
        Exit Function
    End If
    
    intResultCode = PopCommand("PASS", strPassword, strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopLogin = POP_SOCKET_ERROR
        Exit Function
    End If
    
    PopLogin = intResultCode
End Function
Public Function PopGetMessageCount(intMessageNum As Integer) As Integer
    'Stores the number of messages into intMessageNum using the STAT command
    'returns the result code or POP_SOCKET_ERROR if there is an error
    Dim intResultCode As Integer
    Dim lngResultCode As Long
    Dim strResultString As String
    Dim intPos As Integer
   
    
    intResultCode = PopCommand("STAT", "", strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopGetMessageCount = POP_SOCKET_ERROR
        Exit Function
    End If
    
    intPos = InStr(strResultString, " ")
    lngMessageNum = CLng(Left(strResultString, intPos - 1))
    If lngMessageNum > 5000 Then
      intMessageNum = 5000
    Else
      intMessageNum = Val(Left(strResultString, intPos - 1))
    End If
    PopGetMessageCount = intResultCode
Form1.dbg2f ("Messagecount: " + strResultString)
End Function
Public Function PopGetMessageHeader(intMessage As Integer, strHeader As String) As Integer
    'Stores the message header into strHeader using the TOP command
    'returns the result code or POP_SOCKET_ERROR if there is an error
    Dim strResultString As String
    Dim intResultCode As Integer
    Dim strBuffer As String
    Dim intBuffer As Integer
   
    intResultCode = PopCommand("TOP", trm(str(intMessage)) & " 0", strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopGetMessageHeader = POP_SOCKET_ERROR
        Exit Function
    End If
    
    strBuffer = String(BUFFERSIZE, 0)
    strHeader = ""
    Do
        intBuffer = popmain.Socket1.Read(strBuffer, BUFFERSIZE)
        If (intBuffer < 0) Then
            PopGetMessageHeader = POP_SOCKET_ERROR
            Exit Function
        End If
        
        strHeader = strHeader + strBuffer
    Loop While (Right(strHeader, 5) <> Chr(13) & Chr(10) & "." & Chr(13) & Chr(10))
    'The server has finished sending the header if it sends the sequence <CRLF>.<CRLF>
    
    PopGetMessageHeader = intResultCode
End Function

Public Function GetHeaderValue(ByVal strMessageHeader As String, ByVal strHeaderName As String) As String
    'Returns the header value specified by the HeaderName
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim intStart As Integer
    Dim mhead As String
    
    mhead = Left(strMessageHeader, 32768)
    intPos1 = InStr(mhead, strHeaderName & ":")
    If intPos1 = 0 Then intPos1 = InStr(LCase(mhead), LCase(strHeaderName) & ":")
    mhead = Right(mhead, (Len(mhead) - intPos1) + 1)
    
    intPos1 = InStr(mhead, " ")
    intStart = 1
    Do
        intPos2 = InStr(intStart, mhead, Chr(13) & Chr(10))
        intStart = intPos2 + 3
        ' If there is a tab after the <crlf>, the header is multi-line
        ' so find the end of the next line
    Loop While (Mid(mhead, intPos2 + 2, 1) = Chr(9))
    
    If intPos1 < intPos2 Then
        GetHeaderValue = Mid(mhead, intPos1 + 1, (intPos2 - intPos1) - 1)
    Else
        GetHeaderValue = "<" & mhead & " header value missing>"
    End If
    
End Function

Public Function PopDisconnect() As Integer
    'Logs off and disconnects from the server
    'returns the result code
    Dim strResultString As String
    Dim intResultCode As Integer
    
    intResultCode = PopCommand("QUIT", "", strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopDisconnect = POP_SOCKET_ERROR
        Exit Function
    End If
    
    If (popmain.Socket1.Disconnect() <> 0) Then
        PopDisconnect = POP_SOCKET_ERROR
        Exit Function
    End If
    
    PopDisconnect = intResultCode
    
End Function

Public Function PopDeleteMessage(intMessage As Integer)
    'Deletes the specified message and returns the result code
    Dim intResultCode As Integer
    Dim strResultString As String

    intResultCode = PopCommand("DELE", trm(str(intMessage)), strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopDeleteMessage = POP_SOCKET_ERROR
        Exit Function
    End If
    
    PopDeleteMessage = intResultCode
End Function

Public Function PopGetMessage(intMessage As Integer, strMessage As String)
    Dim intResultCode As Integer
    Dim strResultString As String
    Dim strBuffer As String
    Dim intBuffer As Integer
    
    'Read the selected message and store the data in strMessage
    
    intResultCode = PopCommand("RETR", str(intMessage), strResultString)
    If (intResultCode = POP_SOCKET_ERROR) Then
        PopGetMessage = POP_SOCKET_ERROR
        Exit Function
    End If
    
    strBuffer = String(BUFFERSIZE, 0)
    strMessage = ""
    Do
        intBuffer = popmain.Socket1.Read(strBuffer, BUFFERSIZE)
        'Debug.Print intBuffer
        If (intBuffer < 0) Then
            PopGetMessage = POP_SOCKET_ERROR
            Exit Function
        End If
    
        strMessage = strMessage + strBuffer
    Loop While (Right(strMessage, 5) <> Chr(13) & Chr(10) & "." & Chr(13) & Chr(10))
    'The server has finished sending the message if it sends the sequence <CRLF>.<CRLF> is sent
    
    PopGetMessage = intResultCode

End Function

Public Function PopGetLongMessage(intMessage As Integer, strMessageFile As String)
Dim o%
Dim intResultCode As Integer
Dim strResultString As String
Dim strBuffer As String
Dim intBuffer As Integer
Dim stm$, tlen As Long, rd As Long

'Read the selected message and store the data in strMessage
    
intResultCode = PopCommand("RETR", trm(intMessage), strResultString)
If (intResultCode = POP_SOCKET_ERROR) Then
  PopGetLongMessage = POP_SOCKET_ERROR
  Exit Function
End If
  
tlen = Val(word1(strResultString))
rd = 0
Call Form1.dbg2f("PopGetLongMessage getting to file " + strMessageFile)
Call tm_start(1)
o% = FreeFile
Open strMessageFile For Output As #o%
Call popmain.add_diskwrite(tm_stop(1))
strBuffer = String(BUFFERSIZE, 0)
Call Form1.dbg2f("starting do-loop ... " + trm(i))
Do
  Call tm_start(1)
  intBuffer = popmain.Socket1.Read(strBuffer, BUFFERSIZE)
  Call popmain.add_boxread(tm_stop(1))
  If (intBuffer < 0) Then
    PopGetLongMessage = POP_SOCKET_ERROR
    Close #o%
    Exit Function
  End If
  rd = rd + Len(strBuffer)
  Call popmain.popfire(rd, tlen)
  Call tm_start(1)
  Print #o%, strBuffer;
  Call popmain.add_diskwrite(tm_stop(1))
  stm$ = stm$ + strBuffer
  If Len(stm$) > 10 Then stm$ = Right$(stm$, 10)
Loop While (Right(stm$, 5) <> Chr(13) & Chr(10) & "." & Chr(13) & Chr(10))
Call Form1.dbg2f("loop finished " + trm(i))
'The server has finished sending the message if it sends the sequence <CRLF>.<CRLF>
Call tm_start(1)
Close #o%
Call popmain.add_diskwrite(tm_stop(1))
Call Form1.dbg2f("File closed" + trm(i))

PopGetLongMessage = intResultCode

End Function

Public Function SmtpCommand(strCmd As String, strParam As String) As Integer
    Dim strCommand As String
    Dim intBuffer As Integer
    Dim strResultString As String, strp As String
    
    ' Send a command to the server and return the result code
    ' along with the string which describes the result (this
    ' can be particularly useful with errors)
    
    ' All commands must be terminated with a carriage-return
    ' linfeed sequence
    strCommand = strCmd & " " & strParam & Chr(13) & Chr(10)
    strp = strParam
    If strCmd = "PASS" Then strp = "xxxxxxxx"
    Debug.Print "Command: " & strCmd & " " & strp
    intBuffer = popmain.Socket1.Write(strCommand, Len(strCommand))
    If (intBuffer < 1) Then
        SmtpCommand = SMTP_ERROR
        Exit Function
    End If
    
    ' Get the result code back from the server that
    ' indicates if the command was successful or not
    SmtpCommand = SmtpGetResultCode(strResultString)

End Function

Public Function SmtpGetResultCode(strResultString As String) As Integer

    Dim n As Integer
    Dim strResultCode As String
    Dim intBuffer As Integer
    Dim strBuffer As String
    Dim intBufferLength As Integer
    Dim bMultiLine As Boolean
    
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
            intBuffer = popmain.Socket1.Read(strBuffer, 1)
            strResultCode = strResultCode + strBuffer
        Next
        
        'If the last digit is a - then it is a multiline response
        If (Mid(strResultCode, 4, 1) <> "-") Then
            bMultiLine = False
        Else
            bMultiLine = True
        End If
    
        strResultCode = Left(strResultCode, 3)
        strBuffer = String(1, 0)
            
        'The read method will read only one line at a time because the binary property is false
        intBuffer = popmain.Socket1.Read(strBuffer, BUFFERSIZE)
        strResultString = strResultString + strBuffer
        If (intBuffer < 1) Then
            SmtpGetResultCode = SMTP_ERROR
            Exit Function
        End If
            
        If (bMultiLine = False) Then
            'This is the end of the last line of the response
            Debug.Print "Result String: " & strResultString
            SmtpGetResultCode = Val(strResultCode)
            Exit Function
        End If
    Loop
    SmtpGetResultCode = SMTP_ERROR 'Error if the above Do Loop is broken
End Function
Public Function SmtpConnect(strHostName As String, lPort As Long, lTimeout As Long) As Integer
    'Connect to the server and recieve the result code
    Dim intError As Integer
    Dim intResult As Integer
    Dim strResultString As String
    
    popmain.Socket1.HostName = strHostName
    popmain.Socket1.RemotePort = lPort
    popmain.Socket1.Timeout = lTimeout * 1000 'Turn milliseconds into seconds
    intError = popmain.Socket1.Connect
    
    If (intError <> 0) Then
        SmtpConnect = SMTP_ERROR
        Exit Function
    End If
    
    intResult = SmtpGetResultCode(strResultString)
    
    SmtpConnect = intResult
End Function

Public Function SmtpHelloEx(strDomain As String, bExtended As Boolean) As Integer
    Dim intResultCode As Integer
    Dim intPos As Integer
    
    'StrDomain is your domain name.  If a null string was passed in
    'Socket Wrench will determine the domain name.
    If (strDomain = "") Then
        strDomain = popmain.Socket1.LocalName
        intPos = InStr(strDomain, ".")
        strDomain = Left(strDomain, intPos - 1)
    End If
    
    'If bExtended is true, then the server's extended options will be sent back in the Result String
    If (bExtended) Then
        intResultCode = SmtpCommand("EHLO", strDomain)
    Else
        intResultCode = SmtpCommand("HELO", strDomain)
    End If

    SmtpHelloEx = intResultCode
End Function

Public Function SmtpBeginMessage(strFrom As String) As Integer
    Dim strParam As String
    Dim intResult As Integer
    
    'Reset the server to start a new message
    intResult = SmtpCommand("RSET", "")
    
    If (intResult = SMTP_ERROR) Then
        SmtpBeginMessage = SMTP_ERROR
        Exit Function
    End If
      
    'Tell the server who the message is from
    strParam = "FROM: <" & strFrom & ">"
    intResult = SmtpCommand("MAIL", strParam)
    SmtpBeginMessage = intResult

End Function

Public Function SmtpAddRecipient(strAddress As String) As Integer
    Dim strParam As String
    Dim intResult As Integer
    
    'Tell the server who to send the messaage to
    strParam = "TO: <" & strAddress & ">"
    intResult = SmtpCommand("RCPT", strParam)
    SmtpAddRecipient = intResult

End Function

Public Function SmtpWrite(strBuffer As String, intLength As String) As Integer

    Dim intBuffer As Integer
    
    'If this is the first part of the message, use the DATA command to let the server know
    If (bFirstWrite) Then
        intBuffer = SmtpCommand("DATA", "")
        bFirstWrite = False
    End If
    
    'Write strBuffer to the server
    intBuffer = popmain.Socket1.Write(strBuffer, intLength)
    
    SmtpWrite = intBuffer
End Function

Public Function SmtpEndMessage() As Integer
    Dim intBuffer As Integer
    Dim strResultString As String
    
   'Use the <crlf>.<crlf> format to tell the server this is the end of your message
    intBuffer = SmtpWrite(Chr(13) & Chr(10) & "." & Chr(13) & Chr(10), 5)
    bFirstWrite = True
    If (intBuffer < 1) Then
        SmtpEndMessage = SMTP_ERROR
        Exit Function
    End If
    
    intBuffer = SmtpGetResultCode(strResultString)
    
    SmtpEndMessage = intBuffer
End Function

Public Function SmtpDisconnect()
    Dim intBuffer As Integer
    
    'Tell the server you are quiting and disconnect
    intBuffer = SmtpCommand("QUIT", "")
    bFirstWrite = True
    popmain.Socket1.Disconnect
    
    SmtpDisconnect = intBuffer
End Function

Function GetFormattedTime() As String
    Dim Months(12) As String
    Dim Days(7) As String
    Dim SysTime As SYSTEMTIME
    
    Months(1) = "Jan"
    Months(2) = "Feb"
    Months(3) = "Mar"
    Months(4) = "Apr"
    Months(5) = "May"
    Months(6) = "Jun"
    Months(7) = "Jul"
    Months(8) = "Aug"
    Months(9) = "Sep"
    Months(10) = "Oct"
    Months(11) = "Nov"
    Months(12) = "Dec"
    
    Days(1) = "Sun"
    Days(2) = "Mon"
    Days(3) = "Tue"
    Days(4) = "Wed"
    Days(5) = "Thu"
    Days(6) = "Fri"
    Days(7) = "Sat"
    
    GetLocalTime SysTime
    GetFormattedTime = Days(SysTime.wDayOfWeek + 1) & ", "
    GetFormattedTime = GetFormattedTime & Format(SysTime.wDay, "0#") & " " & Months(SysTime.wMonth) & " " & SysTime.wYear & " "
    GetFormattedTime = GetFormattedTime & Format(SysTime.wHour, "0#") & ":" & Format(SysTime.wMinute, "0#") & ":" & Format(SysTime.wSecond, "0#")
    
End Function

