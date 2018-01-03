Attribute VB_Name = "Constants"
Global wrkJet As Workspace
'Global sqla As Database, dbname$, dbpara$
Global ftp_bytes_sent_this_file As Long
Global ftp_bytes_got_this_file As Long
'----------------------------------------------------------------------------
'Function to play a wav file
Public Const SND_SYNC = &H0
Declare Function sndPlaySound Lib "winmm.dll" _
          Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                 ByVal uFlags As Long) As Long
'----------------------------------------------------------------------------

'Declare Function DecodeFileEx Lib "UUCODE32.DLL" Alias "DecodeFileExA" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'Declare Function EncodeFileEx Lib "UUCODE32.DLL" Alias "EncodeFileExA" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'Declare Function CompressFile Lib "UUCODE32.DLL" Alias "CompressFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
'Declare Function ExpandFile Lib "UUCODE32.DLL" Alias "ExpandFileA" (ByVal strInputFile As String, ByVal strOutputFile As String) As Long
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

Global Const FTP_REPLY_RESTMARK = 110     ' Restart marker reply
Global Const FTP_REPLY_NOTREADY = 120     ' Service available in n minutes
Global Const FTP_REPLY_DATAOPEN = 125     ' Data connection open, transfer started
Global Const FTP_REPLY_FILEOK = 150       ' File status okay
Global Const FTP_REPLY_CMDOK = 200        ' Command okay
Global Const FTP_REPLY_IGNORED = 202      ' Command ignored
Global Const FTP_REPLY_SYSSTAT = 211      ' System status
Global Const FTP_REPLY_DIRSTAT = 212      ' Directory status
Global Const FTP_REPLY_FILESTAT = 213     ' File status
Global Const FTP_REPLY_HELPMSG = 214      ' Human-readable help response
Global Const FTP_REPLY_SYSTYPE = 215      ' System type
Global Const FTP_REPLY_READY = 220        ' Service ready for new user
Global Const FTP_REPLY_CLOSED = 221       ' Service closing connection
Global Const FTP_REPLY_DATAOPENED = 225   ' Data connection open
Global Const FTP_REPLY_DATACLOSED = 226   ' Closing data connection
Global Const FTP_REPLY_PASVMODE = 227     ' Entering passive mode
Global Const FTP_REPLY_LOGIN = 230        ' User logged in
Global Const FTP_REPLY_DONE = 250         ' Requested file action completed
Global Const FTP_REPLY_PATHEXISTS = 257   ' Pathname exists, created, etc.
Global Const FTP_REPLY_GETPASS = 331      ' Username okay, need password
Global Const FTP_REPLY_GETACCT = 332      ' Need account for login
Global Const FTP_REPLY_PENDING = 350      ' File action pending
Global Const FTP_REPLY_NOTAVAIL = 421     ' Service not available
Global Const FTP_REPLY_OPENFAIL = 425     ' Cannot open data connection
Global Const FTP_REPLY_ABORTED = 426      ' Connection closed, transfer aborted
Global Const FTP_REPLY_FILEBUSY = 450     ' File is not available
Global Const FTP_REPLY_LOCALERR = 451     ' Local error
Global Const FTP_REPLY_NOSPACE = 452      ' No space on server system
Global Const FTP_REPLY_BADSYN = 500       ' Syntax error
Global Const FTP_REPLY_BADARG = 501       ' Invalid command arguments
Global Const FTP_REPLY_BADCMD = 502       ' Command not implemented
Global Const FTP_REPLY_BADSEQ = 503       ' Bad sequence of commands
Global Const FTP_REPLY_BADPARM = 504      ' Command parameter not implemented
Global Const FTP_REPLY_NOLOGIN = 530      ' User not logged in
Global Const FTP_REPLY_ACCTREQ = 532      ' Account required for storing files
Global Const FTP_REPLY_NOFILE = 550       ' File unavailable
Global Const FTP_REPLY_BADPAGE = 551      ' Page type unknown
Global Const FTP_REPLY_EXQUOTA = 552      ' Exceeded file storage quota
Global Const FTP_REPLY_BADFILE = 553      ' Invalid file name
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
Public Function GetHeaderValue(ByVal strMessageHeader As String, ByVal strHeaderName As String) As String
    'Returns the header value specified by the HeaderName
    Dim intPos1 As Long
    Dim intPos2 As Long
    Dim intStart As Long
    Dim mhead As String, z$
    
    mhead = Left(strMessageHeader, 32768)
    If InStr(mhead, Chr$(10)) = 0 Then
      mhead = strrepl(mhead, Chr$(13), vbCrLf)
    End If
    intPos1 = InStr(LCase(mhead), vbCrLf + LCase(strHeaderName) + ":")
    If intPos1 = 0 Then intPos1 = InStr(LCase(mhead), LCase(strHeaderName) & ":")
    If intPos1 = 0 Then
      GetHeaderValue = ""
      Exit Function
    End If
    mhead = Right(mhead, (Len(mhead) - intPos1) - 1)
    
    intPos1 = InStr(mhead, " ")
    intStart = 1
    Do
        intPos2 = InStr(intStart, mhead, Chr(13))
        intStart = intPos2 + 3
        ' If there is a tab after the <crlf>, the header is multi-line
        ' so find the end of the next line
        z$ = Mid(mhead, intPos2 + 2, 1)
    Loop While z$ = Chr(9) Or z$ = Chr(32)
    
    If intPos1 < intPos2 Then
        GetHeaderValue = Mid(mhead, intPos1 + 1, (intPos2 - intPos1) - 1)
    Else
        GetHeaderValue = "<" & mhead & " header value missing>"
    End If
    
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

