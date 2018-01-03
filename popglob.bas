Attribute VB_Name = "popglob"
Option Explicit
Dim glob_d$(99)
Dim tm_value(9) As Long
Dim datchgmode As String

Const MAX_PATH = 260

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
    Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal _
        dwMilliSeconds As Long)

Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, _
   ByVal lpSubKey As String, ByVal ulOptions As Long, _
   ByVal samDesired As Long, phkResult As Long) _
   As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
   lpValueName As String, ByVal lpReserved As Long, _
   lpType As Long, ByVal lpData As String, lpcbData As Long) _
   As Long
                                                                                                  
Private Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long

Private Const REG_SZ As Long = 1
Private Const KEY_ALL_ACCESS = &H3F
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Type GUID
    PartOne As Long
    PartTwo As Integer
    PartThree As Integer
    PartFour(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
(ptrGuid As GUID) As Long

Public Enum NetConnTypeConstants
   INTERNET_CONNECTION_MODEM = &H1&
   INTERNET_CONNECTION_LAN = &H2&
   INTERNET_CONNECTION_PROXY = &H4&
   INTERNET_RAS_INSTALLED = &H10&
   INTERNET_CONNECTION_OFFLINE = &H20&
   INTERNET_CONNECTION_CONFIGURED = &H40&
End Enum


Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412

Private Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Private Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Private Declare Function RasEnumConnections Lib _
"rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
Any, lpcb As Long, lpcConnections As Long) As Long

Private Declare Function RasHangUp Lib "rasapi32.dll" Alias _
"RasHangUpA" (ByVal hRasConn As Long) As Long

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" _
(ByRef lpdwFlags As Long, _
ByVal lpszConnectionName As String, _
ByVal dwNameLen As Long, _
ByVal dwReserved As Long _
) As Long
Private Declare Sub keybd_event Lib "user32" ( _
  ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_KEYUP = &H2

' Virtual KeyCodes
Private Enum eVirtualKeyCode
  VK_BAK = &H8
  VK_TAB = &H9
  VK_CLEAR = &HC
  VK_RETURN = &HD
  VK_SHIFT = &H10
  VK_CONTROL = &H11
  VK_MENU = &H12
  VK_PAUSE = &H13
  VK_CAPITAL = &H14
  VK_ESCAPE = &H1B
  VK_PRIOR = &H21
  VK_NEXT = &H22
  VK_END = &H23
  VK_HOME = &H24
  VK_LEFT = &H25
  VK_UP = &H26
  VK_RIGHT = &H27
  VK_DOWN = &H28
  VK_SELECT = &H29
  VK_SNAPSHOT = &H2C  ' NEU! Windows-Taste
  VK_INSERT = &H2D
  VK_DELETE = &H2E
  VK_HELP = &H2F
  VK_F1 = &H70
  VK_F2 = &H71
  VK_F3 = &H72
  VK_F4 = &H73
  VK_F5 = &H74
  VK_F6 = &H75
  VK_F7 = &H76
  VK_F8 = &H77
  VK_F9 = &H78
  VK_F10 = &H79
  VK_F11 = &H7A
  VK_F12 = &H7B
  VK_F13 = &H7C
  VK_F14 = &H7D
  VK_F15 = &H7E
  VK_F16 = &H7F
  VK_NUMLOCK = &H90
  VK_SCROLL = &H91
  VK_WIN = &H5B     ' NEU! Windows-Taste
  VK_APPS = &H5D    ' NEU! Taste für Kontextmenü
End Enum

' Text durch Simulieren von Tastenanschlägen
' an das aktive Control senden
Public Sub SendKeysEx(ByVal sText As String)
  Dim VK As eVirtualKeyCode
  Dim sChar As String
  Dim i As Integer
  Dim bShift As Boolean
  
  ' Jedes Zeichen einzeln senden
  For i = 1 To Len(sText)
    ' aktuelles Zeichen extrahieren
    sChar = Mid$(sText, i, 1)
    
    ' Sonderzeichen?
    bShift = False
    If sChar = "{" Then
      If UCase$(Mid$(sText, i + 1, 9)) = "BACKSPACE" Then
        VK = VK_BAK
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "BS" Then
        VK = VK_BAK
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "BKSP" Then
        VK = VK_BAK
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "BREAK" Then
        VK = VK_PAUSE
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 8)) = "CAPSLOCK" Then
        VK = VK_CAPITAL
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "DELETE" Then
        VK = VK_DELETE
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "DEL" Then
        VK = VK_DELETE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "DOWN" Then
        VK = VK_DOWN
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "UP" Then
        VK = VK_UP
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "LEFT" Then
        VK = VK_LEFT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "RIGHT" Then
        VK = VK_RIGHT
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "END" Then
        VK = VK_END
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "ENTER" Then
        VK = VK_RETURN
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HOME" Then
        VK = VK_HOME
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "ESC" Then
        VK = VK_ESCAPE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HELP" Then
        VK = VK_HELP
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "INSERT" Then
        VK = VK_INSERT
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "INS" Then
        VK = VK_INSERT
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 7)) = "NUMLOCK" Then
        VK = VK_NUMLOCK
        i = i + 8
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGUP" Then
        VK = VK_PRIOR
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGDN" Then
        VK = VK_NEXT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 10)) = "SCROLLLOCK" Then
        VK = VK_SCROLL
        i = i + 11
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "TAB" Then
        VK = VK_TAB
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F1" Then
        VK = VK_F1
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F2" Then
        VK = VK_F2
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F3" Then
        VK = VK_F3
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F4" Then
        VK = VK_F4
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F5" Then
        VK = VK_F5
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F6" Then
        VK = VK_F6
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F7" Then
        VK = VK_F7
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F8" Then
        VK = VK_F8
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F9" Then
        VK = VK_F9
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F10" Then
        VK = VK_F10
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F11" Then
        VK = VK_F11
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F12" Then
        VK = VK_F12
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F13" Then
        VK = VK_F13
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F14" Then
        VK = VK_F14
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F15" Then
        VK = VK_F15
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F16" Then
        VK = VK_F16
        i = i + 4
        
      ' NEU! Windows-Taste
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "WIN" Then
        VK = VK_WIN
        i = i + 4
      
      ' NEU! Kontextmenü
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "APPS" Then
        VK = VK_APPS
        i = i + 5
        
      ' NEU! PrintScreen-Taste (DRUCK)
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "PRINT" Then
        VK = VK_SNAPSHOT
        i = i + 6
      End If
      
    ElseIf sChar = "+" Then
      ' Umschalttaste
      VK = VK_SHIFT
      
    ElseIf sChar = "%" Then
      ' ALT
      VK = VK_MENU
    
    ElseIf sChar = "^" Then
      ' STRG
      VK = VK_CONTROL
      
    Else
      ' Großbuchstabe...?
      bShift = (UCase$(sChar) = sChar And Not IsNumeric(sChar))
      If bShift Then
        ' ... dann zusätzlich Shift (Umsch)-Taste "drücken"
        keybd_event VK_SHIFT, 1, 0, 0
      End If
        
      ' Virtual KeyCode ermitteln...
      VK = Asc(UCase$(sChar))
    End If
    
    ' niederdrücken und wieder loslassen
    keybd_event VK, 1, 0, 0
    keybd_event VK, 1, KEYEVENTF_KEYUP, 0
        
    ' Shift (Umsch)-Taste wieder loslassen
    If bShift Then
      keybd_event VK_SHIFT, 1, KEYEVENTF_KEYUP, 0
    End If
  Next i
End Sub

Private Function InternetConnected( _
    Optional ByRef eConnectionInfo As NetConnTypeConstants, _
    Optional ByRef sConnectionName As String _
    ) As Boolean
   
    Dim dwFlags As Long
    Dim sNameBuf As String
    Dim lR As Long
    Dim iPos As Long
    
    On Error GoTo exfu_InternetConnected
    sNameBuf = String$(513, 0)
    lR = 0
    lR = InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&)
    eConnectionInfo = dwFlags
    iPos = InStr(sNameBuf, vbNullChar)
    
    If iPos > 0 Then
        sConnectionName = Left$(sNameBuf, iPos - 1)
    ElseIf Not sNameBuf = String$(513, 0) Then
        sConnectionName = sNameBuf
    End If
exfu_InternetConnected:
    InternetConnected = (lR = 1)
End Function


Public Property Get IsConnected() As Boolean
    IsConnected = InternetConnected()
End Property

Public Property Get ConnType() As Long
    Dim connInfo As NetConnTypeConstants
    InternetConnected connInfo
    ConnType = connInfo
End Property

Public Function ConnTypeDevice(nType As Long) As String
    Dim strReturn As String
    
    If nType And INTERNET_CONNECTION_LAN Then
        strReturn = "LAN"
    ElseIf nType And INTERNET_CONNECTION_MODEM Then
        strReturn = "Modem"
    ElseIf nType And INTERNET_CONNECTION_PROXY Then
        strReturn = "Proxy"
    ElseIf nType And INTERNET_CONNECTION_OFFLINE Then
        strReturn = "Offline"
    End If
    
    ConnTypeDevice = strReturn
End Function

Public Property Get ConnName() As String
    Dim strName As String
    InternetConnected , strName
    ConnName = strName
End Property

Public Sub HangUp()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    Dim ReturnCode As Long
    Dim gstrISPName As String
    
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
    lpcConnections)

    If ReturnCode = 0 Then
        For i = 0 To lpcConnections - 1
            If trm(ByteToString(lpRasConn(i).szEntryName)) = trm(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
            End If
        Next i
    End If
End Sub

Private Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & ChrB$(bytString(i))
        i = i + 1
    Wend
End Function

Public Function MyMin(a, b) As Variant
If a < b Then
  MyMin = a
Else
  MyMin = b
End If

End Function

Public Function hex2dec(hx$) As Integer
Dim rc%, a$, h$

If Len(hx$) <> 2 Then rc% = 0
h$ = UCase(hx$)
a$ = "0123456789ABCDEF"
rc% = (InStr(a$, Left$(h$, 1)) - 1) * 16
rc% = rc% + (InStr(a$, Right$(h$, 1)) - 1)
hex2dec = rc%
End Function

Public Function imin(a%, b%) As Integer

imin = a%
If b% < a% Then imin = b%

End Function
Public Function imax(a%, b%) As Integer

imax = a%
If b% > a% Then imax = b%

End Function
Public Function nouml(l$) As String

l$ = strrepl(l$, "ä", "ae")
l$ = strrepl(l$, "ö", "oe")
l$ = strrepl(l$, "ü", "ue")
l$ = strrepl(l$, "Ö", "Oe")
l$ = strrepl(l$, "Ä", "Ae")
l$ = strrepl(l$, "Ü", "Ue")
nouml = l$
End Function
Public Function mkalphanum(wrd$) As String
Dim rc$, i%, z$

rc$ = ""
For i% = 1 To Len(wrd$)
  z$ = Mid$(wrd$, i%, 1)
  If (z$ >= "0" And z$ <= "9") Or (z$ >= "a" And z$ <= "z") Or (z$ >= "A" And z$ <= "Z") Then
    rc$ = rc$ + z$
  End If
Next i%
mkalphanum = rc$
End Function

Public Function IsTime(ByVal Time As Variant) As Boolean
'*****************************************************************
'* This subroutine determines if a value is a valid time
'* (not  date).
'************************************************************

  IsTime = False
  If IsDate(Time) = True Then
    If CStr(Time) Like "*#*.*#*" = False Then
      If Fix(CDate(Time)) = 0 Then
        If CStr(Time) Like "1[3-9]*[aApP]*" = False Then
          If CStr(Time) Like "2[0-3]*[aApP]*" = False Then
            IsTime = True
          End If
        End If
      End If
    End If
  End If
  
End Function

Public Function domainofemail(adr$)
Dim l$, p%
domainofemail = ""
If IsValidEmail(adr$) = False Then Exit Function
p% = InStr(adr$, "@")
If p% > 0 And p% < Len(adr$) - 1 Then
  domainofemail = Mid$(adr$, p% + 1)
End If
End Function

Public Function emailonly(adr$) As String
Dim p%, l$
Debug.Print adr$
emailonly = ""
p% = InStr(adr$, "<")
If p% > 0 And p% < Len(adr$) - 1 Then
  l$ = Mid$(adr$, p% + 1)
  p% = InStr(l$, ">")
  If p% > 1 Then l$ = Left(l$, p% - 1)
  l$ = strrepl(l$, ">", "")
Else
  l$ = adr$
End If
p% = InStr(l$, ">")
If p% > 1 Then l$ = Left(l$, p% - 1)
l$ = strrepl(l$, ">", "")
If IsValidEmail(l$) = True Then
  emailonly = l$
End If
Debug.Print emailonly
End Function

Public Function IsValidEmail(sEMail As String) As Boolean
    ' original by Brad Murray
    ' optimized by Rob Hofker, email: rob@eurocamp.nl,
     '23 august 2000
    
    Dim sInvalidChars As String
    Dim bTemp As Boolean
    Dim i As Integer
    Dim sTemp As String

    ' Disallowed characters
    ' sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "
    ' spammers used %
    sInvalidChars = "!#$^&*()=+{}[]|\;:'/?>,< "

    ' Check that there is at least one '@'
    bTemp = InStr(sEMail, "@") <= 0
    If bTemp Then GoTo exit_function

    ' Check that there is at least one '.'
    bTemp = InStr(sEMail, ".") <= 0
    If bTemp Then GoTo exit_function

    ' and that the length is at least six (a@a.ca)
    bTemp = Len(sEMail) < 6
    If bTemp Then GoTo exit_function

    ' Check that there is only one '@'
    i = InStr(sEMail, "@")
    sTemp = Mid(sEMail, i + 1)
    bTemp = InStr(sTemp, "@") > 0
    
    If bTemp Then GoTo exit_function
    'extra checks
    ' AFTER '@' space is not allowed
    bTemp = InStr(sTemp, " ") > 0
    If bTemp Then GoTo exit_function

    ' do not Check that there is one dot AFTER '@'
    'bTemp = InStr(sTemp, ".") = 0
    'If bTemp Then GoTo exit_function
    
    ' Check if there's a quote (")
    bTemp = InStr(sEMail, Chr(34)) > 0
    If bTemp Then GoTo exit_function
    
        
    ' Check if there's any other disallowed chars
    ' optimize a little if sEmail longer than sInvalidChars
    ' check the other way around
    If Len(sEMail) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(sEMail, Mid(sInvalidChars, i, 1)) > 0 _
                  Then bTemp = True
            If bTemp Then Exit For
        Next
    Else
        For i = 1 To Len(sEMail)
            If InStr(sInvalidChars, Mid(sEMail, i, 1)) > 0 _
                   Then bTemp = True
            If bTemp Then Exit For
        Next
    End If
    If bTemp Then GoTo exit_function
    
    ' extra check
    ' no two consecutive dots
    bTemp = InStr(sEMail, "..") > 0
    If bTemp Then GoTo exit_function
    
exit_function:
    ' if any of the above are true, invalid e-mail
    IsValidEmail = Not bTemp

End Function


Public Sub change_field_size(DBPath As String, _
  tblName As String, fldName As String, fldSize As Integer)
    ' this routine changes the field size
    
    Dim db As Database
    Dim td As TableDef
    Dim fld As Field
        
    On Error GoTo errhandler

    Set db = OpenDatabase(DBPath)
    Set td = db.TableDefs(tblName)
    
    If td.Fields(fldName).Type <> dbText Then
        ' wrong field type
        db.Close
        Exit Sub
    End If
    
    If td.Fields(fldName).Size = fldSize Then
        ' the field width is correct
        db.Close
        Exit Sub
    End If
    
    ' create a temp feild
    td.Fields.Append td.CreateField("temp", dbText, fldSize)
    td.Fields("temp").AllowZeroLength = True
    td.Fields("temp").DefaultValue = """"""

    ' copy the info into the temp field
    db.Execute "Update " & tblName & " set temp = " & fldName & " "
    
    ' delete the field
    td.Fields.Delete fldName
    
    ' rename the field
    td.Fields("temp").Name = fldName
    db.Close
    
'======================================================================
Exit Sub

errhandler:
MsgBox CStr(Err.Number) & vbCrLf & Err.Description & vbCrLf & "Change Field Size Routine", vbCritical, App.Title

End Sub

Public Function isvaldate(d$)
Dim i%

isvaldate = 0
For i% = 1 To 4
  If isdigit(Mid$(d$, i%, 1)) = 0 Then Exit Function
Next i%
For i% = 9 To 10
  If isdigit(Mid$(d$, i%, 1)) = 0 Then Exit Function
Next i%
For i% = 6 To 7
  If isdigit(Mid$(d$, i%, 1)) = 0 Then Exit Function
Next i%
If isdigit(Mid$(d$, 5, 1)) = 1 Then Exit Function
If isdigit(Mid$(d$, 8, 1)) = 1 Then Exit Function
isvaldate = 1
End Function
Public Function FileName(fn$) As String
Dim r$, p%

FileName = fn$
If InStr(fn$, "\") = 0 Then Exit Function
r$ = fn$
p% = Len(r$)
While p% > 0 And Mid$(r$, p%, 1) <> "\":  p% = p% - 1: Wend
If p% > 0 Then FileName = Mid$(r$, p% + 1)

End Function
Public Function FileExtension(fn$) As String
Dim p%, f$

FileExtension = ""
f$ = FileName(fn$)
p% = InStr(f$, ".")
If p% = 0 Then Exit Function
FileExtension = Mid$(f$, p% + 1)
While Left(FileExtension, 1) = ".": FileExtension = Mid(FileExtension, 2): Wend

End Function
Public Function DirName(fn$) As String
Dim r$, p%

DirName = ""
If InStr(fn$, "\") = 0 Then Exit Function
r$ = fn$
p% = Len(r$)
While p% > 0 And Mid$(r$, p%, 1) <> "\":  p% = p% - 1: Wend
If p% > 0 Then DirName = Left$(r$, p% - 1)

End Function
Public Function basename(fn$, ext$) As String
Dim r$, p%

basename = fn$
r$ = fn$
While InStr(r$, "\") > 0: r$ = Mid$(r$, InStr(r$, "\") + 1): Wend
If ext$ = "" Then
  basename = r$
  Exit Function
End If
p% = InStr(LCase(r$), LCase(ext$))
If p% > 0 Then
  basename = Left$(r$, p% - 1)
  Exit Function
End If
p% = InStr(LCase(r$), ".")
If p% > 0 Then
  basename = Left$(r$, p% - 1)
  Exit Function
End If

End Function
Public Function word1(l1$) As String
Dim l$

l$ = trm(l1$)
If InStr(l$, " ") > 0 Then
  word1 = Left$(l$, InStr(l$, " ") - 1)
Else
  word1 = l$
End If
End Function
Public Function word2(l1) As String
Dim i As Integer, l As String

l = ""
i = InStr(trm(l1), " ")
If i > 0 Then
  l = Mid(trm(l1), i + 1)
  l = word1(l)
End If
word2 = l
End Function
Public Function word2bis(l1$) As String
Dim i%, l$

l$ = ""
i% = InStr(trm(l1$), " ")
If i% > 0 Then
  l$ = Mid$(trm(l1$), i% + 1)
End If
word2bis = l$
End Function
Public Function isnumber(l1$) As Boolean
Dim i%, l$, z$

isnumber = True
l$ = trm(l1$)
If l$ = "" Then
  isnumber = False
  Exit Function
End If
For i% = 1 To Len(l$)
  z$ = Mid$(l$, i%, 1)
  If isdigit(z$) = 0 And z$ <> "-" And z$ <> "+" Then
    isnumber = False
    Exit Function
  End If
Next i%

End Function
Public Function isnumberrange(l1$) As Boolean
Dim i%, l$, z$

isnumberrange = True
l$ = trm(l1$)
For i% = 1 To Len(l$)
  z$ = Mid$(l$, i%, 1)
  If isdigit(z$) = 0 And z$ <> " " And z$ <> "-" Then
    isnumberrange = False
    Exit Function
  End If
Next i%

End Function
Public Function lastword(l1$) As String
Dim l$, rl$

rl$ = l1$
Do
  l$ = word1(rl$)
  rl$ = word2bis(rl$)
Loop Until rl$ = ""
lastword = l$
End Function
Public Function d2db(Text) As String
Dim t$

t$ = trm(Text)
t$ = strrepl(t$, ",", ".")
d2db = t$

End Function
Public Function fixl(l$, le%) As String

fixl = l$
If Len(l$) >= le% Then Exit Function
fixl = l$ + Space$(le% - Len(l$))

End Function

Public Function fixl0(l$, le%) As String
Dim sp As String

sp = Space$(le% + 1): sp = strrepl(sp, " ", "0")
fixl0 = l$
If Len(l$) >= le% Then Exit Function
fixl0 = Left(sp, le% - Len(l)) + l$

End Function

Public Function isdigit(char$)

isdigit = InStr("1234567890", char$)

End Function

Public Function hasdigit(C$) As Boolean
Dim i%
hasdigit = False
For i% = 1 To Len(C$)
  If isdigit(Mid$(C$, i%, 1)) Then
    hasdigit = True
    Exit Function
  End If
Next i%

End Function
Public Function isalpha(C$)

isalpha = 0
If (C$ >= "a" And C$ <= "z") Or (C$ >= "A" And C$ <= "Z") Then isalpha = 1

End Function
Public Function mknam(l%) As String
Dim i%, rc$, v%, K%, z$

rc$ = Chr$(Rnd * 25 + 65)
v% = 0
K% = 0
i% = l% - 1
While i% > 0
  Do
    z$ = Chr$(Int(Rnd * 25 + 65))
  Loop Until Sgn(isvocal(z$)) <> Sgn(isvocal(Right$(rc$, 1)))
  rc$ = rc$ + z$
  i% = i% - 1
Wend

mknam = UCase(Left$(rc$, 1)) + LCase$(Mid$(rc$, 2))

End Function
Public Function mkkey(l%) As String
Dim i%, rc$, v%, K%, z$

i% = l%
While i% > 0
  z$ = Chr$(Int(Rnd * 25 + 65))
  rc$ = rc$ + z$
  i% = i% - 1
Wend

mkkey = rc$

End Function
Function isvocal(z$)

isvocal = InStr("aeiouAEIOU", z$)

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
Public Function UcaseFirstLetter(Text$) As String
Dim t$

t$ = Text$
If Len(t$) > 1 Then
  t$ = UCase(Left$(t$, 1)) + Mid$(t$, 2)
Else
  t$ = UCase(t$)
End If

UcaseFirstLetter = t$

End Function

Public Function nexist(fn$) As Boolean
Dim o%, rrr

o% = FreeFile
On Error Resume Next
Open fn$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Close #o%
  nexist = False
Else
  nexist = True
End If

End Function

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

Public Function onlynums(i$) As String
Dim rc$, j%, z$

rc$ = ""

For j% = 1 To Len(i$)
  z$ = Mid$(i$, j%, 1)
  If isdigit(z$) > 0 Then
    rc$ = rc$ + z$
  End If
Next j%
onlynums = rc$

End Function

Public Function fixeur(d As Double) As String
Dim rc$, p%, r2$, i%, s As String, usgn As String
fixeur = "0.00"
rc$ = str(Int(100 * d + 0.5) / 100)
'rc$ = str(Int(100 * d) / 100)
p% = InStr(rc$, ".")
If p% = 0 Then
  rc$ = rc$ + ".00"
Else
  While p% > Len(rc$) - 2
    rc$ = rc$ + "0"
    p% = InStr(rc$, ".")
  Wend
End If
If Left$(rc$, 1) = "," Then rc$ = "0" & rc$
rc$ = trm(strrepl(rc$, ".", ","))
s = "": usgn = rc$
If Left(rc$, 1) = "-" Then
  s = "-"
  usgn = Mid(rc$, 2)
End If
p% = InStr(usgn, ",")
If p% > 4 Then
  r2$ = Right(usgn, 3)
  i% = 0
  For p% = Len(usgn) - 3 To 1 Step -1
    i% = i% + 1
    If i% > 3 Then
      i% = 1
      r2$ = "." & r2$
    End If
    r2$ = Mid$(usgn, p%, 1) & r2$
  Next p%
Else
  r2$ = usgn
End If
fixeur = s + r2$
End Function

Public Function fixeurnozerotail(d As Double) As String
Dim r$
r$ = fixeur(d)
While Right$(r$, 1) = "0" And r$ <> ""
  r$ = Left(r$, Len(r$) - 1)
Wend
If Right$(r$, 1) = "," Then r$ = Left(r$, Len(r$) - 1)
fixeurnozerotail = r$
End Function

Public Sub flist(d$, ofn$)
Dim o%, tr$

tr$ = Dir(d$ & "\*.*")
While tr$ <> ""
  If tr$ <> "." And tr$ <> ".." Then
    Debug.Print d$ & "\" & tr$
    o% = FreeFile
    Open ofn$ For Append As #o%
    Print #o%, d$ & "\" & tr$
    Close #o%
  End If
  tr$ = Dir
Wend

End Sub
Public Sub dlist(d$, ofn$)
Dim o%, tr$, i%

Call flist(d$, ofn$)
tr$ = Dir(d$ & "\*.*", vbDirectory)
i% = 0
While tr$ <> ""
  If tr$ <> "." And tr$ <> ".." Then
    If (GetAttr(d$ & "\" & tr$) And vbDirectory) = vbDirectory Then
      glob_d$(i%) = d$ & "\" & tr$
      i% = i% + 1
    End If
  End If
  tr$ = Dir
Wend
If i% > 0 Then
  While i% > 0
    i% = i% - 1
    Call flist(glob_d$(i%), ofn$)
  Wend
End If

End Sub
'* Parameters: 0=cut, 1=copy, 2=paste, 3=delete
Sub CutCopyPaste(DoWhat As Integer)
    ' ActiveForm refers to the active form in the MDI form.
    If TypeOf Screen.ActiveControl Is TextBox Then
        Select Case DoWhat
            Case 0                      ' Cut.
                ' Copy selected text to Clipboard.
                Clipboard.SetText Screen.ActiveControl.SelText
                ' Delete selected text.
                Screen.ActiveControl.SelText = ""
            Case 1                      ' Copy.
                ' Copy selected text to Clipboard.
                Clipboard.SetText Screen.ActiveControl.SelText
            Case 2                      ' Paste.
                ' Put Clipboard text in text box.
                Screen.ActiveControl.SelText = Clipboard.GetText()
            Case 3                      ' Delete.
                ' Delete selected text.
                Screen.ActiveControl.SelText = ""
        End Select
    End If
End Sub
'RETURNS:  GUID if successful; blank string otherwise.
'Unlike the GUIDS in the registry, this function returns GUID
'without "-" characters.  See comments for how to modify if you
'want the dash.

Public Function GUID() As String
    Dim lRetVal As Long
    Dim udtGuid As GUID
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
   
    On Error GoTo errorhandler
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
    
       'First 8 chars
        sPartOne = Hex$(udtGuid.PartOne)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.PartOne)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartOne)
        
        'Next 4 Chars
        sPartTwo = Hex$(udtGuid.PartTwo)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.PartTwo)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartTwo)
           
        'Next 4 Chars
        sPartThree = Hex$(udtGuid.PartThree)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.PartThree)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartThree)   'Next 2 bytes (4 hex digits)
           
        'Final 16 chars
        For iCtr = 0 To 7
            sPartFour = sPartFour & _
            Format$(Hex$(udtGuid.PartFour(iCtr)), "00")
        Next
 
     'To create GUID with "-", change line below to:
     'sAns = sPartOne & "-" & sPartTwo & "-" & sPartThree _
     '& "-" & sPartFour
       
       sAns = sPartOne & sPartTwo & sPartThree & sPartFour
            
        End If
        
        GUID = sAns
Exit Function


errorhandler:
'return a blank string if there's an error
Exit Function
End Function
'*****************************************************
'These functions return the path to the specified office
'application or a 0-length string if the application does not
'exist on the machine.  This is one good way to check whether a
'specific office application is present before trying to run
'automation code for that application
'*****************************************************
Public Function GetWordPath() As String
    GetWordPath = GetOfficeAppPath("Word.Application")
End Function

Public Function GetExcelPath() As String
    GetExcelPath = GetOfficeAppPath("Excel.Application")
End Function

Public Function GetAccessPath() As String
    GetAccessPath = GetOfficeAppPath("Access.Application")
End Function

Public Function GetOutlookPath() As String
    GetOutlookPath = GetOfficeAppPath("Outlook.Application")
End Function

Public Function GetPowerPointPath() As String
    GetPowerPointPath = _
       GetOfficeAppPath("PowerPoint.Application")
End Function

Public Function GetFrontPagePath() As String
    GetFrontPagePath = GetOfficeAppPath("FrontPage.Application")
End Function

Private Function GetOfficeAppPath(ByVal ProgID As String) _
   As String

Dim lKey As Long
Dim lRet As Long
Dim sClassID As String
Dim sAns As String
Dim lngBuffer As Long
Dim lPos As Long

   'GetClassID
   lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
          "Software\Classes\" & ProgID & "\CLSID", 0&, _
           KEY_ALL_ACCESS, lKey)
   If lRet = 0 Then
 
      lRet = RegQueryValueEx(lKey, "", 0&, REG_SZ, "", lngBuffer)
      sClassID = Space(lngBuffer)
      lRet = RegQueryValueEx(lKey, "", 0&, REG_SZ, sClassID, _
          lngBuffer)

      'drop null-terminator
      sClassID = Left(sClassID, lngBuffer - 1)
      RegCloseKey lKey
   End If
   
    
   'Get AppPath
    lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
        "Software\Classes\CLSID\" & sClassID & _
        "\LocalServer32", 0&, KEY_ALL_ACCESS, lKey)
 
  If lRet = 0 Then
      lRet = RegQueryValueEx(lKey, "", 0&, REG_SZ, "", lngBuffer)
      sAns = Space(lngBuffer)
      lRet = RegQueryValueEx(lKey, "", 0&, REG_SZ, sAns, _
        lngBuffer)
      sAns = Left(sAns, lngBuffer - 1)
      
      RegCloseKey lKey
   End If
    
    
    'Sometimes the registry will return a switch
       'beginning with "/" e.g., "/automation"
    
    lPos = InStr(sAns, "/")
        If lPos > 0 Then
            sAns = trm(Left(sAns, lPos - 1))
        End If
    
    GetOfficeAppPath = sAns
    
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
End Function


Public Function hex2(w) As String
Dim hx$

hx$ = Hex$(w)
If Len(hx$) < 2 Then hx$ = "0" & hx$
hex2 = hx$
End Function

Public Function mkhttp(url$) As String
Dim rc$, u$

u$ = trm(url$)
If InStr(u$, "://") = 0 Then
  mkhttp = "http://" & u$
Else
  mkhttp = u$
End If

End Function

Public Function trm(l) As String
Dim rrr
On Error Resume Next
trm = Trim("" & l)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then trm = ""
End Function

Public Function trm0(l) As String
Dim rrr, rc$

On Error Resume Next
rc$ = trm("" & l)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  trm0 = "0"
  Exit Function
End If
If rc$ = "" Then rc$ = "0"
trm0 = rc$

End Function
Public Function datum2sql(dtg) As String
Dim y$, rrr, M$, d$

datum2sql = ""
If Len(dtg) > 0 Then

If datchgmode = "en" And InStr(dtg, "-") > 0 Then
  datum2sql = dtg
End If

On Error Resume Next
y$ = Year(dtg)
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  M$ = Format$(Month(dtg), "00")
  d$ = Format$(Day(dtg), "00")
  datum2sql = y$ + "-" + M$ + "-" + d$
End If
End If

End Function
Public Function datfromsql(dtg) As String

datfromsql = ""
If trm(dtg) <> "" Then
  Select Case datchgmode
    Case "en": datfromsql = Left(dtg, 4) + "-" + Mid(dtg, 6, 2) + "-" + Right(dtg, 2)
    Case Else: datfromsql = Right(dtg, 2) + "." + Mid(dtg, 6, 2) + "." + Left(dtg, 4)
  End Select
End If

End Function
Public Function datfromsqlshort(dtg) As String

datfromsqlshort = ""
If trm(dtg) <> "" Then
  Select Case datchgmode
    Case "en": datfromsqlshort = Mid(dtg, 3, 2) + "-" + Mid(dtg, 6, 2) + "-" + Right(dtg, 2)
    Case Else: datfromsqlshort = Right(dtg, 2) + "." + Mid(dtg, 6, 2) + "." + Mid(dtg, 3, 2)
  End Select
End If

End Function

Public Function trmx1(l1) As String
Dim l$, r$

l$ = "" & l1
r$ = trm(l$)
r$ = strrepl(r$, "'", "´")
trmx1 = r$
End Function

Public Function initialen(l1) As String
Dim l$, r$, w1$

l$ = trm(l1)
While Len(l$) > 0
  w1$ = word1(l$)
  r$ = r$ & UCase(Left$(w1$, 1))
  l$ = trm(Mid$(l$, Len(w1$) + 1))
Wend
initialen = r$
End Function

Function linesof(l$) As Integer
Dim fl$, ll$, rl$, i%, n%, p%

fl$ = trm(l$)
i% = 0
While InStr(fl$, Chr$(10)) > 0
  p% = InStr(fl$, Chr$(10))
  If p% > 0 Then
    ll$ = trm(Left$(fl$, p% - 1))
    rl$ = trm(Mid$(fl$, p% + 1))
    ll$ = strrepl(ll$, Chr$(13), "")
    rl$ = strrepl(rl$, Chr$(13), "")
    fl$ = rl$
  End If
  i% = i% + 1
Wend
If Len(fl$) > 0 Then i% = i% + 1
linesof = i%

End Function

Function lineof(n%, l$) As String
Dim brk As Boolean
Dim fl$, ll$, rl$, i%, p%

fl$ = trm(l$)
i% = n%
brk = False
While i% > 0 And Not brk
  p% = InStr(fl$, Chr$(10))
  If p% > 0 Then
    ll$ = trm(Left$(fl$, p% - 1))
    rl$ = trm(Mid$(fl$, p% + 1))
    ll$ = strrepl(ll$, Chr$(13), "")
    rl$ = strrepl(rl$, Chr$(13), "")
    fl$ = rl$
  Else
    If fl$ = "" Then brk = True
    ll$ = fl$
    fl$ = ""
  End If
  i% = i% - 1
Wend
lineof = ll$
End Function
Function lineof_notrim(n%, l$) As String
Dim brk As Boolean
Dim fl$, ll$, rl$, i%, p%

fl$ = l$
i% = n%
brk = False
While i% > 0 And Not brk
  p% = InStr(fl$, Chr$(10))
  If p% > 0 Then
    ll$ = Left$(fl$, p% - 1)
    rl$ = Mid$(fl$, p% + 1)
    ll$ = strrepl(ll$, Chr$(13), "")
    rl$ = strrepl(rl$, Chr$(13), "")
    fl$ = rl$
  Else
    If fl$ = "" Then brk = True
    ll$ = fl$
    fl$ = ""
  End If
  i% = i% - 1
Wend
lineof_notrim = ll$
End Function

Function lastlineof(l$) As String
Dim fl$, ll$, rl$, i%, n%, p%

fl$ = trm(l$)
i% = n%
While InStr(fl$, Chr$(10)) > 0
  p% = InStr(fl$, Chr$(10))
  If p% > 0 Then
    ll$ = trm(Left$(fl$, p% - 1))
    rl$ = trm(Mid$(fl$, p% + 1))
    ll$ = strrepl(ll$, Chr$(13), "")
    rl$ = strrepl(rl$, Chr$(13), "")
    fl$ = rl$
  End If
  i% = i% - 1
Wend
lastlineof = fl$
End Function

Sub app2file(fn$, l$)
Dim o%
o% = FreeFile
Open fn$ For Append As #o%
Print #o%, l$
Close #o%
End Sub

Public Function ohnePLZ(s$) As String
Dim brk%, C$, s99$

s99$ = s$
brk% = 0
Do
  C$ = Left$(s99$, 1)
  If C$ = " " Or (C$ >= "0" And C$ <= "9") Then
    s99$ = Mid$(s99$, 2)
  Else
    brk% = 1
  End If
Loop Until brk% = 1
ohnePLZ = s99$

End Function

Public Function nurdiePLZ(s$) As String
Dim rc$, brk%, C$, s99$

s99$ = s$
rc$ = ""
brk% = 0
Do
  C$ = Left$(s99$, 1)
  If C$ = " " Or (C$ >= "0" And C$ <= "9") Then
    rc$ = rc$ & C$
    s99$ = Mid$(s99$, 2)
  Else
    brk% = 1
  End If
Loop Until brk% = 1
nurdiePLZ = rc$

End Function

Public Function encrypt(was$, womit$) As String
Dim aKey() As Byte, rc$, i%, erg$, enc$, z$

encrypt = ""
enc$ = womit$
aKey = enc$
Call blf_KeyInit(aKey())
rc$ = blf_StringEnc(was$)
For i% = 1 To Len(rc$)
  z$ = Hex$(Asc(Mid$(rc$, i%, 1)))
  If Len(z$) = 1 Then z$ = "0" & z$
  erg$ = erg$ & z$
Next i%
encrypt = erg$

End Function

Public Function decrypt(was$, womit$) As String
Dim aKey() As Byte, rc$, i%, erg$
    
    aKey() = womit$
    For i% = 1 To Len(was$) - 1 Step 2
      rc$ = rc$ & Chr$(popglob.hex2dec(Mid$(was$, i%, 2)))
    Next i%
    Call blf_KeyInit(aKey())
    erg$ = blf_StringDec(rc$)
    decrypt = erg$

End Function

Public Function opjahr(anfangszeichen$, jahr$, von$, bis$, endezeichen$) As String
Dim j$, v$, b$, r$

opjahr = "": r$ = ""

j$ = jahr$: If j$ = "0" Then j$ = ""
v$ = von$: If v$ = "0" Then v$ = ""
b$ = bis$: If b$ = "0" Then b$ = ""

If j$ <> "" Then
  r$ = j$
Else
  If b$ <> "" Then
    r$ = b$
  Else
    r$ = v$
  End If
End If
If r$ <> "" Then
  opjahr = anfangszeichen$ & r$ & endezeichen$
End If

End Function
Public Function cut_d1(w$, term$) As String
Dim p%

p% = InStr(w$, term$)
If p% = 0 Then
  cut_d1 = w$
Else
  cut_d1 = Left$(w$, p% - 1)
End If

End Function
Public Function cut_d2bis(w$, term$) As String
Dim p%

p% = InStr(w$, term$)
If p% = 0 Then
  cut_d2bis = ""
Else
  cut_d2bis = trm(Mid$(w$, p% + 1))
End If

End Function

Public Sub tm_start(nr%)
tm_value(nr%) = GetTickCount()

End Sub

Public Function tm_stop(nr%) As Long
Dim l As Long

l = GetTickCount()
tm_stop = l - tm_value(nr%)
End Function

Public Function transex1(txt, delim As String) As String
Dim p As Integer, tx As String

If trm(txt) = "" Then
  transex1 = ""
  Exit Function
End If
p = InStr(txt, delim)
If p > 1 Then
  tx = Left(txt, p - 1)
  transex1 = Form1.inmylanguage(tx) + Mid(txt, p)
Else
  tx = txt
  transex1 = Form1.inmylanguage(tx)
End If

End Function

Public Function transe(txt) As String
Dim tx As String

If trm(txt) = "" Then
  transe = ""
  Exit Function
End If
tx = txt
transe = Form1.inmylanguage(tx)
End Function

Public Function transo(txt) As String
Dim tx As String

If trm(txt) = "" Then
  transo = ""
  Exit Function
End If
tx = txt
transo = Form1.outmylanguage(tx)
End Function

Public Sub set_datchgmode(mde As String)
datchgmode = mde
End Sub
Public Function nurstrasse(s$) As String
Dim i%, rc$, z$

nurstrasse = s$
rc$ = ""
i% = 1
While i% <= Len(s$)
  z$ = Mid$(s$, i%, 1)
  If isdigit(z$) Then
    nurstrasse = trm(rc$)
    Exit Function
  End If
  rc$ = rc$ + z$
  i% = i% + 1
Wend

End Function

Public Sub delay(s As Long)
Dim i As Long

  i = s * 10
  While i > 0
    i = i - 1
    Sleep (100)
    DoEvents
  Wend

End Sub

Public Function removegarbtrailfromnumber(l$) As String
Dim r$

r$ = l$
While Len(r$) > 0 And isdigit(Right$(r$, 1)) = 0
  r$ = Left$(r$, Len(r$) - 1)
Wend
If r$ = "" Then r$ = 0
removegarbtrailfromnumber = r$
End Function

Sub wait(s%)
Dim wt As Double, wt0 As Double

wt0 = (Date + Time) + (s% / 85400)
Do
  DoEvents
  wt = Date + Time
  wt = (wt0 - wt) * 86400
  Debug.Print wt
Loop Until wt < 0
End Sub

Public Function FileSize(ByVal sFile As String) As Long
  ' Größe einer Datei ermitteln
  ' funktioniert bis 4 GB!
  Dim nSize As Long
  
  On Error Resume Next
  If Dir$(sFile) <> "" Then
    nSize = FileLen(sFile)
    If Err.Number <> 0 Then
      ' Fehler: evtl. ist Datei größer als 4 GB
      FileSize = 2147483648# + 2147483648#
    Else
      If nSize < 0 Then
        ' Datei ist größer als 2GB!
        FileSize = 2147483648# + (2147483648# - Abs(nSize))
      Else
        FileSize = nSize
      End If
    End If
  Else
    ' Falls Datei nicht gefunden,
    ' -1 als Wert zurückgeben
    FileSize = -1
  End If
  On Error GoTo 0
End Function

Public Function dirlist(pfad As String) As String
Dim tr

dirlist = ""
On Error GoTo exdirlist
    tr = Dir(pfad, vbDirectory)
    Do While tr <> ""
      If (GetAttr(pfad + tr) And vbDirectory) = vbDirectory Then
        If tr <> "." And tr <> ".." Then
          If dirlist <> "" Then dirlist = dirlist + "|"
          dirlist = dirlist + tr
        End If
      End If
      tr = Dir
    Loop
exdirlist:
On Error GoTo 0

End Function

Public Function time2minutes(tstr As String) As Integer
Dim C$, rc As Integer, rrr

C$ = tstr
time2minutes = -1
If C$ = "" Then Exit Function
If Len(C$) = 3 Then C$ = "0" + C$
If Len(C$) = 4 Then C$ = Left$(C$, 2) + ":" + Right$(C$, 2)
If Mid$(C$, 2, 1) = ":" Then C$ = "0" + C$
On Error Resume Next
rc = Val(Left$(C$, 2)) * 60 + Val(Mid$(C$, 4, 2))
rrr = Err
On Error GoTo 0
If rrr = 0 Then time2minutes = rc
End Function

Public Function var2dbl(ByVal wert As Variant) As Double
Dim vz As Boolean, usgn As String, rc As Double

vz = False: usgn = trm(wert)
If Left(usgn, 1) = "-" Then
  vz = True
  usgn = trm(Mid(usgn, 2))
End If
rc = CDbl("0" + usgn)
If vz Then rc = -rc
var2dbl = rc

End Function

Public Function GetDriveStrings() As String
    ' Wrapper for calling the GetLogicalDriveStrings API

    Dim Result As Long          ' Result of our api calls
    Dim strDrives As String     ' String to pass to api call
    Dim lenStrDrives As Long    ' Length of the above string

    ' Call GetLogicalDriveStrings with a buffer size of zero to
    ' find out how large our stringbuffer needs to be
    Result = GetLogicalDriveStrings(0, strDrives)

    strDrives = String(Result, 0)
    lenStrDrives = Result

    ' Call again with our new buffer
    Result = GetLogicalDriveStrings(lenStrDrives, strDrives)

    If Result = 0 Then
        ' There was some error calling the API
        ' Pass back an empty string
        ' NOTE - TODO: Implement proper error handling here
        GetDriveStrings = ""
    Else
        GetDriveStrings = strDrives
    End If
End Function

Public Function ausklammern(wert$, K$) As String
Dim lvl As Integer, rc$, i As Integer, z As String
Dim ka As String, ke As String

rc$ = "": lvl = 0
ka = Left(K$, 1)
ke = Mid(K$, 2, 1)
For i = 1 To Len(wert$)
  z = Mid(wert$, i, 1)
  If z = ka Then lvl = lvl + 1
  If lvl = 0 Then rc$ = rc$ + z
  If z = ke Then lvl = lvl - 1
Next i
ausklammern = rc$
End Function

Public Function lMin(a As Long, b As Long) As Long

lMin = a
If b < a Then lMin = b

End Function

Public Function lMax(a As Long, b As Long) As Long

lMax = a
If b > a Then lMax = b

End Function


