VERSION 5.00
Begin VB.Form setupmain 
   Caption         =   "Agencyprof Demo Setup"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   2280
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "setupmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
            Alias "DeleteUrlCacheEntryA" _
            (ByVal lpszUrlName As String) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Dim swait As Integer, myip$, myid$

Private Sub Timer1_Timer()
Dim x, cmd As String, o%, pw$, xb As Boolean, nfn$, url$, pack$, adm$, brw$

If swait < -1000 Then
  Timer1.Interval = 10000
  If Not nexist("C:\Agencyprof\ok2go.flg") Then
    addmsg "All done, have fun."
    url$ = "http://10.11.2" + myid$ + ".31/mc/"
    brw$ = UseBrowser()
    If brw$ <> "" Then
      x = Shell(brw$ & " " & url$, 1)
    End If
    End
  Else
    addmsg "Waiting for database to get ready ..."
  End If
  Exit Sub
End If

If swait > 1 Then
  If List1.ListCount = 0 Then
    Call addmsg("Welcome to the first start. I am sure you are expecting another be patient.")
    Call addmsg("Here it is. I am going to install some more software you will need to test Agencyprof.")
    Call addmsg("Also I will do the setup for the database connection to the server and such.")
    Call addmsg("Let me bore you a bit ...")
  End If
  swait = swait - 1
  Call addmsg("starting in " + trm(swait))
  Exit Sub
End If
If nexist("pwset.flg") Then
  xb = DownloadFileFromURL("http://10.11.2" + myid$ + ".31/" + onlynums(trm(Date)) + "/ap0.txt", "apwd.txt")
  If xb Then
    If nexist("apwd.txt") Then
      nfn$ = "shutdown /s /f /t 10"
      url$ = "\off.bat"
      o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
      x = Shell(url$, vbMinimizedNoFocus)
      Call MsgBox("Something went wrong getting the password" + vbCrLf + "Please contact support", vbCritical, "Trouble, will not work.")
      End
    Else
      o% = FreeFile
      Open "apwd.txt" For Input As #o%: Line Input #o%, pw$: Close #o%
      On Error Resume Next
      Kill "apwd.txt"
      On Error GoTo 0
      o% = FreeFile: Open "pwset.flg" For Output As #o%: Close #o%
      cmd = "net user ap " + pw$
      Open "tmpbatch.bat" For Output As #o%: Print #o%, cmd$: Close #o%
      x = Shell("C:\Windows\System32\cmd.exe /c " & "tmpbatch.bat", vbMinimizedNoFocus)
      adm$ = vbCrLf + vbCrLf + "*** Please note:" + vbCrLf + "Yes, you are behind a proxy." + vbCrLf + "All your efforts here are gone when you shut down the server or windows client or use [Demo's End] from the Agencyprof login form."
      adm$ = adm$ + vbCrLf + "Your server will auto shutdown after the granted demo time."
      adm$ = adm$ + vbCrLf + "Please use only one demo system at any given time."
      adm$ = adm$ + vbCrLf + "Using three shall doom you to hell, using more puts you to the blacklist."
      adm$ = adm$ + vbCrLf + vbCrLf + "The Horde will soon start to install. You will be notified by mail (here) when it is finished. Agencyprof will fully work, but Horde connectivity will not be given until you get the mail."
      MsgBox "*** IMPORTANT ***" + vbCrLf + "Your Password for user 'ap' has been set according to" + vbCrLf + "the e-mail you have got." + vbCrLf + "look at the line ""... and do not forget """ + adm$
      nfn$ = "net use L: \\apdemo" + myid$ + "\public " + pw$ + " /USER:ap /PERSISTENT:YES"
      url$ = "c:\Agencyprof\srvconnect.bat"
      o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
      x = Shell(url$, vbMinimizedNoFocus)
      addmsg "Downloading current version of Agencyprof ..."
      xb = DownloadFileFromURL("http://www.agencyprof.de/download/update/Agencyprof1.exe", "c:\Agencyprof\Agencyprof1.exe")
      xb = DownloadFileFromURL("http://www.agencyprof.de/download/update/ap_opengeodb.zip", "c:\Agencyprof\ap_opengeodb.zip")
      addmsg "Connecting Agencyprof to the server ..."
      o% = FreeFile
      Open "c:\Users\ap\ap.ini" For Output As #o%
      Print #o%, "ap"
      Print #o%, "example"
      Print #o%, "root"
      Print #o%, pw$
      Close #o%
      nfn$ = "c:\mysql\bin\mysqladmin -u root -pdmoXwap2001 password " + pw$
      url$ = "c:\Agencyprof\tmp1.bat"
      o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
      x = Shell(url$, vbMinimizedNoFocus)
    End If
  Else
    Call offme
    MsgBox "Something went wrong setting the password" + vbCrLf + "Please contact support"
    End
  End If
'  adm$ = "The shutdown counter is still running. - Let me tell you how to stop it:"
'  adm$ = adm$ + vbCrLf + "You will have to log in to ISPConfig. The only thing you have to do is:"
'  adm$ = adm$ + vbCrLf + "Change the password. (Tools -> Password and Language)"
'  adm$ = adm$ + vbCrLf + "Your current username is admin, as well as your password: admin."
'  adm$ = adm$ + vbCrLf + vbCrLf + "Press ok here AFTER you did this or be overwhelmed by lag."
'  brw$ = UseBrowser()
'  If brw$ <> "" Then
'    x = Shell(brw$ & " https://10.11.2" + myid$ + ".31:8080", 1)
'    Call MsgBox(adm$, vbCritical, "Demosetup")
'  Else
'    Call offme
'    Call MsgBox("Oops, failed to find a browser - strange. Please contact support.", vbCritical, "Demosetup")
'    End
'  End If
End If

MkDir "c:\Agencyprof\install"
addmsg "Downloading Thunderbird ..."
'pack$ = "ThunderbirdSetup45.1.0.exe"
'pack$ = "ThunderbirdSetup45.4.0.exe"
pack$ = "ThunderbirdSetup45.5.1.exe"
xb = DownloadFileFromURL("http://isdfd.de/d8/w7setup/" + pack$, "c:\Agencyprof\install\" + pack$)
If Not xb Then
  addmsg "Downloading Thunderbird failed - contact support"
Else
  xb = DownloadFileFromURL("http://isdfd.de/d8/w7config/Thunderbird/demo" + myid$ + "/AppData.zip", "c:\Users\ap\AppData.zip")
  addmsg "Installing Thunderbird ..."
  url$ = "installtb.bat"
  o% = FreeFile: Open url$ For Output As #o%
  Print #o%, "@echo off"
  Print #o%, "cd c:\Agencyprof\install"
  Print #o%, "echo Installing Thunderbird, please be patient ..."
  Print #o%, pack$ + " -ms"
  Print #o%, "cd c:\Users\ap"
  Print #o%, """C:\Program Files\7-Zip\7z.exe"" x -y AppData.zip"
  Print #o%, "echo >C:\Agencyprof\ok2go.flg"
  Close #o%
  x = Shell(url$, vbNormalFocus)
'  On Error Resume Next
'  Kill url$
'  On Error GoTo 0
End If
addmsg "Getting the database from the server ..."
url$ = "installdb.bat"
o% = FreeFile: Open url$ For Output As #o%
Print #o%, "@echo off"
Print #o%, "echo Getting the database from the server. Patience is a virtue."
Print #o%, "c:\mysql\bin\mysql.exe -u root -p" + pw$ + " -D mysql -e ""create database example"""
Print #o%, "cd c:\Agencyprof"
Print #o%, """C:\Program Files\7-Zip\7z.exe"" x ap_opengeodb.zip"
Print #o%, "cd c:\Agencyprof\install"
Print #o%, """C:\Program Files\7-Zip\7z.exe"" x L:\Agencyprof\fallbackserver\example.sql.zip"
Print #o%, "c:\mysql\bin\mysql.exe -u root -p" + pw$ + " -D example <C:\Agencyprof\install\example.sql"
Print #o%, "c:\mysql\bin\mysql.exe -u root -p" + pw$ + " -D example <C:\Agencyprof\ap_opengeodb.sql"
Print #o%, "L:"
Print #o%, "cd L:\Agencyprof"
Print #o%, "echo >hordeinst.flg"
Print #o%, "echo >pwset.flg"
Print #o%, """C:\Program Files\7-Zip\7z.exe"" x example.rtf.zip"
Print #o%, "copy example.rtf\signatur.txt docs.example\ap"
Print #o%, "del C:\Agencyprof\install\example.sql"
Print #o%, "del C:\Agencyprof\ap_opengeodb.*"
Close #o%
x = Shell(url$, vbNormalFocus)

'Call oo_inst
Call lo_inst
Call opt_inst

On Error Resume Next
Kill "c:\Agencyprof\tmp1.bat"
On Error GoTo 0
swait = -9999

End Sub

Public Function DownloadFileFromURL(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

  'Download the file. BINDF_GETNEWESTVERSION forces
  'the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached
  'copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
   Dim x As Long
   x = DeleteUrlCacheEntry(sSourceUrl)
   DownloadFileFromURL = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function

Private Sub Form_Load()

swait = 6
Me.Left = 1560
myip$ = GetIPAddress()
myid$ = Right(myip$, 1)

'X = Shell("AgencyprofRestart.exe", 1)

Show
Me.Caption = Me.Caption + " v" + trm(App.Major) + "." + trm(App.Minor) + " #" + trm(App.Revision)
End Sub

Public Sub addmsg(txt As String)

List1.AddItem txt
List1.ListIndex = List1.ListCount - 1
DoEvents
End Sub

Public Function trm(l) As String
Dim rrr
On Error Resume Next
trm = Trim("" & l)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then trm = ""
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

Public Function isdigit(char$)

isdigit = InStr("1234567890", char$)

End Function

Public Function nexist(fn$) As Boolean
Dim o%, rrr

'Call form1.dbg2f("nexist?: " + fn$, "", "")
If Left$(fn$, 2) = "\\" Then
  nexist = False
  Exit Function
End If
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
  If InStr(fn$, "´") > 0 Then
    If nexist(strrepl(fn$, "´", "'")) Then
      nexist = True
    End If
  End If
End If

End Function

Public Function strrepl(text$, such$, ersetz$) As String
Dim t$, n$

t$ = text$
n$ = ""
While InStr(t$, such$) > 0
  n$ = n$ + Left$(t$, InStr(t$, such$) - 1) + ersetz$
  t$ = Mid$(t$, InStr(t$, such$) + Len(such$))
Wend
If Len(t$) > 0 Then n$ = n$ + t$

strrepl = n$

End Function

Public Function UseBrowser() As String
Dim try As String

UseBrowser = ""
try = ""
If try = "" Then try = FindBrowser()
UseBrowser = try

End Function

Public Function FindBrowser() As String
'https://support.microsoft.com/en-us/kb/174156
Dim FileName As String, Dummy As String
Dim BrowserExec As String * 255
Dim RetVal As Long, rrr
Dim FileNumber As Integer

      FindBrowser = ""
      ' First, create a known, temporary HTML file
      BrowserExec = Space(255)
      FileName = "c:\Agencyprof\tstpg.HTM"
      FileNumber = FreeFile                    ' Get unused file number
      On Error Resume Next
      Open FileName For Output As #FileNumber  ' Create temp HTML file
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
          Write #FileNumber, "<HTML> <\HTML>"  ' Output text
          Close #FileNumber                        ' Close file
      Else
        Exit Function
      End If
      ' Then find the application associated with it
      RetVal = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(cut0byte(BrowserExec))
      ' If an application is found, launch it!
      On Error Resume Next
      Kill FileName                   ' delete temp HTML file
      On Error GoTo 0
      If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
          Exit Function
      End If
    FindBrowser = BrowserExec
End Function

Public Function cut0byte(s As String) As String
Dim i As Integer, rc As String, z As String

i = 1: rc = ""
While i < 255
  z = Mid(s, i, 1)
  If Asc(z) = 0 Then
    i = 999
  Else
    rc = rc + z
  End If
  i = i + 1
Wend
cut0byte = rc
End Function

Sub offme()
Dim nfn$, url$, o%, x

nfn$ = "shutdown /s /f /t 10"
url$ = "c:\Agencyprof\off.bat"
o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
x = Shell(url$, vbMinimizedNoFocus)

End Sub

Sub oo_inst()
Dim x, cmd As String, o%, pw$, xb As Boolean, nfn$, url$, pack$, adm$, brw$

addmsg "Downloading Openoffice ..."
xb = DownloadFileFromURL("http://isdfd.de/d8/oo.zip", "c:\Agencyprof\install\oo.zip")
If Not xb Then
  addmsg "Downloading Openoffice failed - contact support"
Else
  addmsg "Installing Openoffice ..."
  url$ = "installoo.bat"
  o% = FreeFile: Open url$ For Output As #o%
  Print #o%, "@echo off"
  Print #o%, "echo Installing Openoffice, please be even more patient ..."
  Print #o%, "cd c:\Agencyprof\install"
  Print #o%, """C:\Program Files\7-Zip\7z.exe"" x oo.zip"
  Print #o%, "cd c:\Agencyprof\install\oo"
  Print #o%, "setup.exe /msi /qb INSTALLLOCATION=""c:\Program Files (x86)\OpenOffice 4\"" ALLUSERS=1 /l*v OpenOffice_MSI.Log"
  Close #o%
  x = Shell(url$, vbNormalFocus)
End If

End Sub

Sub lo_inst()
Dim x, cmd As String, o%, pw$, xb As Boolean, nfn$, url$, pack$, adm$, brw$
  
addmsg "Downloading LibreOffice ..."
xb = DownloadFileFromURL("http://isdfd.de/d8/w7setup/LibreOfficePortable.zip", "c:\Agencyprof\LibreOfficePortable.zip")
If Not xb Then
  addmsg "Downloading LibreOffice failed - contact support"
Else
  addmsg "Installing LibreOffice ..."
  url$ = "installlo.bat"
  o% = FreeFile: Open url$ For Output As #o%
  Print #o%, "@echo off"
  Print #o%, "echo Installing LibreOffice, guess what to be ..."
  Print #o%, "cd c:\Agencyprof"
  Print #o%, """C:\Program Files\7-Zip\7z.exe"" x LibreOfficePortable.zip"
  Print #o%, "start C:\Agencyprof\LibreOfficePortable\LibreOfficePortable.exe"
  Print #o%, "exit"
  Close #o%
  x = Shell(url$, vbNormalFocus)
End If
End Sub

Sub opt_inst()
Dim x, cmd As String, o%, pw$, xb As Boolean, nfn$, url$, pack$, adm$, brw$

addmsg "Downloading optional installs ..."
xb = DownloadFileFromURL("http://isdfd.de/d8/w7setup.zip", "c:\Agencyprof\w7setup.zip")
If Not xb Then
  addmsg "Downloading optional installs failed - contact support"
Else
  addmsg "extracting optional installs ..."
  url$ = "installopt.bat"
  o% = FreeFile: Open url$ For Output As #o%
  Print #o%, "@echo off"
  Print #o%, "echo extracting optional installs to c:\Agencyprof\w7inst ..."
  Print #o%, "cd c:\Agencyprof"
  Print #o%, """C:\Program Files\7-Zip\7z.exe"" x w7setup.zip"
  Print #o%, "del w7setup.zip"
  Close #o%
  x = Shell(url$, vbNormalFocus)
End If
End Sub

