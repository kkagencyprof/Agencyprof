VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form ps2pdf 
   Caption         =   "PDF-Konverter"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form2"
   ScaleHeight     =   2580
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   2160
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ps2pdf.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "per Email an Agencyprof"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ListBox logwin 
      Height          =   1995
      IntegralHeight  =   0   'False
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1680
      Picture         =   "ps2pdf.frx":00B2
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ps2pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nupd%, ftpuser$, ftphost$, ftpport$, ftppass$
Dim uId$, psdir$, trrrcnt%, nfl As Long
Dim aKey() As Byte


Sub rlist1()
Dim tr$

List1.Clear
tr$ = Dir(psdir$ & "\*.ps")
While tr$ <> ""
  List1.AddItem tr$
  tr$ = Dir
Wend
tr$ = Dir(psdir$ & "\*.prn")
While tr$ <> ""
  List1.AddItem tr$
  tr$ = Dir
Wend
End Sub
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Dim nError As Integer
Dim strLocalFile As String
Dim strRemoteFile As String
Dim bTransfered As Integer, i%

Static bCanceled As Integer, ftprdir$
Dim o%, l$, l1$, p%, z$, vac$, enc$, ftpp$, gem$


If List1.ListCount = 0 Then
  logwin.AddItem "Keine Dateien zu übertragen"
  logwin.ListIndex = logwin.ListCount - 1
  Exit Sub
End If
gem$ = Me.Caption
trrrcnt% = -1

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
        ftpp$ = decrypt(ftppass$, "hihallohuhu4716")
        If Not FtpLogin(Trim$(ftpuser$), ftpp$) Then
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
trrrcnt% = 0
    For i% = 0 To List1.ListCount - 1
      List1.ListIndex = i%
      DoEvents
      strLocalFile = psdir$ & "\" & List1.List(i%)
      strRemoteFile = "ps2pdf_" & List1.List(i%)
      nfl = FileLen(strLocalFile)
      Timer1.Interval = 1000
      Timer1.Enabled = True
      bTransfered = FtpPutFile(strLocalFile, strRemoteFile)
      Timer1.Enabled = False: Me.Caption = gem$
      If Not bTransfered Then
        logwin.AddItem "Übertragung abgebrochen: " & List1.List(i%)
        trrrcnt% = trrrcnt% + 1
      Else
        logwin.AddItem "Übertragung ok: " & List1.List(i%)
        On Error Resume Next
        Kill strLocalFile
        On Error GoTo 0
      End If
      logwin.ListIndex = logwin.ListCount - 1
    Next i%
    Call rlist1
    If trrrcnt% = 0 Then Call Command1_Click
End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset, c$, rrr
axsResizer1.SaveControlPositions

If IsConnected() = False Then Exit Sub
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Call ftpsetlogwin("ps2pdf")


uId$ = form1.getuserid()
psdir$ = form1.getusersetting("pspath")
Show
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM poplist where id='" + uId$ + "_PDFServer'", form1.adoc, dbOpenDynaset, dbReadOnly)
If r.EOF Then
  logwin.AddItem "PDFServer nicht konfiguriert"
Else
  ftpuser$ = r!user
  ftphost$ = r!server
  ftpport$ = r!port
  ftppass$ = r!psswd
  Call rlist1
  DoEvents
  If List1.ListCount > 0 Then Call Command2_Click
End If
End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
If form1.sockCmd.Connected Then form1.sockCmd.Abort
If form1.sockData.Listening Or form1.sockData.Connected Then form1.sockData.Abort
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

Private Sub Timer1_Timer()

Me.Caption = trm(Int((ftp_bytes_sent_this_file / nfl) * 100)) & "% gesendet (" & trm(Int(ftp_bytes_sent_this_file / 1024)) & " kB)"

End Sub
