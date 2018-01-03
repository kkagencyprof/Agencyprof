VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form acclogload 
   Caption         =   "Accesslog download"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   600
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "acclogload.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Abbruch"
      Top             =   2400
      Width           =   375
   End
   Begin VB.ListBox logwin 
      Height          =   1890
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   120
      Top             =   1920
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2175
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "acclogload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ftprdir$, ftphost$, ftpuser$, ftppass$, adrid$, aclog$
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
axsResizer1.SaveControlPositions
Call ftpsetlogwin("acclogload")

Me.Top = form1.mylasttop(Me.Name)
Me.Left = form1.mylastleft(Me.Name)
Show

End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.Name, Me.Top)
Call form1.setmylastleft(Me.Name, Me.Left)
If form1.sockCmd.Connected Then form1.sockCmd.Abort
If form1.sockData.Listening Or form1.sockData.Connected Then form1.sockData.Abort

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me

End Sub
Public Sub initload(adr$, k$)
Dim id$, c$, r As Recordset, x
Dim nError As Integer
Dim strLocalFile As String
Dim strRemoteFile As String
Dim bTransfered As Integer
Static bCanceled As Integer


id$ = adr$
If k$ <> "" Then id$ = adr$ & k$
ftprdir$ = "": ftphost$ = "": ftpuser$ = "": ftppass$ = ""
c$ = "select felddaten from auftritthigru where feldname='Username' and auftrittstyp='accesslog' and auftrittsid='" & id$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  ftpuser$ = trm(r!felddaten)
End If
c$ = "select felddaten from auftritthigru where feldname='Servername' and auftrittstyp='accesslog' and auftrittsid='" & id$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  ftphost$ = trm(r!felddaten)
End If
c$ = "select felddaten from auftritthigru where feldname='Passwort' and auftrittstyp='accesslog' and auftrittsid='" & id$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  ftppass$ = trm(r!felddaten)
End If
c$ = "select felddaten from auftritthigru where feldname='accesslog' and auftrittstyp='accesslog' and auftrittsid='" & id$ & "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
If Not r.EOF Then
  aclog$ = trm(r!felddaten)
End If
If ftphost$ = "" Or ftpuser$ = "" Or ftppass$ = "" Or aclog$ = "" Then
  logwin.AddItem "Die Zugriffsdaten sind nicht vollständig ausgeführt"
  Exit Sub
End If
MousePointer = 11
DoEvents
    If form1.sockCmd.Connected Then
        nError = form1.sockCmd.Disconnect
        If (nError <> 0) Then
            MsgBox "Unable to disconnect.  Error " & nError
        End If
    Else
        Dim strHostName As String
        Dim strDirectory As String
        Dim nRemotePort As Integer
        
'        logwin.Clear

        strHostName = ftphost$
        nRemotePort = 21
        
        '
        ' Establish the connection with the FTP server
        '
        If Not FtpConnect(strHostName, nRemotePort) Then
            logwin.AddItem "Verbindung mit " & strHostName & " fehlgeschlagen"
            logwin.ListIndex = logwin.ListCount - 1
            GoTo errout
        End If
        
        '
        ' Login to the server using the supplied username
        ' and password
        '
        If Not FtpLogin(ftpuser$, ftppass$) Then
            form1.sockCmd.Action = SOCKET_DISCONNECT
            logwin.AddItem "Benutzername / Passwort ungültig"
            logwin.ListIndex = logwin.ListCount - 1
            GoTo errout
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
            GoTo errout
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
        GoTo errout
    End If

    strLocalFile = form1.mydatadir() & "\access_log.tmp"
    strRemoteFile = aclog$

    On Error Resume Next
    Kill strLocalFile
    On Error GoTo 0
    bTransfered = FtpGetFile(strLocalFile, strRemoteFile)
    If Not bTransfered Then
        logwin.AddItem "Übertragung fehlgeschlagen"
        logwin.ListIndex = logwin.ListCount - 1
        DoEvents
        GoTo errout
    End If
  DoEvents
  x = Shell("notepad.exe " & strLocalFile, 1)
errout:

MousePointer = 0

End Sub
