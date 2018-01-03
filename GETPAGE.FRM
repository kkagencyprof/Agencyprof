VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form getpg 
   Caption         =   "Text einer Internetseite laden"
   ClientHeight    =   6330
   ClientLeft      =   3555
   ClientTop       =   2835
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "getpg"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6330
   ScaleWidth      =   8205
   Begin SocketWrenchCtrl.Socket sockClient 
      Left            =   4320
      Top             =   600
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
   Begin VB.ListBox llist1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Picture         =   "GETPAGE.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browserpage"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox editView 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   480
      Width           =   3555
   End
   Begin VB.TextBox editAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   3720
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label labelStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4380
      Width           =   7455
   End
   Begin VB.Label labelAddress 
      Caption         =   "&Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   795
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu menuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "getpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Catalyst SocketWrench 3.6
' Copyright 1995-2002, Catalyst Development Corporation
' All rights reserved
'
' This product is licensed to you pursuant to the terms of the
' Catalyst license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'
Option Explicit

Dim g_strHostName As String
Dim g_nRemotePort As Integer
Dim g_strResource As String
Dim g_strFileName As String
Dim g_nLineCount As Integer
Dim g_strLine() As String
Dim g_nTotalBytes As Long

Private Sub cmdAction_Click()
    Dim strBuffer As String
    Dim cchBuffer As Integer
    Dim strHeaderBuffer As String
    Dim cchHeaderBuffer As Integer
    Dim strHeader() As String
    Dim cchHeader As Integer
    Dim nHeader As Integer
    Dim nHeaderCount As Integer
    Dim strHeaderField As String
    Dim strHeaderValue As String
    Dim nResultCode As Integer
    Dim nLine As Integer
    Dim nPos As Integer
    Dim nCopied As Long
    Dim bInHeader As Integer
    Static bCanceled As Integer
    Dim pg$, i%, url$, ln$, p%

    If cmdAction.Caption = "Cancel" Then
        bCanceled = True
        cmdAction.Enabled = False
        If sockClient.State > SOCKET_IDLE Then sockClient.Cancel
        Exit Sub
    End If

    UpdateForm False
    
Restart:

    ReDim g_strLine(0)
    g_nLineCount = 0
    g_nTotalBytes = 0
    
    '
    ' Parse the URL that was entered by the user
    '
    If ParseAddress(editAddress.Text, g_strHostName, g_nRemotePort, g_strResource, g_strFileName) = False Then
        ShowError "ungültige Adresse"
        UpdateForm True
        Exit Sub
    End If

    cmdAction.Caption = "Cancel"
    labelStatus.Caption = "": editView.Text = ""
    
    '
    ' Update the editAddress control with the full URL
    ' for the specified resource
    '
    If g_nRemotePort <> 80 Then
        editAddress.Text = "http://" & g_strHostName & ":" & g_nRemotePort & g_strResource
    Else
        editAddress.Text = "http://" & g_strHostName & g_strResource
    End If


    '
    ' Create the request header that will be sent to the server;
    ' strHeader(1) is the actual command (ie: GET) and the remaining
    ' values are the actual header fields
    '
    ' By default we simply send the 'Host' and 'Accept' header fields
    ' in the request headers. Additional headers may be added by
    ' increasing nHeaderCount and modifying the strHeader array. Make
    ' sure that you end each header with a CRLF pair and that the
    ' last value is a single CRLF pair (the empty line is what tells
    ' the server that it has read all of the request header).
    '
    nHeaderCount = 4
    ReDim strHeader(nHeaderCount)
    strHeader(1) = "GET " & g_strResource & " HTTP/1.0" & Chr(13) & Chr(10)
    strHeader(2) = "Host: " & g_strHostName & Chr(13) & Chr(10)
    strHeader(3) = "Accept: text/*" & Chr(13) & Chr(10)
    strHeader(4) = Chr(13) & Chr(10)

    labelStatus.Caption = "Establishing connection to " & g_strHostName & "..."
    DoEvents

    '
    ' Create the socket, setting the HostName and RemotePort properties
    ' to the values specified in the URL (note that if no port number
    ' was explicitly given, the ParseAddress function will return the
    ' default port number of 80).
    '
    sockClient.AutoResolve = False
    sockClient.Blocking = True
    sockClient.HostName = g_strHostName
    sockClient.RemotePort = g_nRemotePort
    sockClient.Timeout = 60000 ' 60 second timeout

    '
    ' Attempt the connection to the server
    '
    If sockClient.Connect() <> 0 Then
        If bCanceled Then
            bCanceled = False
            labelStatus.Caption = "Connection canceled"
        Else
            ShowError "Unable to connect to " & g_strHostName & " (error " & sockClient.LastError & ")"
        End If
        UpdateForm True
        Exit Sub
    End If

    labelStatus.Caption = "Requesting " & g_strFileName & " from server..."

    '
    ' Send the request header block to the server, which consists of
    ' the command (GET) and the other header fields defined above
    '
    For nHeader = 1 To nHeaderCount
        '
        ' If the number of bytes written doesn't match the length
        ' of the header string, then something has gone wrong;
        ' for a blocking socket, these values should be the same
        '
        cchHeader = Len(strHeader(nHeader))
        If sockClient.Write(strHeader(nHeader), cchHeader) <> cchHeader Then
            If bCanceled Then
                bCanceled = False
                labelStatus.Caption = "Download canceled"
            Else
                ShowError "Unable to send request to " & g_strHostName
            End If
            sockClient.Disconnect
            UpdateForm True
            Exit Sub
        End If
    Next

    labelStatus.Caption = "Downloading " & g_strFileName & "..."

    '
    ' The server will reply with a response header block, followed
    ' by the data for the requested resource; we will re-use the
    ' strHeader array to contain the response header values, and
    ' store each line of the resource in the g_strLine array
    '
    ' Note that this sample, as written, expects that only textual
    ' data (such as HTML pages) will be returned by the server
    '
    ReDim strHeader(0)
    
    nHeaderCount = 0
    cchHeaderBuffer = 0
    cchBuffer = 0
    bInHeader = True

    Do
        If bCanceled Then
            sockClient.Disconnect
            labelStatus.Caption = "Download canceled"
            g_nLineCount = 0
            bCanceled = False
            UpdateForm True
            Exit Sub
        End If

        '
        ' Read the data from the socket, and store it in strBuffer;
        ' the actual number of bytes read is stored in cchBuffer
        '
        cchBuffer = sockClient.Read(strBuffer, 2048)
        If cchBuffer = 0 Then
            '
            ' The server has closed the connection and we have
            ' reached the end of the data stream
            '
            Exit Do
        ElseIf cchBuffer = -1 Then
            '
            ' An error has occurred while reading data from the
            ' server; this should be considered a fatal error
            '
            If bCanceled Then
                bCanceled = False
                labelStatus.Caption = "Download canceled"
            Else
                ShowError "Unable to receive data from server (error " & sockClient.LastError & ")"
            End If
            sockClient.Disconnect
            UpdateForm True
            Exit Sub
        End If

        If bInHeader Then
            '
            ' If we are processing the response header block, then
            ' store the data into the header buffer
            '
            strHeaderBuffer = strHeaderBuffer + strBuffer
            cchHeaderBuffer = cchHeaderBuffer + cchBuffer

            '
            ' Look for the end of the header block, which is a
            ' blank line (a pair of CRLF sequences)
            '
            nPos = InStr(strHeaderBuffer, Chr(13) & Chr(10) & Chr(13) & Chr(10))
            If nPos > 0 Then
                '
                ' The end of the header block has been reached; the
                ' entire response header is stored in strHeaderBuffer
                ' and the remaining data is left in strBuffer to be
                ' processed later
                '
                cchBuffer = cchBuffer - (nPos + 3)
                strBuffer = Right(strHeaderBuffer, cchBuffer)
                strHeaderBuffer = Left(strHeaderBuffer, nPos + 1)
                bInHeader = False

                '
                ' Break strHeaderBuffer apart, with each response
                ' header field being placed into the strHeader array;
                ' this will make it simple to search for specific
                ' header values, etc.
                '
                Do
                    nPos = InStr(strHeaderBuffer, Chr(10))
                    If nPos = 0 Then
                        Exit Do
                    Else
                        nHeaderCount = nHeaderCount + 1
                        ReDim Preserve strHeader(nHeaderCount)
                        strHeader(nHeaderCount) = Trim(Left(strHeaderBuffer, nPos - 2))
                        strHeaderBuffer = Right(strHeaderBuffer, Len(strHeaderBuffer) - nPos)
                    End If
                Loop

                '
                ' Note that strHeader(1) will contain the command
                ' response from the server, and will typically look
                ' something like this:
                '
                '       HTTP/1.0 200 OK
                '
                ' The first part contains the protocol version (in
                ' this case 1.0), the second is the result code
                ' and what follows is a textual description of the
                ' result. A result code in the range of 200-299
                ' indicates success; for a complete description of
                ' the result codes, refer to RFC 2616
                '
                nPos = InStr(strHeader(1), " ")
                If nPos > 0 Then
                    nResultCode = Val(Right(strHeader(1), Len(strHeader(1)) - nPos))
                End If

                If nResultCode >= 300 And nResultCode <= 303 Then
                    '
                    ' A result code in this range indicates that the
                    ' resource has been moved; the new location is
                    ' specified in the Location header field
                    '
                    For nHeader = 2 To nHeaderCount
                        nPos = InStr(strHeader(nHeader), ":")
                        If nPos > 0 Then
                            strHeaderField = UCase(Left(strHeader(nHeader), nPos - 1))
                            If strHeaderField = "LOCATION" Then
                                editAddress.Text = Trim(Right(strHeader(nHeader), Len(strHeader(nHeader)) - nPos))
                                sockClient.Disconnect
                                GoTo Restart
                            End If
                        End If
                    Next
                End If

                If nResultCode < 200 Or nResultCode > 299 Then
                    sockClient.Disconnect
                    ShowError "Result code " & nResultCode & " was returned by the server"
                    UpdateForm True
                    Exit Sub
                End If

                '
                ' Determine the content type of the data being returned
                '
                For nHeader = 2 To nHeaderCount
                    nPos = InStr(strHeader(nHeader), ":")
                    If nPos > 0 Then
                        strHeaderField = UCase(Left(strHeader(nHeader), nPos - 1))
                        
                        If strHeaderField = "CONTENT-TYPE" Then
                            '
                            ' If the content type is not textual, then disconnect
                            ' and warn the user that we cannot display it
                            '
                            strHeaderValue = Trim(Right(strHeader(nHeader), Len(strHeader(nHeader)) - nPos))
                            If Left(strHeaderValue, 5) <> "text/" Then
                                sockClient.Disconnect
                                ShowError "The specified resource is not text, and cannot be displayed"
                                UpdateForm True
                                Exit Sub
                            End If
                            Exit For
                        End If
                    End If
                Next
                '
                ' Any additional checks for specific header field values
                ' could be placed here
                '
            End If
        End If
        '
        ' If we are not processing the header block, the data into
        ' individual lines to make it easier to process; this will
        ' also handle the different end-of-line character sequences
        ' used by UNIX and Windows servers
        '
        If Not bInHeader Then
            g_nTotalBytes = g_nTotalBytes + cchBuffer
            If g_nLineCount = 0 Then
                g_nLineCount = 1
                ReDim Preserve g_strLine(g_nLineCount)
            End If
            '
            ' If the buffer contains carriage returns, then strip
            ' them out and use only linefeeds to mark the end-of-line
            '
            Do
                nPos = InStr(strBuffer, Chr(13))
                If nPos = 0 Then Exit Do
                strBuffer = Left(strBuffer, nPos - 1) & Right(strBuffer, cchBuffer - nPos)
                cchBuffer = cchBuffer - 1
            Loop
            Do
                nPos = InStr(strBuffer, Chr(10))
                If nPos = 0 Then
                    Exit Do
                Else
                    '
                    ' If the linefeed is at the beginning of the line, then
                    ' simply append CRLF; otherwise append the remaining
                    ' characters and then CRLF.
                    '
                    If nPos = 1 Then
                        g_strLine(g_nLineCount) = g_strLine(g_nLineCount) & Chr(13) & Chr(10)
                    Else
                        g_strLine(g_nLineCount) = g_strLine(g_nLineCount) & Left(strBuffer, nPos - 1) & Chr(13) & Chr(10)
                    End If
                    g_nLineCount = g_nLineCount + 1
                    ReDim Preserve g_strLine(g_nLineCount)
                    cchBuffer = cchBuffer - nPos
                    strBuffer = Right(strBuffer, cchBuffer)
                End If
            Loop
            If cchBuffer > 0 Then
                g_strLine(g_nLineCount) = g_strLine(g_nLineCount) + strBuffer
            End If
        End If
        labelStatus.Caption = "Downloaded " & g_nTotalBytes & " bytes of " & g_strFileName
        DoEvents
    Loop

    '
    ' All of the data base been read, so disconnect from the
    ' server and report the total number of bytes read
    '
    sockClient.Disconnect
    labelStatus.Caption = "Downloaded " & g_nTotalBytes & " bytes of " & g_strFileName

    '
    ' Update the editView control to contain the text returned;
    ' note that we check the total number of bytes being copied
    ' to the edit control because the control is limited to a
    ' maximum of about 32K
    '
    editView.Visible = False
    nCopied = 0

    For nLine = 1 To g_nLineCount
        nCopied = nCopied + Len(g_strLine(nLine))
        If nCopied > 30000 Then
            ShowError "Unable to display all data returned by the server"
            Exit For
        End If
        editView.Text = editView.Text & g_strLine(nLine)
    Next
    
    editView.Visible = True
    pg$ = strrepl(editView.Text, Chr$(10), "")
    pg$ = strrepl(pg$, """", "")
    pg$ = LCase(strrepl(pg$, Chr$(13), ""))
    i% = InStr(pg$, "<a href=")
    While i% > 0
      url$ = Mid$(pg$, i%)
      p% = InStr(url$, "="): url$ = Mid$(url$, p% + 1)
      p% = InStr(url$, ">"):
      ln$ = Mid$(url$, p% + 1)
      url$ = Left$(url$, p% - 1)
      p% = InStr(ln$, "</a")
      ln$ = Left$(ln$, p% - 1)
      'llist1.AddItem ln$ & ":" & url$
      llist1.AddItem url$
      pg$ = Mid$(pg$, i% + 1)
      i% = InStr(pg$, "<a href=")
    Wend
    '
    ' Update the form controls
    '
    UpdateForm True

End Sub

Private Sub Command1_Click()
editAddress.Text = frmBrowser.brwWebBrowser.LocationURL
Call cmdAction_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub editAddress_GotFocus()
    editAddress.SelStart = 0
    editAddress.SelLength = Len(editAddress.Text)
End Sub

Private Sub editView_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 3 Then KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 And cmdAction.Caption <> "Cancel") Or (KeyAscii = 27 And cmdAction.Caption = "Cancel") Then
        KeyAscii = 0
        cmdAction_Click
    End If
End Sub

Private Sub Form_Load()
Me.Top = form1.mylasttop(Me.Name)
Me.Left = form1.mylastleft(Me.Name)

Show
    UpdateForm True
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        '
        ' The minimum height and width for the form
        '
        If Width < 3000 Then Width = 3000
        If Height < 2500 Then Height = 2500

        '
        ' Adjust the controls based on the new width and/or
        ' height of the form
        '
        llist1.Top = ScaleHeight - llist1.Height
        llist1.Width = ScaleWidth - llist1.Left - 250
        Command2.Top = llist1.Top
        editAddress.Width = cmdAction.Left - editAddress.Left - 250
        Command1.Left = ScaleWidth - Command1.Width
        cmdAction.Left = Command1.Left - 1000
        labelStatus.Width = ScaleWidth
        labelStatus.Top = Command2.Top - labelStatus.Height
        editView.Width = ScaleWidth
        editView.Height = labelStatus.Top - editView.Top - labelStatus.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call form1.setmylasttop(Me.Name, Me.Top)
Call form1.setmylastleft(Me.Name, Me.Left)
Hide
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

Private Sub menuSaveAs_Click()
    Dim hFile As Integer
    Dim strFileName As String
    Dim nLine As Integer

    strFileName = g_strFileName
    If strFileName = "/" Then strFileName = "index.htm"

    cmnDialog.CancelError = True
    cmnDialog.DefaultExt = "htm"
    cmnDialog.filename = strFileName
    cmnDialog.Filter = "Web Page (*.htm;*.html)|*.htm;*.html|Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cmnDialog.FilterIndex = 1
    cmnDialog.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
    
    On Error GoTo Canceled
    cmnDialog.Action = 2 ' Save As

    On Error GoTo 0
    strFileName = cmnDialog.filename

    hFile = FreeFile
    Open strFileName For Binary Access Write As #hFile
    For nLine = 1 To g_nLineCount
        Put #hFile, , g_strLine(nLine)
    Next
    Close #hFile
    Exit Sub

Canceled:
    Exit Sub

End Sub

Private Function ParseAddress(ByVal strAddress As String, strHostName As String, nRemotePort As Integer, strResource As String, strFileName As String) As Integer
    Dim nPos As Integer
    Static strCurrentHost As String

    ParseAddress = False

    '
    ' Trim the address and set the default port number to 80
    '
    strAddress = Trim(strAddress)
    nRemotePort = 80

    '
    ' Check if the address specifies the protocol and remove
    ' it from the string if necessary; note that only the HTTP
    ' protocol is supported.
    '
    nPos = InStr(strAddress, "://")
    If nPos > 0 Then
        Dim strProtocol As String
        
        strProtocol = Left(strAddress, nPos - 1)
        strAddress = Right(strAddress, Len(strAddress) - (nPos + 2))
        
        If LCase(strProtocol) <> "http" Then
            Exit Function
        End If
    End If

    '
    ' Look for the path delimiter (forward slash) in the address;
    ' this is what seperates the hostname from the path. If no
    ' delimiter is found, then the address just contains a hostname
    '
    nPos = InStr(strAddress, "/")
    If nPos = 0 Then
        strHostName = strAddress
        strResource = "/"
    Else
        strHostName = Left(strAddress, nPos - 1)
        strResource = Right(strAddress, Len(strAddress) - (nPos - 1))
    End If
    
    '
    ' Look for the port specification in the hostname; if it exists,
    ' this value overrides the default port of 80
    '
    nPos = InStr(strHostName, ":")
    If nPos > 0 Then
        nRemotePort = Val(Right(strHostName, Len(strHostName) - nPos))
        strHostName = Left(strHostName, nPos - 1)
        If nRemotePort = 0 Then Exit Function
    End If

    '
    ' Determine the base file name for the resource; this is strictly
    ' for display purposes
    '
    strFileName = strResource
    Do
        nPos = InStr(strFileName, "/")
        If nPos = 0 Then
            nPos = InStr(strFileName, "?")
            If nPos > 0 Then strFileName = Left(strFileName, nPos - 1)
            Exit Do
        Else
            strFileName = Right(strFileName, Len(strFileName) - nPos)
        End If
    Loop
    If Len(strFileName) = 0 Then strFileName = "/"

    If Len(strHostName) > 0 Then
        strCurrentHost = strHostName
    Else
        '
        ' If no hostname has been specified, then this is a
        ' relative address (ie: /index.html); use the current
        ' hostname if it exists, otherwise return an error
        '
        If Len(strCurrentHost) = 0 Then Exit Function
        strHostName = strCurrentHost
    End If

    ParseAddress = True

End Function

Private Sub ShowError(ByVal strError As String)
    labelStatus.Caption = strError
    MsgBox strError, vbExclamation, "SocketWrench"
End Sub

Private Sub sockClient_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    Debug.Print "Error " & ErrorCode & " (" & ErrorString & ")"
End Sub

Private Sub UpdateForm(ByVal bEnable As Integer)
    If bEnable Then
        '
        ' The resource has been downloaded or the operation
        ' has been canceled
        '
        cmdAction.Enabled = True
        cmdAction.Cancel = False
        cmdAction.Caption = "Get"
        menuExit.Enabled = True
        
        '
        ' If g_nLineCount is greater than zero, then this means
        ' that data has been read from the server
        '
        If g_nLineCount > 0 Then
            editView.Enabled = True
            menuSaveAs.Enabled = True
        Else
            editView.Enabled = False
            menuSaveAs.Enabled = False
        End If
    Else
        '
        ' The resource is currently being downloaded, so
        ' disable the editView control and Save As menu
        ' option
        '
        editView.Enabled = False
        menuSaveAs.Enabled = False
        menuExit.Enabled = False
    End If
End Sub

