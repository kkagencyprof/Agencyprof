VERSION 5.00
Begin VB.Form AutoAnwahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anwahl"
   ClientHeight    =   1050
   ClientLeft      =   3765
   ClientTop       =   3720
   ClientWidth     =   3465
   LinkTopic       =   "AutoAnwahl"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   1050
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "Dialer.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "&Anwahl"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox nummer 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nummer:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "AutoAnwahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Public Sub cmdDial_Click()
    Dim buff As String
    Dim nResult As Long, o%, url$
    Dim amt$, brw$, X, dp$

    url$ = form1.getusersetting("fb7050url", "")
    
    If url$ = "callto:" Or url$ = "tsip:" Then
      o% = FreeFile
      amt$ = form1.s0dir() + "\" + form1.docs() + "\" & form1.getuserid() & "\dialer.htm"
      Open amt$ For Output As #o%
      Print #o%, "<head>"
      dp$ = form1.getusersetting("fb7050dialprefix", "")
      Print #o%, "<meta http-equiv='refresh' content='0; URL=" + url$ + dp$ + trm(nummer) + "'>"
      Print #o%, "</head><body>"
      Print #o%, "Dialing: <a href='" + url$ + dp$ + trm(nummer) + "'>" + dp$ + trm(nummer) + "</a><br>"
      Print #o%, "</body>"
      Close #o%
      Call Command3_Click
      Unload frmBrowser
      DoEvents
      brw$ = form1.UseBrowser()
      If brw$ <> "" Then
        X = Shell(brw$ & " file:///" + strrepl(amt$, "\", "/"), 1)
      Else
        frmBrowser.StartingAddress = "file:///" + strrepl(amt$, "\", "/")
        Load frmBrowser
        frmBrowser.locdialer = amt$
        frmBrowser.cboAddress.Visible = False
        frmBrowser.lblAddress.Visible = False
      End If
      DoEvents
      Exit Sub
    End If
    If url$ <> "" Then
      o% = FreeFile
      amt$ = form1.getusersetting("fb7050dialer", form1.s0dir() + "\" + form1.docs() + "\" & form1.getuserid() & "\dialer.htm")
      Open amt$ For Output As #o%
      Print #o%, "<form method=""POST"" action=""" + url$ + """ target=""_self"" id=""uiPostForm"" name=""uiPostForm"">"
      Print #o%, "<input type=""hidden"" name=""login:command/password"" value=""" + form1.getusersetting("fb7050pass", "") + """ id=""uiPostPassword"">"
      Print #o%, "<input type=""hidden"" name=""telcfg:settings/UseClickToDial"" value=""1"" id=""uiPostClickToDial"">"
      Print #o%, "<input type=""hidden"" name=""telcfg:settings/DialPort"" value=""" + form1.getusersetting("fb7050dialport", "50") + """ id=""uiPostDialPort"">"
      Print #o%, "<input type=""hidden"" name=""getpage"" value="""" id=""uiPostGetPage"">"
      Print #o%, "<table border=0>"
      dp$ = form1.getusersetting("fb7050dialprefix", "")
      If Left(nummer, 1) = "~" Then
        dp$ = ""
        nummer = Mid(nummer, 2)
      End If
      Print #o%, "<tr><td>Nummer:</td><td><input name=""telcfg:command/Dial"" value=""" + dp$ + trm(nummer) + "#"" size=40 id=""uiPostDial""></td></tr>"
      Print #o%, "<tr><td></td><td><input type=""submit"" value=""" + transe("Nummer wählen") + """></td></tr>"
      Print #o%, "</table></form>"
      Close #o%
      Call Command3_Click
      Unload frmBrowser
      DoEvents
      brw$ = form1.UseBrowser()
      If brw$ <> "" Then
        X = Shell(brw$ & " file:///" + strrepl(amt$, "\", "/"), 1)
      Else
        frmBrowser.StartingAddress = "file:///" + strrepl(amt$, "\", "/")
        Load frmBrowser
        frmBrowser.locdialer = amt$
        frmBrowser.cboAddress.Visible = False
        frmBrowser.lblAddress.Visible = False
      End If
      DoEvents
    Else
      'Invoke tapiRequestMakeCall. If tapiRequestMakeCall returns 0, the
      'request has been accepted. It is up to the call manager application
      'to do any further work. The second-to-last argument should be
      'changed to be the name of the person you are dialing.
      amt$ = trm(form1.getusersetting("Amtsanholung", ""))
      nResult = tapiRequestMakeCall&(amt$ & Trim$(nummer), CStr(Caption), "Wählen...", "")
      'Display message if error
      If nResult <> 0 Then
        buff = "Fehler bei der Anwahl : "
        Select Case nResult
            Case TAPIERR_NOREQUESTRECIPIENT
                buff = buff & form1.inmylanguage("Windows TAPI läuft nicht und kann nicht gestartet werden.")
            Case TAPIERR_REQUESTQUEUEFULL
                buff = buff & form1.inmylanguage("Die Wählwarteschlange ist voll.")
            Case TAPIERR_INVALDESTADDRESS
                buff = buff & form1.inmylanguage("Die Telefonnummer ist ungültig.")
            Case Else
                buff = buff & form1.inmylanguage("Unbekannter Fehler.")
        End Select
        MsgBox buff
      End If
    End If
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
AutoAnwahl.Caption = transe("Anwahlfenster")
cmdDial.Caption = transe("&Anwahl")
Label1.Caption = transe("Nummer:")
Command3.ToolTipText = transe("Formular schliessen")
Show
EnableDial
cmdDial.Enabled = True
End Sub


Private Sub txtNumber_Change()
    EnableDial
End Sub

Private Sub EnableDial()
    cmdDial.Enabled = Len(Trim$(nummer)) > 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0
End Sub

