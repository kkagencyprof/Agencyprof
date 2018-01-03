VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Anmeldung - AgencyProf"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7380
   Begin VB.Timer Timer2 
      Interval        =   59000
      Left            =   0
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "demo's end"
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      ToolTipText     =   "Server and client will be nuked asap."
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox dbdrv 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   34
      ToolTipText     =   "OBDC-Datenquelle eingeben"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "login.frx":0CCA
      Left            =   4800
      List            =   "login.frx":0CCC
      Sorted          =   -1  'True
      TabIndex        =   32
      ToolTipText     =   "Sprache wählen"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CheckBox pingb4 
      Alignment       =   1  'Rechts ausgerichtet
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      ToolTipText     =   "sofort anmelden"
      Top             =   1920
      Width           =   255
   End
   Begin VB.ComboBox langu 
      Height          =   315
      ItemData        =   "login.frx":0CCE
      Left            =   4800
      List            =   "login.frx":0CD0
      Sorted          =   -1  'True
      TabIndex        =   26
      Text            =   "de"
      ToolTipText     =   "Sprache wählen"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   6720
      Picture         =   "login.frx":0CD2
      Style           =   1  'Grafisch
      TabIndex        =   25
      ToolTipText     =   "Liste löschen"
      Top             =   1875
      Width           =   375
   End
   Begin VB.ComboBox dbserver 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   24
      Text            =   "localhost"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   6720
      Picture         =   "login.frx":11C2
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "Liste löschen"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   6720
      Picture         =   "login.frx":16B2
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Liste löschen"
      Top             =   240
      Width           =   375
   End
   Begin VB.ComboBox user 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox dbname 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   20
      Text            =   "apdemo.mdb"
      Top             =   780
      Width           =   1815
   End
   Begin VB.CommandButton Command19 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CheckBox autoenter 
      Alignment       =   1  'Rechts ausgerichtet
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      ToolTipText     =   "sofort anmelden"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   2760
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      Picture         =   "login.frx":1BA2
      Style           =   1  'Grafisch
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox dbdsn 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "OBDC-Datenquelle eingeben"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox dbpsswd 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   8
      ToolTipText     =   "Bitte geben Sie Ihr Passwort ein"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox dbuid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Text            =   "root"
      ToolTipText     =   "Datenbank-Verzeichnis"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      Picture         =   "login.frx":1DF2
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "AgencyProf starten"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   120
      Picture         =   "login.frx":2043
      ScaleHeight     =   1755
      ScaleWidth      =   2715
      TabIndex        =   10
      ToolTipText     =   "Releasenotes"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Driver"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      ToolTipText     =   "Hier klicken, um Eingabe zu ermöglichen"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Anmeldep&rofil"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "&Opts"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   30
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "&Ping"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label autoenterlabel 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "&Auto-Login"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Server"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      ToolTipText     =   "Hier klicken, um Eingabe zu ermöglichen"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "DSN"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      ToolTipText     =   "Hier klicken, um Eingabe zu ermöglichen"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Contains cryptography software by David Ireland of DI Management Services Pty Ltd <www.di-mgt.com.au>."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.Agencyprof.de"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: kk@agencyprof.de"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ihre Benutzer-Kennung"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Bitte Ihre Kennung eingeben"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "(oder neue Kennung wählen)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Datenbank"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "Hier klicken, um Eintrag zu ermöglichen"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Passwort"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Hier klicken, um Passwort-Eingabe zu ermöglichen"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Datenbank-UID"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Hier klicken, um Eingabe zu ermöglichen"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2007,2008  Karsten Kaus"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3735
      Left            =   3120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   3015
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nodo%, t1cnt, autoclosecount As Integer
Dim uuid$, dbpara$, dbdriver$, hppth$, currlang
Public login_exv%

Private Sub autoenter_Click()
Dim o As Integer

If autoenter.value = 0 Then
  o = FreeFile
  Open hppth$ & "\ap.ini" For Output As #o%
  Print #o, uuid$
  Print #o, dbname.text
  Print #o, dbuid.text
  Print #o, dbpsswd.text
  Print #o, dbdsn.text
  Print #o, dbserver.text
  Print #o, 0
  Print #o, langu.text
  Print #o, trm(pingb4.value)
  Close #o
End If
End Sub

Private Sub autoenterlabel_Click()
If autoenter.value = 1 Then
  autoenter.value = 0
Else
  autoenter.value = 1
End If

End Sub

Private Sub Combo1_Click()
Dim fn$, o%, enc$, rrr

fn$ = hppth$ + "\" + trm(Combo1.text) + ".app"
If Not nexist(fn$) Then
  enc$ = genc()
  o% = FreeFile
  Open fn$ For Input As #o%
  Line Input #o%, fn$: user.text = fn$
  Line Input #o%, fn$: dbname.text = fn$
  Line Input #o%, fn$: dbdsn.text = fn$
  Line Input #o%, fn$: dbserver.text = fn$
  Line Input #o%, fn$: dbuid.text = fn$
  Line Input #o%, fn$: dbpsswd.text = decrypt(fn$, enc$)
  Line Input #o%, fn$: langu.text = fn$
  Line Input #o%, fn$: pingb4.value = Val(fn$)
  Line Input #o%, fn$: autoenter.value = Val(fn$)
  On Error Resume Next
  Line Input #o%, fn$
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then fn$ = "MySQL ODBC 3.51 Driver"
  dbdrv.text = fn$
  Close #o%
Else
  MsgBox ("Das Profil " + fn$ + " ist nicht lesbar.")
End If

End Sub

Private Sub Combo1_DropDown()
Dim tr, fn$

Combo1.Clear
tr = Dir(hppth$ + "\*.app")
While tr <> ""
  fn$ = Left(tr, Len(tr) - 4)
  Combo1.AddItem fn$
  tr = Dir
Wend

End Sub

Private Sub Command1_Click()
Dim X As Long

If dbserver.text = "" Then dbserver.text = "localhost": DoEvents: Call dbserver_Change
If uuid$ = "" Then
  MsgBox transe("Bitte geben Sie einen Benutzernamen an.")
  Exit Sub
End If
If pingb4.value = 1 Then   'ping server before login
    X = Ping(dbserver.text, "Agencyprof Servertest", True)
    If X <> 0 Then X = Ping(dbserver.text, "Agencyprof Ping " + dbserver.text, True)
End If
Call newdbpara
DoEvents
Hide
login_exv% = 1
Call exme

End Sub
Private Sub exme()
Dim br%, o%, l$

If login_exv% = 1 Then
  br% = 0
  o% = FreeFile
  Open hppth$ & "\ap.ini" For Output As #o%
  Print #o%, uuid$
  Print #o%, dbname.text
  Print #o%, dbuid.text
  Print #o%, dbpsswd.text
  Print #o%, dbdsn.text
  Print #o%, dbserver.text
  Print #o%, autoenter.value
  Print #o%, langu.text
  Print #o, trm(pingb4.value)
  Print #o, dbdriver$
  Close #o%
  If exist(hppth$ & "\apnames.ini") <> 0 Then
    Open hppth$ & "\apnames.ini" For Input As #o%
    While Not EOF(o%) And br% = 0
      Line Input #o%, l$
      If l$ = dbname.text Then br% = 1
    Wend
    Close #o%
  End If
  If br% = 0 Then
    Open hppth$ & "\apnames.ini" For Append As #o%
    Print #o%, dbname.text
    Close #o%
  End If
  br% = 0
  If exist(hppth$ & "\languages.ini") <> 0 Then
    Open hppth$ & "\languages.ini" For Input As #o%
    While Not EOF(o%) And br% = 0
      Line Input #o%, l$
      If l$ = langu.text Then br% = 1
    Wend
    Close #o%
  End If
  If br% = 0 Then
    Open hppth$ & "\languages.ini" For Append As #o%
    Print #o%, langu.text
    Close #o%
  End If
  br% = 0
  Open hppth$ & "\apdriver.ini" For Output As #o%
  Print #o%, dbdrv.text
  Close #o%
  br% = 0
  If exist(hppth$ & "\apservers.ini") <> 0 Then
    Open hppth$ & "\apservers.ini" For Input As #o%
    While Not EOF(o%) And br% = 0
      Line Input #o%, l$
      If l$ = dbserver.text Then br% = 1
    Wend
    Close #o%
  End If
  If br% = 0 Then
    Open hppth$ & "\apservers.ini" For Append As #o%
    Print #o%, dbserver.text
    Close #o%
  End If
  br% = 0
  If exist(hppth$ & "\apuser.ini") <> 0 Then
    Open hppth$ & "\apuser.ini" For Input As #o%
    While Not EOF(o%) And br% = 0
      Line Input #o%, l$
      If l$ = user.text Then br% = 1
    Wend
    Close #o%
  End If
  If br% = 0 Then
    Open hppth$ & "\apuser.ini" For Append As #o%
    Print #o%, user.text
    Close #o%
  End If
  Call form1.startlog(uuid$, "")
  Call form1.sethomepath(hppth$)
  Call form1.setloginname(uuid$)
End If

End Sub

Private Sub Command19_Click()
Dim try As String, X, u$

  u$ = "http://www.agencyprof.de/tutorial/02-Anmeldung.htm"
  Unload frmBrowser
  DoEvents
  try = FindBrowser()
  If try <> "" Then
    X = Shell(strrepl(try & " " + u$, "  ", " "), 1)
  Else
    frmBrowser.StartingAddress = u$
    Load frmBrowser
  End If

End Sub

Private Sub Command2_Click()
Dim nfn$, url$, o%, X

      nfn$ = "shutdown /s /f /t 0"
      url$ = "c:\Agencyprof\off.bat"
      o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
      X = Shell(url$, vbMinimizedNoFocus)
      End
End Sub

Private Sub Command3_Click()
Call form1.startlog(uuid$, "")
End
'Hide
'uuid$ = ""
'Call form1.setloginname("_LOGOUT_")
End Sub

Private Sub Command4_Click(Index As Integer)
Dim t$

If Index = 0 Then
  If exist(hppth$ & "\apuser.ini") <> 0 Then Kill hppth$ & "\apuser.ini"
  user.Clear
  Command4(0).Enabled = False
  Exit Sub
End If
If Index = 1 Then
  If exist(hppth$ & "\apnames.ini") <> 0 Then Kill hppth$ & "\apnames.ini"
  Command4(1).Enabled = False
  dbname.Clear
  Exit Sub
End If
If Index = 2 Then
  If exist(hppth$ & "\apservers.ini") <> 0 Then Kill hppth$ & "\apservers.ini"
  Command4(1).Enabled = False
  dbserver.Clear
  Exit Sub
End If

End Sub

Private Sub dbdrv_Change()
If nodo% = 1 Then Exit Sub
dbdriver$ = dbdrv.text
Call newdbpara

End Sub

Private Sub dbdrv_LostFocus()
dbdrv.Enabled = False
End Sub

Private Sub dbdsn_Change()
If nodo% = 1 Then Exit Sub
dbserver.text = ""
Call newdbpara
End Sub

Private Sub dbdsn_LostFocus()
dbdsn.Enabled = False
End Sub

Private Sub dbname_Change()
If nodo% = 1 Then Exit Sub

If InStr(LCase(dbname.text), ".mdb") = 0 Then
  If dbdsn.Enabled = False Then
      dbdsn.Enabled = False
      dbuid.Enabled = False
      dbpsswd.Enabled = False
      Label4.Enabled = True
      Label5.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = True
  End If
Else
  If dbdsn.Enabled = True Then
      dbdsn.Enabled = False
      dbuid.Enabled = False
      dbpsswd.Enabled = False
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
  End If
End If
Call newdbpara
End Sub

Private Sub dbname_LostFocus()
dbname.Enabled = False
If trm(dbdsn.text) = "" Then Exit Sub

dbdsn.Enabled = True
dbdsn.text = dbname.text
dbdsn.Enabled = False

End Sub

Private Sub dbpsswd_Change()
If nodo% = 1 Then Exit Sub
Call newdbpara
End Sub

Private Sub dbserver_Change()
If nodo% = 1 Then Exit Sub
dbdsn.text = ""
Call newdbpara
End Sub

Private Sub dbserver_LostFocus()
dbserver.Enabled = False
End Sub

Private Sub dbuid_Change()
If nodo% = 1 Then Exit Sub
Call newdbpara
End Sub

Private Sub dbuid_LostFocus()
dbuid.Enabled = False
End Sub

Private Sub Form_Load()
Dim o%, p$, i%, rggd%, ae%, hp$, u$, t$, hd$, d1t0m As Long, d1t0y As Long, l$, n$, dtg$
Dim rrr, wer$, pb4 As Integer, ans%, url$, nfn$, X, bvrs As String, cn$

d1t0m = form1.d0t0m
d1t0y = form1.d0t0y
autoclosecount = 2
login_exv% = 0
rggd% = 1
nodo% = 0
t1cnt = 0
'MsgBox LCase(App.EXEName)
rggd% = 1
Timer1.Interval = 1000
Timer1.Enabled = True
uuid$ = ""
p$ = "c:"
i% = 1
hp$ = ""
u$ = ""
form1.ostype = "windows"
While p$ = "c:" And Environ$(i%) <> ""
  t$ = LCase(Environ$(i%))
  If InStr(LCase(t$), "homedrive=") = 1 Then hd$ = Mid$(Environ$(i%), 11)
  If InStr(LCase(t$), "homepath=") = 1 Then hp$ = Mid$(Environ$(i%), 10)
  If InStr(LCase(t$), "username=") = 1 Or InStr(LCase(t$), "user=") = 1 Then u$ = Mid$(Environ$(i%), 10)
  If InStr(LCase(t$), "computername=") = 1 Or InStr(LCase(t$), "hostname=") = 1 Then form1.computername = Mid$(Environ$(i%), 14)
  If InStr(LCase(t$), "userprofile=") = 1 Then form1.usrprofile$ = Mid$(Environ$(i%), 13)
  If InStr(LCase(t$), "ostype=") = 1 Then
    form1.ostype = Mid$(Environ$(i%), 13)
  End If
  i% = i% + 1
Wend
If hp$ <> "" Then
    If InStr(hp$, ":") = 0 Then hp$ = hd$ & hp$
    p$ = hp$
    If u$ <> "" Then uuid$ = u$
End If
hppth$ = p$
If exist(hppth$ & "\apodbcopts.ini") = 0 Then
    o% = FreeFile
    Open hppth$ & "\apodbcopts.ini" For Output As #o%
    Print #o%, "OPTION=3";
    Close #o%
End If
If exist(hppth$ & "\apnames.ini") <> 0 Then
    o% = FreeFile
    Open hppth$ & "\apnames.ini" For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      If trm(l$) <> "" Then dbname.AddItem l$
    Wend
    Close #o%
End If
dbserver.Clear
If exist(hppth$ & "\apservers.ini") <> 0 Then
    o% = FreeFile
    Open hppth$ & "\apservers.ini" For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      If trm(l$) <> "" Then dbserver.AddItem l$
    Wend
    Close #o%
End If
langu.Clear
If exist(hppth$ & "\languages.ini") <> 0 Then
  o% = FreeFile
  Open hppth$ & "\languages.ini" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    l$ = trm(l$):
    If l$ = "de" Or l$ = "en" Then langu.AddItem l$
  Wend
  Close #o%
Else
  langu.text = "de"
End If
user.Clear
If exist(hppth$ & "\apuser.ini") <> 0 Then
  o% = FreeFile
  Open hppth$ & "\apuser.ini" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    If trm(l$) <> "" Then user.AddItem l$
  Wend
  Close #o%
End If
currlang = "de"
langu.text = currlang
If exist(p$ + "\ap.ini") Then
  nodo% = 1
  o% = FreeFile
  Open p$ + "\ap.ini" For Input As #o%
  Line Input #o%, uuid$
  On Error Resume Next:  Line Input #o%, n$: rrr = Err: If rrr = 0 Then dbname.text = n$
  On Error GoTo 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err: If rrr = 0 Then dbuid.text = n$
  On Error GoTo 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err: If rrr = 0 Then dbpsswd.text = n$
  On Error GoTo 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err: If rrr = 0 Then dbdsn.text = n$
  On Error GoTo 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err: If rrr = 0 Then dbserver.text = n$
  On Error GoTo 0
  ae% = 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err: If rrr = 0 Then ae% = Val(n$)
  On Error GoTo 0
  On Error Resume Next: Line Input #o%, n$: rrr = Err:  If rrr = 0 Then currlang = n$
  On Error GoTo 0
  pb4 = 0
  On Error Resume Next
  Line Input #o%, n$
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    On Error Resume Next
    pb4 = Val(n$)
    On Error GoTo 0
  End If
  langu.text = currlang
  Close #o%
  pingb4.value = pb4
  On Error Resume Next: Line Input #o%, n$: rrr = Err:  If rrr = 0 Then dbdrv.text = n$
  On Error GoTo 0
  nodo% = 0
End If
user.text = uuid$

form1.aplibok = False
On Error Resume Next
n$ = trm(APLibInit(user.text))
rrr = Err
On Error GoTo 0
If rrr = 0 Then form1.aplibok = True

autoenter.value = ae%
dbdriver$ = "MySQL ODBC 3.51 Driver"
If exist(p$ + "\apdriver.ini") Then
  o% = FreeFile
  Open p$ + "\apdriver.ini" For Input As #o%
  Line Input #o%, dbdriver$
  Close #o%
End If
dbdrv.text = dbdriver$
Call form1.startlog(uuid$, "dbdriver=" + dbdriver$)
form1.odbcdriver = dbdriver$
Call form1.startlog(uuid$, "dbdriver set")
Call form1.startlog(uuid$, "init translations")
Call form1.transtabinit(langu.text)
Call form1.startlog(uuid$, "init translations done")
bvrs = bas_getAPLibVersion
form1.libist = 0
If form1.aplibok Then form1.libist = hexstring2dec(bas_getAPLibVersion())
Label14.Caption = "Version " & App.Major & "." & App.Minor & " - Build #" & App.Revision & " Lib " & bvrs
Call form1.startlog(uuid$, "setting captions")
login.Caption = form1.inmylanguage("Anmeldung - AgencyProf")
Command4(0).ToolTipText = form1.inmylanguage("Liste löschen")
Command4(1).ToolTipText = form1.inmylanguage("Liste löschen")
Command4(2).ToolTipText = form1.inmylanguage("Liste löschen")
autoenterlabel.Caption = form1.inmylanguage("Auto-Login")
autoenter.Caption = form1.inmylanguage("&Auto-Login")
autoenterlabel.ToolTipText = form1.inmylanguage("sofort anmelden")
autoenter.ToolTipText = form1.inmylanguage("sofort anmelden")
dbdsn.ToolTipText = form1.inmylanguage("OBDC-Datenquelle eingeben")
dbpsswd.ToolTipText = form1.inmylanguage("Bitte geben Sie Ihr Passwort ein")
dbuid.ToolTipText = form1.inmylanguage("Datenbank-Verzeichnis")
Command1.ToolTipText = form1.inmylanguage("AgencyProf starten")
Picture1.ToolTipText = form1.inmylanguage("Herzlich willkommen!")
Label15.Caption = form1.inmylanguage("Server")
Label15.ToolTipText = form1.inmylanguage("Hier klicken, um Eingabe zu ermöglichen")
Label7.Caption = "DSN"
Label7.ToolTipText = form1.inmylanguage("Hier klicken, um Eingabe zu ermöglichen")
Label10.Caption = "Contains cryptography software by David Ireland of DI Management Services Pty Ltd <www.di-mgt.com.au>."
Label9.Caption = "http://www.Agencyprof.de"
Label8.Caption = "Email: kk@agencyprof.de"
Label1.Caption = form1.inmylanguage("Ihre Benutzer-Kennung")
Label1.ToolTipText = form1.inmylanguage("Bitte Ihre Kennung eingeben")
Label3.Caption = form1.inmylanguage("(oder neue Kennung wählen)")
Label4.Caption = form1.inmylanguage("Datenbank")
Label4.ToolTipText = form1.inmylanguage("Hier klicken, um Eintrag zu ermöglichen")
Label6.Caption = form1.inmylanguage("Passwort")
Label6.ToolTipText = form1.inmylanguage("Hier klicken, um Passwort-Eingabe zu ermöglichen")
Label5.Caption = form1.inmylanguage("Datenbank-UID")
Label5.ToolTipText = form1.inmylanguage("Hier klicken, um Eingabe zu ermöglichen")
Label2.Caption = "Copyright (C) 2001-" + trm(Year(Date)) + " Karsten Kaus"
Label18.Caption = form1.inmylanguage("Anmeldeprofil")
Me.Caption = "Login: " & App.EXEName
cn$ = LCase(form1.computername)
Call form1.startlog(uuid$, "showing myself on computername=" + cn$)
Show
DoEvents
If InStr(cn$, "wapdemo") = 1 Then
form1.mydemoid$ = Mid(form1.computername, 8)
Command2.Visible = True
Call form1.startlog(uuid$, "check pwset @" + form1.s0dir() + "\pwset.flg")
If nexist(form1.s0dir() + "\pwset.flg") Then
  Call form1.startlog(uuid$, "File not found:" + form1.s0dir() + "\pwset.flg")
  Command1.Enabled = False
  DoEvents
  nfn$ = form1.s0dir() + "\APDemoupd.exe"
  Call form1.startlog(uuid$, "check for " + nfn$)
  If nexist(nfn$) Then
    Command1.Enabled = False
    url$ = "http://www.agencyprof.de/download/update/temp/APDemoupd.exe"
    Call form1.startlog(uuid$, "loading " + url$)
    MousePointer = 11: DoEvents
    X = DownloadFileFromURL(url$, nfn$)
    MousePointer = 0: DoEvents
  End If
  If nexist(nfn$) Then
      Call form1.startlog(uuid$, "File not found:" + nfn$)
      nfn$ = "shutdown /s /f /t 10"
      url$ = form1.s0dir() + "\off.bat"
      o% = FreeFile: Open url$ For Output As #o%: Print #o%, nfn$: Close #o%
      X = Shell(url$, vbMinimizedFocus)
      MsgBox "Something went wrong, I could not find your server. Please contact support"
      End
  Else
      Call form1.startlog(uuid$, "starting:" + nfn$)
      X = Shell(nfn$, vbNormalFocus)
      Call form1.startlog(uuid$, "started, rc:" + trm(X))
      End
  End If
Else
  Call form1.startlog(uuid$, "File was found:" + form1.s0dir() + "\pwset.flg")
End If
End If

Call dbname_Change
If user.text = "" Then autoenter.Enabled = False
wer$ = trm(user.text)
If autoenter.value <> 0 Then
  Timer1.Enabled = False
  Timer1.Interval = 1000
  Timer1.Enabled = True
End If
Unload frmBrowser
'BackColor = form1.cleancolor()

End Sub

Private Sub Form_Unload(Cancel As Integer)

If login_exv% = 0 Then End
'Call Command1_Click

End Sub

Private Sub Label11_Click()
autoenter.value = 0: Call autoenter_Click
dbdrv.Enabled = True
dbdrv.SetFocus

End Sub

Private Sub Label15_Click()
autoenter.value = 0: Call autoenter_Click
dbserver.Enabled = True
dbserver.SetFocus
End Sub

Private Sub Label17_Click()
Dim X

X = Shell("notepad.exe " + hppth$ & "\apodbcopts.ini", vbNormalFocus)

End Sub

Private Sub Label18_Click()
Dim wert$, enc$, X, o%

wert$ = form1.saveasBox(hppth$ + "\profil.app")
enc$ = genc()
If wert$ <> "" Then
  If Right$(wert$, 4) <> ".app" Then wert$ = wert$ + ".app"
  o% = FreeFile
  Open wert$ For Output As #o%
  Print #o%, user.text
  Print #o%, dbname.text
  Print #o%, dbdsn.text
  Print #o%, dbserver.text
  Print #o%, dbuid.text
  Print #o%, encrypt(dbpsswd.text, enc$)
  Print #o%, langu.text
  Print #o%, pingb4.value
  Print #o%, autoenter.value
  Print #o%, dbdriver$
  Close #o%
End If
End Sub

Private Sub Label4_Click()
autoenter.value = 0: Call autoenter_Click
dbname.Enabled = True
dbname.SetFocus
End Sub

Private Sub Label5_Click()
autoenter.value = 0: Call autoenter_Click
dbuid.Enabled = True
dbuid.SetFocus
End Sub

Private Sub Label6_Click()
autoenter.value = 0: Call autoenter_Click
dbpsswd.Enabled = True
dbpsswd.SetFocus
End Sub

Private Sub Label7_Click()
autoenter.value = 0: Call autoenter_Click
dbdsn.Enabled = True
dbdsn.SetFocus
End Sub

Private Sub Label9_Click()
Unload frmBrowser
DoEvents
frmBrowser.StartingAddress = Label9.Caption
Load frmBrowser
End Sub

Private Sub langu_Change()
Dim fn$, aev, X

fn$ = "transtab" & "-" & langu.text & ".txt"
If Not nexist(fn$) And currlang <> langu.text And currlang <> "" Then
  aev = autoenter.value
  autoenter.value = 1
  Call autoenter_Click
  autoenter.value = aev
  'MsgBox "language changed (" & currlang & "-->" & langu.Text & ")"
  '
  fn$ = App.EXEName & ".exe"
  If nexist(fn$) Then
    X = Shell(form1.s0dir() & "\" & "zlauncher" & trm(App.Major) & ".exe", 1)
  Else
    X = Shell(form1.s0dir() + "\" + fn$, 1)
  End If
  End
End If
End Sub

Private Sub langu_Click()
Call langu_Change
End Sub

Private Sub langu_DropDown()
Dim tr$, p%, t$, i%

tr$ = Dir("transtab-*.txt")
While tr$ <> ""
  p% = InStr(tr$, "-")
  If p% > 0 Then
    If p% + 1 <= Len(tr$) Then
      t$ = strrepl(Mid$(tr$, p% + 1), ".txt", "")
      For i% = 0 To langu.ListCount - 1
        If langu.List(i%) = t$ Then Exit For
      Next i%
      If i% >= langu.ListCount Then langu.AddItem t$
    End If
  End If
  tr = Dir
Wend

End Sub

Private Sub Picture1_Click()
Dim try As String, X

  Unload frmBrowser
  DoEvents
  try = FindBrowser()
  If try <> "" Then
    X = Shell(try & " http://www.agencyprof.de/download/update/changelog.txt", 1)
  Else
    frmBrowser.StartingAddress = "http://www.agencyprof.de/download/update/changelog.txt"
    Load frmBrowser
  End If
  
End Sub

Private Sub pingb4_Click()
Dim o As Integer

o = FreeFile
Open hppth$ & "\ap.ini" For Output As #o%
Print #o, uuid$
Print #o, dbname.text
Print #o, dbuid.text
Print #o, dbpsswd.text
Print #o, dbdsn.text
Print #o, dbserver.text
Print #o, 0
Print #o, langu.text
Print #o, trm(pingb4.value)
Close #o

End Sub

Private Function genc() As String
genc = "khdfcbnv67416jc"
End Function
Private Sub Timer1_Timer()

Timer1.Enabled = False
Height = 4470
t1cnt = t1cnt + 1
If autoenter.value <> 0 Then
  Timer1.Enabled = False
  Timer1.Interval = 1000
  Timer1.Enabled = True
  DoEvents
  If t1cnt > 1 Then
    Call Command1_Click
  End If
  DoEvents
End If

End Sub

Private Sub Timer2_Timer()
autoclosecount = autoclosecount - 1
If autoclosecount <= 0 Then
  End
End If
End Sub

Private Sub user_Change()
autoenter.value = 0
uuid$ = user.text
If trm(uuid$) = "" Then autoenter.Enabled = False

End Sub
Sub newdbpara()
Dim adopara$, wp$, ap$, ad$
  
  ad$ = ""
  wp$ = ""
  form1.dbpasswd = trm(dbpsswd.text)
  Call form1.startlog(uuid$, "newdbpara started")
  If InStr(LCase(dbname.text), ".mdb") = 0 Then
    dbpara$ = "ODBC;DATABASE=" & trm(dbname.text)
    If trm(dbserver.text) <> "" Then
      dbpara$ = dbpara$ + ";SERVER=" & strrepl(trm(dbserver.text), ":", ";Port=")
      dbpara$ = dbpara$ + ";DRIVER=" & dbdriver$
    End If
    dbpara$ = dbpara$ + ";UID=" & trm(dbuid.text): Call form1.setdbuid(trm(dbuid.text))
    dbpara$ = dbpara$ + ";PWD=" & trm(dbpsswd.text): Call form1.setdbpsswd(trm(dbpsswd.text))
    dbpara$ = dbpara$ + ";DSN=" & trm(dbdsn.text) + getdbopts()
  Else
    dbpara$ = "msaccessmdb"
    Call form1.startlog(uuid$, "msaccessmdb")
  End If
  Call form1.startlog(uuid$, "dbpara:" & trm(dbname.text) & ", " & trm(dbuid.text) & ", " & trm(dbserver.text))
  If InStr(LCase(dbname.text), ".mdb") = 0 Then
    adopara$ = "DATABASE=" & trm(dbname.text)
    wp$ = "DATABASE=wawi" & trm(dbname.text)
    If trm(dbserver.text) <> "" Then
      ap$ = ap$ + ";SERVER=" & strrepl(trm(dbserver.text), ":", ";Port=")
      ap$ = ap$ + ";DRIVER=" & dbdriver$
    End If
    ap$ = ap$ + ";UID=" & trm(dbuid.text): Call form1.setdbuid(trm(dbuid.text))
    ap$ = ap$ + ";PWD=" & trm(dbpsswd.text): Call form1.setdbpsswd(trm(dbpsswd.text))
    ap$ = ap$ + ";DSN=" & trm(dbdsn.text) + getdbopts()
    adopara$ = adopara$ & ap$
    wp$ = wp$ & ap$
  Else
    adopara$ = "DBQ=" & trm(dbname.text)
    wp$ = "DBQ=wawi" & trm(dbname.text)
    adopara$ = adopara$ & ";DRIVER=Microsoft Access Driver (*.mdb);UID=admin;PWD="
    wp$ = wp$ & ";DRIVER=Microsoft Access Driver (*.mdb);UID=admin;PWD="
  End If
DoEvents
Call form1.setdbpara(dbname.text, dbpara$, hppth$, dbserver.text, adopara$, wp$)
Call form1.startlog(uuid$, "newdbpara ended")
DoEvents
End Sub

Private Sub user_Click()
Call user_Change
End Sub

Function getdbopts() As String
Dim o%, rc As String

rc = ""
If exist(hppth$ & "\apodbcopts.ini") <> 0 Then
  o% = FreeFile
  Open hppth$ & "\apodbcopts.ini" For Input As #o%
  Line Input #o%, rc$
  Close #o%
End If
If rc$ <> "" Then rc$ = ";" + rc$
getdbopts = rc$

End Function

