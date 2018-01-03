VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form apupd 
   Caption         =   "Agencyrof Update"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "changelog zeigen"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "zeigt die letzten Programmänderungen"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Datenbank mit neuer Struktur erstellen"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "nur für MS-Access-Datenbanken"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   960
      Top             =   4440
   End
   Begin VB.ListBox msg 
      Height          =   1650
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Datenstruktur testen"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   960
      Top             =   2160
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "apupd.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Aktualisierung starten, auf Updates prüfen"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "apupd.frx":0B66
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Schliessen"
      Top             =   4440
      Width           =   375
   End
   Begin VB.ListBox logwin 
      Height          =   1530
      IntegralHeight  =   0   'False
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   4935
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   5760
      Top             =   4440
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   6495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   1440
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2760
      Width           =   5175
   End
End
Attribute VB_Name = "apupd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f_n$(0 To 1)
Dim f_ldn$(100), t1yp$(100)
Dim aex As Boolean, aplocked As Boolean, pdt, t3wrt%

Sub aplock()
Dim tr, whu$, dtg

tr = Dir(form1.s00dir() & "\*.run")
aplocked = True
dtg = Now
While tr <> ""
  whu$ = basename(trm(tr), ".run")
  If whu$ <> form1.getuserid() Then
    If (dtg - pdt) * 86400 > 10 Then
      logwin.AddItem Time & " Agencyprof läuft bei " & whu$ & ". warte auf Abbruch ...": DoEvents
      logwin.ListIndex = logwin.ListCount - 1
      pdt = dtg
    End If
    Exit Sub
  End If
  tr = Dir
Wend
aplocked = False

End Sub
Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command2_Click()

Call dbverok
End Sub

Private Sub Command3_Click()
Dim e As Recordset, werte$, ff$, rrr, X
Dim s As Database, tbd As TableDef, fn$, o%, l$, dbx$, i, flst$, j, r As Recordset
Dim rev, lsttbl$, t$, p%, ofn$, rdbrk As Boolean, t1$, f1$, ty1, le1, c$
Dim ftyp, fnam, fsiz, dcnt As Long, f As Field, idx As Index, indIndexObj As Index


dbx$ = "": If InStr(form1.getdbname(), ".mdb") > 0 Then dbx$ = ".mdb"
ofn$ = form1.mydatadir() & "\" & App.Major & "-" & App.Minor & dbx$ & ".dbini"
If nexist(ofn$) Then
  MsgBox ofn$ & " " + transe("nicht gefunden")
  Exit Sub
End If
MousePointer = 11: DoEvents
fn$ = form1.mydatadir() & "\neu-" & form1.getdbname()
On Error Resume Next
Kill form1.getdbname() & ".err"
On Error GoTo 0
msg.AddItem form1.sqla.name & " mit" & form1.sqla.TableDefs.Count - 1 & " Tabellen": msg.ListIndex = msg.ListCount - 1
On Error Resume Next
Kill fn$
On Error GoTo 0
Set s = wrkJet.CreateDatabase(fn$, dbLangGeneral)
s.Close
Set s = wrkJet.OpenDatabase(fn$, False, False)
For i = 0 To form1.sqla.TableDefs.Count - 1
  flst$ = ""
  If Left$(LCase(form1.sqla.TableDefs(i).name), 4) <> "msys" _
     And Left$(LCase(form1.sqla.TableDefs(i).name), 6) <> "mysql." Then
    Debug.Print form1.sqla.TableDefs(i).name
    msg.AddItem "Tabelle " & form1.sqla.TableDefs(i).name & " wird erstellt": msg.ListIndex = msg.ListCount - 1: DoEvents
    Set tbd = s.CreateTableDef(form1.sqla.TableDefs(i).name)
    For j = 0 To 99: f_ldn$(j) = "": Next j
    For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
      ftyp = form1.sqla.TableDefs(i).Fields(j).Type
      fnam = form1.sqla.TableDefs(i).Fields(j).name
      fsiz = form1.sqla.TableDefs(i).Fields(j).Size
      If flst$ <> "" Then flst$ = flst$ & ","
      flst$ = flst$ & fnam
      f_ldn$(j) = fnam
      Debug.Print ftyp; "; "; fnam; "; s="; fsiz
      If ftyp = 0 Then ftyp = 3
      t1yp$(j) = trm(ftyp)
      Debug.Print "Teste Feld: "; form1.sqla.TableDefs(i).name; " "; fnam; " "; ftyp; " "; fsiz
      o% = FreeFile
      rdbrk = False
      Open ofn$ For Input As #o%
      While Not EOF(o%) And Not rdbrk
        Line Input #o%, l$
        If InStr(l$, form1.sqla.TableDefs(i).name & "." & fnam & "." & trm(ftyp)) = 1 Then rdbrk = True
      Wend
      Close #o%
      If rdbrk Then
        p% = InStr(l$, "."): t1$ = Left$(l$, p% - 1): l$ = Mid$(l$, p% + 1)
        p% = InStr(l$, "."): f1$ = Left$(l$, p% - 1): l$ = Mid$(l$, p% + 1)
        p% = InStr(l$, "."): ty1 = Val(Left$(l$, p% - 1)): l$ = Mid$(l$, p% + 1)
        le1 = Val(l$)
        If ty1 <> ftyp Then
          ftyp = ty1
        End If
        If le1 <> fsiz Then
          fsiz = le1
        End If
      End If
      Set f = tbd.CreateField(fnam, ftyp)
      f.Size = fsiz
      tbd.Fields.Append f
    Next j
    For j = 0 To form1.sqla.TableDefs(i).Indexes.Count - 1
        Set idx = form1.sqla.TableDefs(i).Indexes(j)
        ' *** Create the index
        Set indIndexObj = tbd.CreateIndex(idx.name)
        Debug.Print "idx#" & trm(j); " Name="; idx.name
        indIndexObj.Fields = idx.Fields
        indIndexObj.Unique = idx.Unique
        If idx.name = "PRIMARY" Then
          indIndexObj.Primary = True
        Else
          Debug.Print idx.name; " nonprimary?"
        End If
        ' *** Add this index
        tbd.Indexes.Append indIndexObj
    Next j
    ' *** Append this new table
    s.TableDefs.Append tbd
    msg.AddItem "Tabelle " & form1.sqla.TableDefs(i).name & " Daten werden kopiert": msg.ListIndex = msg.ListCount - 1: DoEvents
    dcnt = 0
    c$ = "select * from " & form1.sqla.TableDefs(i).name
    Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
    Set e = s.OpenRecordset(form1.sqla.TableDefs(i).name)
    While Not r.EOF
      e.AddNew
      flst$ = ""
      werte$ = ""
      For j = 0 To form1.sqla.TableDefs(i).Fields.Count - 1
        If trm(r.Fields(j).value) <> "" Then
        ff$ = "'": If t1yp$(j) = "3" Or t1yp$(j) = "4" Or t1yp$(j) = "7" Then ff$ = ""
        flst$ = flst$ & f_ldn$(j) & ","
        If ff$ = "" Then
          werte$ = werte$ & strrepl(trm(r.Fields(j).value), ",", ".") & ","
        Else
          werte$ = werte$ & ff$ & r.Fields(j).value & ff$ & ","
        End If
        e.Fields(j).value = r.Fields(j).value
        End If
      Next j
      werte$ = Left$(werte$, Len(werte$) - 1)
      flst$ = Left$(flst$, Len(flst$) - 1)
      c$ = "insert into " & form1.sqla.TableDefs(i).name & " (" & flst$ & ") values(" & werte$ & ");"
      Debug.Print c$
      On Error Resume Next
      e.Update
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        o% = FreeFile
        Open form1.getdbname() & ".err" For Append As #o%
        Print #o%, Error$(rrr); ": "; vbCrLf; c$
        Close #o%
      End If
      dcnt = dcnt + 1
      If dcnt Mod 1000 = 0 Then
        msg.AddItem trm(dcnt) & " Sätze wurden kopiert": msg.ListIndex = msg.ListCount - 1: DoEvents
      End If
      r.MoveNext
    Wend
    msg.AddItem trm(dcnt) & " Sätze wurden kopiert": msg.ListIndex = msg.ListCount - 1: DoEvents
  End If
Next i
If exist(form1.getdbname() & ".err") <> 0 Then
  X = Shell("notepad.exe " & form1.getdbname() & ".err", 1)
  DoEvents
  On Error Resume Next
  Kill form1.getdbname() & ".err"
  On Error GoTo 0
End If
MsgBox "Eine Datenbank mit dem Namen " & fn$ & " wurde erstellt." & vbCrLf & "1. Stellen Sie sicher, daß niemand die Datenbank " & form1.getdbname() & " benutzt." & vbCrLf & "2. Sichern Sie Ihre alte Datenbank " & form1.getdbname() & " indem Sie sie umbenennen." & vbCrLf & "3. Kopieren Sie die neue Datenbank in Ihr Agencyprofverzeichnis " & form1.s00dir() & "."
X = Shell("explorer.exe " & form1.mydatadir(), vbNormalFocus)
X = Shell("explorer.exe " & form1.s00dir(), vbNormalFocus)
MousePointer = 0
End Sub

Public Sub Command4_Click()
Dim nError As Integer, i%
Dim strLocalFile As String, usehtml As Boolean, usefiles As Boolean
Dim strRemoteFile As String
Dim bTransfered As Integer
Dim fdir As String

Static bCanceled As Integer, updu$, updp$
Dim o%, l$, l1$, p%, z$, dbx$, fn$, xrev, X, rst As Boolean

rst = False: usehtml = True: usefiles = False
fdir = form1.getusersetting("autoupdate", "nein")
If Not nexist(fdir + "\1-0.dbini") And Not nexist(fdir + "\1-0.dbini") And _
   Not nexist(fdir + "\Agencyprof1.exe") And Not nexist(fdir + "\zlauncher1.exe") Then
      usefiles = True
End If
If nexist("aploader.exe") Then usehtml = False
Label1.Caption = "Aktualisierung, bitte warten ..."
Command4.Enabled = False
msg.AddItem "Version " & App.Major & "." & App.Minor & " - Build #" & App.Revision
DoEvents

If Not usehtml And Not usefiles Then
  MsgBox ("Eine FTP-Aktualisierung nicht mehr möglich. Bitte installieren Sie aploader.exe." + vbCrLf + "http://www.agencyprof.de/download/update/APLoaderSetup.zip")
  Exit Sub
End If

    MousePointer = 11: DoEvents
Call form1.killxmysettings
    dbx$ = "": If InStr(form1.getdbname(), ".mdb") > 0 Then dbx$ = ".mdb"
    fn$ = Dir(form1.s00dir() & "\0-*.dbini")
    If Not nexist(fn$) Then
      msg.AddItem "Es existiert eine ältere Version von Agencyprof."
      logwin.AddItem "Update abgebrochen"
      logwin.ListIndex = logwin.ListCount - 1
      DoEvents
      Label1.Caption = "Aktualisierung abgebrochen"
      GoTo errout
    End If
    fn$ = App.Major & "-" & App.Minor & dbx$ & ".dbini"
    strLocalFile = form1.mydatadir() & "\" & fn$
    strRemoteFile = fn$

    On Error Resume Next
    Kill strLocalFile
    On Error GoTo 0
    If usefiles Then
      bTransfered = FilesGetFile(strLocalFile, fdir + "\" + strRemoteFile)
    Else
      bTransfered = HTMLGetFile(strLocalFile, strRemoteFile)
    End If
    If Not bTransfered Then
        msg.AddItem fn$ & " konnte nicht übertragen werden."
        logwin.AddItem "Update abgebrochen"
        logwin.ListIndex = logwin.ListCount - 1
        DoEvents
        Label1.Caption = "Aktualisierung abgebrochen"
        GoTo errout
    End If
  msg.AddItem "Datenbankstruktur wird getestet."
  msg.ListIndex = msg.ListCount - 1
  DoEvents
  xrev = dbverok()
  If xrev = 0 Then
    msg.AddItem "Datenbankstruktur nicht kompatibel.": msg.ListIndex = msg.ListCount - 1
    Label1.Caption = "Agencyprof ist nicht aktuell."
    aex = False
    DoEvents
    GoTo errout
  End If
  'msg.AddItem "Die Datenbankstruktur ist kompatibel.": msg.ListIndex = msg.ListCount - 1
  DoEvents
  If xrev <= App.Revision Then
    Label1.Caption = "Aktualisierung erfolgreich"
    msg.AddItem "Version im Internet: " & trm(xrev) & ". Ihr Programm ist aktuell, ein Update nicht erforderlich.": msg.ListIndex = msg.ListCount - 1
    DoEvents
    GoTo okout
  End If
  Unload frmBrowser
  DoEvents
  frmBrowser.StartingAddress = "http://www.agencyprof.de/download/update/changelog.txt"
  Load frmBrowser
  msg.AddItem "Build #" & trm(xrev) & " wird geladen": msg.ListIndex = msg.ListCount - 1
  DoEvents

  DoEvents

  'Unload frmBrowser
  'frmBrowser.StartingAddress = "http://www.agencyprof.de/swbrett/index.html"
  'Load frmBrowser

  fn$ = "zlauncher" & trm(App.Major) & ".exe"
  strLocalFile = form1.s00dir() & "\" & fn$
  strRemoteFile = fn$
  On Error Resume Next
  Kill strLocalFile
  On Error GoTo 0
  msg.AddItem fn$ & " wird geladen": msg.ListIndex = msg.ListCount - 1: DoEvents
  If usefiles Then
      bTransfered = FilesGetFile(strLocalFile, fdir + "\" + strRemoteFile)
  Else
      bTransfered = HTMLGetFile(strLocalFile, strRemoteFile)
  End If
  If Not bTransfered Then
      msg.AddItem fn$ & " konnte nicht übertragen werden.": msg.ListIndex = msg.ListCount - 1
      logwin.ListIndex = logwin.ListCount - 1
      DoEvents
      Label1.Caption = "Aktualisierung wird fortgesetzt ..."
  End If
  DoEvents

  fn$ = "Agencyprof" & trm(App.Major) & ".exe"
  strLocalFile = "neu." & fn$
  strRemoteFile = fn$
  On Error Resume Next
  Kill strLocalFile
  On Error GoTo 0
  msg.AddItem fn$ & " wird geladen": msg.ListIndex = msg.ListCount - 1: DoEvents
  If usefiles Then
      bTransfered = FilesGetFile(strLocalFile, fdir + "\" + strRemoteFile)
  Else
      bTransfered = HTMLGetFile(strLocalFile, strRemoteFile)
  End If
  If Not bTransfered Then
      msg.AddItem fn$ & " konnte nicht übertragen werden.": msg.ListIndex = msg.ListCount - 1
      logwin.AddItem "Update abgebrochen"
      logwin.ListIndex = logwin.ListCount - 1
      DoEvents
      Label1.Caption = "Aktualisierung abgebrochen"
      GoTo errout
  End If
  
  fn$ = "AgencyprofPOPClient.exe"
  strLocalFile = "neu." & fn$
  strRemoteFile = fn$
  On Error Resume Next
  Kill strLocalFile
  On Error GoTo 0
  msg.AddItem fn$ & " wird geladen": msg.ListIndex = msg.ListCount - 1: DoEvents
  If usefiles Then
      bTransfered = FilesGetFile(strLocalFile, fdir + "\" + strRemoteFile)
  Else
      bTransfered = HTMLGetFile(strLocalFile, strRemoteFile)
  End If
  If Not bTransfered Then
      msg.AddItem fn$ & " konnte nicht übertragen werden.": msg.ListIndex = msg.ListCount - 1
      logwin.AddItem "Update abgebrochen"
      logwin.ListIndex = logwin.ListCount - 1
      DoEvents
      Label1.Caption = "Aktualisierung abgebrochen"
      GoTo errout
  End If
  
  msg.AddItem "Programm wird aktualisiert ...": msg.ListIndex = msg.ListCount - 1
  DoEvents
  rst = True
okout:
  Call form1.sqlqry("delete from sysvars where owner='sysvar_system_lastupdate';")
  z$ = "insert into sysvars (id,owner,wert) values('" & form1.newid("sysvars", "id", 16) & "','sysvar_system_lastupdate','" & Date & "');"
  Call form1.sqlqry(z$)
errout:
  If aex And Not rst Then
    Timer1.Interval = 2000
    Timer1.Enabled = True
  End If
noautoout:
MousePointer = 0
If rst Then
  Label1.Caption = "Agencyprof wird beendet ...": DoEvents
  form1.menoquit = True
  o% = FreeFile
  Open form1.s00dir() & "\lock.lck" For Output As #o%
  Print #o%, form1.computername
  Close #o%
  Do
    Call aplock
  Loop Until aplocked = False
  t3wrt% = 10
  Timer3.Interval = 1000
  Timer3.Enabled = True
End If

End Sub

Private Sub Command5_Click()
Unload frmBrowser
DoEvents
frmBrowser.StartingAddress = "http://www.agencyprof.de/download/update/changelog.txt"
Load frmBrowser
End Sub

Private Sub Form_Load()

axsResizer1.SaveControlPositions

aex = False
aplocked = True
pdt = 0
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
Label1.Caption = transe("Bitte die Aktualisierung starten")
apupd.Caption = transe("Agencyrof Update")
Command3.Caption = transe("Datenbank mit neuer Struktur erstellen")
Command3.ToolTipText = transe("nur für MS-Access-Datenbanken")
Command2.Caption = transe("Datenstruktur testen")
Command4.ToolTipText = transe("Aktualisierung starten, auf Updates prüfen")
Command1.ToolTipText = transe("Formular schliessen")
Show
msg.AddItem transe("Bitte die Aktualisierung starten"): msg.ListIndex = msg.ListCount - 1
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

Public Sub autoex()

aex = True
Label1.Caption = "Aktualisierung, bitte warten ..."
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
Unload Me
End Sub

Private Sub Timer3_Timer()
Dim X

t3wrt% = t3wrt% - 1
If t3wrt% > 0 Then
  Label1.Caption = "Fortsetzung in in " & t3wrt% & "s.": DoEvents
Else
  Timer3.Enabled = False
  Label1.Caption = "Programm wird beendet ..."
  Call form1.killlogonexit
  X = Shell("zlauncher" & trm(App.Major) & ".exe", 1)
  End
End If
End Sub

Function dbverok()
Dim dbx$, fn$, o%, rev, somer As Boolean, l$, p%, t$, f$, ty, le, testlen As Boolean
Dim ty1, rrr, le1, c$, fld As Field, restartrd As Boolean, cmds2run As Boolean
Dim sqlcoll$, msq$, ft$, X, gefragt As Boolean, ask%, dbpara$

gefragt = False
restartallover:
dbverok = 0
restartrd = True
dbpara$ = form1.getconnstr()
MousePointer = 11: DoEvents
While restartrd

restartrd = False
dbx$ = "": If InStr(form1.getdbname(), ".mdb") > 0 Then dbx$ = ".mdb"
fn$ = form1.mydatadir() & "\" & App.Major & "-" & App.Minor & dbx$ & ".dbini"
sqlcoll$ = form1.mydatadir() & "\runonce.bat"
msq$ = form1.getusersetting("mysql", "")
cmds2run = False
On Error Resume Next
Kill form1.s00dir() & "\akok.tmp"
On Error GoTo 0
o% = FreeFile
Open sqlcoll$ For Output As #o%
Print #o%, "@echo off"
Print #o%, "echo Um Ihre Datenbankstruktur zu aktualisieren, muessen einige Kommandos"
Print #o%, "echo auf Ihre Datenbank angewendet werden."
Print #o%, "echo -------------------------------------------------------------------"
If msq$ <> "" And Not nexist(msq$) Then
  Print #o%, "echo SICHERN SIE IHRE DATENBANK!"
  Print #o%, "echo UND"
  Print #o%, "echo lassen Sie diese Datei (" & sqlcoll$ & ")"
  Print #o%, "echo von Ihrem technischen Support ueberpruefen,"
  Print #o%, "echo !BEVOR! Sie sie durch Doppelklick vom Arbeisplatz aus starten."
  Print #o%, "echo -------------------------------------------------------------------"
  Print #o%, "echo   Agencyprof kann erst aktualisiert werden wenn diese Aenderungen"
  Print #o%, "echo   durchgefuehrt wurden."
  Print #o%, "echo -------------------------------------------------------------------"
'  Print #o%, "echo Danach koennen Sie Agencyprof ueber die Verwaltungsfunktionen"
'  Print #o%, "echo aktualisieren."
Else
  Print #o%, "echo Ihre Benutzereistellung mysql=... fehlt."
  Print #o%, "echo Das Ausfuehren der Datei ist daher wirkungslos und von Fehlermeldungen"
  Print #o%, "begleitet. Installieren Sie MySQL für Windows und tragen Sie den vollen"
  Print #o%, "Pfadnamen auf mysql.exe (z.B. c:\mysql\bin\mysql.exe je nach Installation)"
  Print #o%, "in Ihren Benutzerdaten ein. Versuchen Sie es danach erneut."
End If
Print #o%, "echo -------------------------------------------------------------------"
Print #o%, "echo Druecken Sie Return um fortzufahren, oder"
Print #o%, "echo              Strg c um abzubrechen."
Print #o%, "pause"
Print #o%, "echo -------------------------------------------------------------------"
Print #o%, "echo Aktualisierung laeuft ..."
Close #o%
If nexist(fn$) Then
  msg.AddItem "nicht gefunden: " + fn$: msg.ListIndex = msg.ListCount - 1
Else
  o% = FreeFile
  Open fn$ For Input As #o%
  Line Input #o%, rev
  somer = False
  While Not EOF(o%) And Not restartrd
    Line Input #o%, l$
    Debug.Print l$
    p% = InStr(l$, "."): t$ = Left$(l$, p% - 1): l$ = Mid$(l$, p% + 1)
    p% = InStr(l$, "."): f$ = Left$(l$, p% - 1): l$ = Mid$(l$, p% + 1)
    p% = InStr(l$, "."): ty = Val(Left$(l$, p% - 1)): l$ = Mid$(l$, p% + 1)
    'If ty <> 8 And ty <> 12 Then
    If ty <> 12 Then
      le = Val(l$)
      testlen = False
'      Debug.Print t$; "."; f$; "."; ty; "."; le
      On Error Resume Next
      ty1 = form1.sqla.TableDefs(t$).Fields(f$).Type
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        le1 = form1.sqla.TableDefs(t$).Fields(f$).Size
        If ty1 <> ty Then
          If ty <> 8 Then
            If Left(f$, 3) <> "opt" Then
              If (ty <> 10 And ty1 <> 12) And (ty <> 12 And ty1 <> 10) Then
                msg.AddItem "Tabelle " & t$ & " Feld: " & f$ & " typ ist " & ty1 & ", sollte " & ty & " sein.": msg.ListIndex = msg.ListCount - 1: DoEvents
                If dbx$ <> "" Then Command3.Enabled = True
                somer = True
                Exit Function
              Else
                testlen = False
              End If
            End If
          Else
            testlen = False
          End If
        End If
        If Abs(le1 - le) > 2 And testlen And Left(f$, 3) <> "opt" Then
          c$ = "Tabelle " & t$ & " Feld: " & f$ + " (" + trm(ty1) + ") Groesse ist " & le1 & ", sollte " & le & " sein."
          msg.AddItem c$: msg.ListIndex = msg.ListCount - 1
          If dbx$ <> "" Then Command3.Enabled = True
          somer = True
          MousePointer = 0
'          Exit Function
          'End If
          'If Not gefragt Then
          '  gefragt = True
          '  ask% = MsgBox("Vor einer Strukturänderung in der Datenbank sollte eine Datensicherung durchgeführt werden." & vbCrLf & "Haben Sie das getan?", vbYesNo + vbCritical + vbDefaultButton2, "Datenstruktur muss geändert werden.")
          '  If ask% = vbNo Then
          '    somer = True
          '    MousePointer = 0
          '    Exit Function
          '  End If
          'End If
        End If
      Else
       If Left(f$, 3) <> "opt" Then
        If rrr = 3265 Then
          If Not gefragt Then
            gefragt = True
            ask% = MsgBox(transe("Vor einer Strukturänderung in der Datenbank sollte eine Datensicherung durchgeführt werden.") & vbCrLf & "Haben Sie das getan?", vbYesNo + vbCritical + vbDefaultButton2, transe("Datenstruktur muss geändert werden."))
            If ask% = vbNo Then
              somer = True
              Exit Function
            End If
          End If
          c$ = "Tabelle " & t$ & ", Feld: " & f$ & " Fehler #" & rrr & " (Feld fehlt?)"
          If dbx$ <> "" Then
            Set fld = form1.sqla.TableDefs(t$).CreateField(f$, ty)
            fld.Size = le
            On Error Resume Next
            form1.sqla.TableDefs(t$).Fields.Append fld
            rrr = Err
            On Error GoTo 0
            If rrr = 0 Then
              restartrd = True
              c$ = "Tabelle " & t$ & ", Feld: " & f$ & " wurde angelegt."
            End If
          Else
            somer = True
            c$ = "Tabelle " & t$ & ", Feld: " & f$ & " Fehler #" & rrr & " " & Error$(rrr)
            Select Case ty
            Case 3: ft$ = "int"
            Case 4: ft$ = "bigint"
            Case 5: ft$ = "char(20)"          ' currency not implemented
            Case 7: ft$ = "double"
            Case 8: ft$ = "timestamp(16)"
            Case 10: ft$ = "char(" & trm(le) & ")"
            Case 12: ft$ = "longtext"
            Case Else: ft$ = ""
            End Select
            c$ = "ALTER TABLE `" & t$ & "` ADD `" & f$ & "` " & ft$
            cmds2run = True
            Call app2file(sqlcoll$, msq$ & " -h " & form1.getdbserver & " -u root -p" & form1.getdbpsswd & " -D " & form1.getdbname & " -e """ & c$ & """")
          End If
        End If
        msg.AddItem c$: msg.ListIndex = msg.ListCount - 1: DoEvents
        somer = True
       Else
         Call form1.addmissingfield(t$, f$)
       End If
      End If
    End If
  Wend
  Close #o%
'  If Not somer Then
    dbverok = rev
'  endif
End If

Wend
On Error Resume Next
Kill form1.s00dir() & "\akok.tmp"
On Error GoTo 0
MousePointer = 0: DoEvents
End Function

Public Function HTMLGetFile(nach$, von$) As Boolean
Dim o%, X, tr, rrr


HTMLGetFile = False
o% = FreeFile
Open "aploader.ini" For Output As #o%
Print #o%, form1.getusersetting("htmlupdate", "http://www.agencyprof.de/download/update") & "/" & von$
Print #o%, nach$
Print #o%, Me.Top + Command4.Top + 20
Print #o%, Me.Left + Command4.Left + 20
Close #o%
X = Shell("aploader.exe", 1)
tr = "x"
While tr <> ""
  On Error Resume Next
  tr = Dir("aploader.ini")
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then tr = ""
  DoEvents
  DoEvents
Wend
If Not nexist(nach$) Then HTMLGetFile = True

End Function

Public Function FilesGetFile(nach$, von$) As Boolean
Dim o%, X, tr, rrr


FilesGetFile = False
On Error Resume Next
Kill nach$
Call FileCopy(von$, nach$)
On Error GoTo 0
If Not nexist(nach$) Then FilesGetFile = True

End Function

