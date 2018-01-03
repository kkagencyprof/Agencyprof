VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form mexplore 
   Caption         =   "Mail-Details"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   LinkTopic       =   "Form2"
   ScaleHeight     =   6240
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command18 
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
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2160
      Picture         =   "mexplore.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Diese Mail im Explorer öffnen"
      Top             =   5760
      Width           =   375
   End
   Begin VB.ListBox usrfnlist 
      Height          =   1455
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.ListBox fnlist 
      Height          =   1575
      IntegralHeight  =   0   'False
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1695
      IntegralHeight  =   0   'False
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox mtxt 
      Height          =   4095
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Text            =   "mexplore.frx":062A
      Top             =   360
      Width           =   8655
   End
   Begin VB.ListBox List2 
      Height          =   3735
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "mexplore.frx":0633
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Formular schliessen"
      Top             =   5760
      Width           =   375
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   6600
      Top             =   120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox mhead 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Text            =   "mexplore.frx":0883
      Top             =   4560
      Width           =   8655
   End
   Begin VB.Label fnam 
      BackStyle       =   0  'Transparent
      Caption         =   "Datei:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "mexplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bd$

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command18_Click()

Call form1.handbuchcall("13-Email.htm")
End Sub

Private Sub Command19_Click()
Dim fn$, X
fn$ = form1.s0dir() & "\tmp\" & basename(fnam.Caption, ".msg")
X = Shell("explorer.exe " & fn$, vbNormalFocus)

End Sub

Private Sub fnam_Change()
Dim fn$, hd%, o%, b$, l$, p%, d$, M%, i%, pnr%, c$
Dim tdsc$, tdsd$, tdsn$, mfn$, tdse$
Dim decde%

mhead.text = ""
mtxt = ""
pnr% = 0
List1.Clear
List2.Clear
fnlist.Clear
usrfnlist.Clear
fn$ = fnam.Caption
If exist(fn$) = 0 Then Exit Sub

o% = form1.myfontsize()
mtxt.Font.Size = o%

d$ = form1.s0dir() & "\tmp"
On Error Resume Next
MkDir d$
On Error GoTo 0
d$ = form1.s0dir() & "\tmp\" & basename(fn$, ".msg")
On Error Resume Next
MkDir d$
On Error GoTo 0
bd$ = d$
o% = FreeFile
Open fn$ For Input As #o%
M% = FreeFile
Open d$ & "\msgin.eml" For Output As #M%
While Not EOF(o%)
  Line Input #o%, l$
  If InStr(l$, Chr$(13)) <> 0 Or InStr(l$, Chr$(10)) <> 0 Then
    c$ = Chr$(13)
    If InStr(l$, c$) = 0 Then
      l$ = strrepl(l$, c$, "")
      c$ = Chr$(10)
    Else
      l$ = strrepl(l$, Chr$(10), "")
    End If
    l$ = strrepl(l$, c$, vbCrLf)
  End If
  Print #M%, l$
Wend
Close #M%
Close #o%
o% = FreeFile
Open d$ & "\msgin.eml" For Input As #o%
M% = FreeFile
Open d$ & "\msg.hdr" For Output As #M%
hd% = 1
While Not EOF(o%)
  Line Input #o%, l$: l$ = trm(l$)
  p% = InStr(LCase(l$), "boundary=")
  If p% > 0 Then
    b$ = Mid$(l$, p% + 9): b$ = trm(strrepl(b$, """", " "))
    List1.AddItem b$
  End If
  If l$ = "" And hd% = 1 Then
    hd% = 0
    Close #M%
    mfn$ = d$ & "\msg.txt"
    Open mfn$ For Output As #M%
  Else
    If hd% = 0 Then
      For i% = 0 To List1.ListCount - 1
        If InStr(l$, "--" & List1.List(i%)) = 1 Then
          pnr% = pnr% + 1
          Close #M%
          If decde% = 1 Then
            If tdsn$ = "" Then tdsn$ = mknam(8) + ".txt"
            Call decodepending(tdse$, mfn$, d$ & "\" & tdsn$)
          End If
          tdsn$ = mknam(8) + ".txt"
          mfn$ = d$ & "\Teil." & trm(pnr%)
          Open mfn$ For Output As #M%
          Line Input #o%, l$
          tdsc$ = "Teil." & trm(pnr%): tdsd$ = "": tdsn$ = "": tdse$ = ""
          decde% = 0
          If InStr(LCase(l$), "content-type: ") = 1 Then
            tdsd$ = ":" & trm(Mid$(l$, 14))
            Do
              Line Input #o%, l$
              p% = InStr(LCase(l$), "transfer-encoding")
              If p% > 0 Then
                tdse$ = trm(strrepl(Mid$(l$, p% + 18), """", ""))
              End If
              p% = InStr(LCase(l$), "name=")
              If p% > 0 Then
                tdsn$ = strrepl(Mid$(l$, p% + 5), """", "")
                tdsn$ = strrepl(tdsn$, ">", "")
              End If
            Loop Until trm(l$) = ""
            If LCase(tdse$) = "base64" And tdsn$ = "" Then
              tdsn$ = tdsc$ + ".decoded.txt"
            End If
          End If
          List2.AddItem tdsc$ & tdsd$
          fnlist.AddItem tdsn$
          If tdsn$ <> "" Then
            DoEvents
            decde% = 1
          End If
          GoTo alldone4thisline
        End If
      Next i%
      If pnr% < 2 Then
        If mtxt.text <> "" Then mtxt.text = mtxt.text & vbCrLf
        mtxt.text = mtxt.text & l$
      End If
    Else
      If mhead.text <> "" Then mhead.text = mhead.text & vbCrLf
      mhead.text = mhead.text & l$
      p% = InStr(LCase(l$), "transfer-encoding")
      If p% > 0 Then
        tdse$ = trm(strrepl(Mid$(l$, p% + 18), """", ""))
        decde% = 1
      End If
      p% = InStr(LCase(l$), "name=")
      If p% > 0 Then
        tdsn$ = strrepl(Mid$(l$, p% + 5), """", "")
        tdsn$ = strrepl(tdsn$, ">", "")
        decde% = 1
      End If
    End If
  End If
  Print #M%, l$
alldone4thisline:
Wend
Close #M%
If decde% = 1 Then
  If tdsn$ = "" Then tdsn$ = mknam(8) + ".txt"
  Call decodepending(tdse$, mfn$, d$ & "\" & tdsn$)
  If Not nexist(d$ & "\" & tdsn$) Then
    o% = FreeFile
    Open d$ & "\" & tdsn$ For Input As #o%
    While Not EOF(o%)
      Line Input #o%, l$
      l$ = strrepl(l$, Chr$(13), "")
      l$ = strrepl(l$, Chr$(10), vbCrLf)
      mtxt.text = mtxt.text + vbCrLf + l$
    Wend
    Close #o%
  End If
End If
Close #o%


End Sub

Private Sub Form_Load()

axsResizer1.SaveControlPositions

Dim ufsze%
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
ufsze% = Val(form1.getusersetting("fontsize"))
If ufsze% < 8 Or ufsze% > 12 Then ufsze% = 8
List1.Font.Size = ufsze%
List2.Font.Size = ufsze%
mtxt.Font.Size = ufsze%
mhead.Font.Size = ufsze%

mexplore.Caption = transe("Mail-Details")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command19.ToolTipText = transe("Diese Mail im Explorer öffnen")
Command1.ToolTipText = transe("Formular schliessen")
fnam.Caption = transe("Datei:")
mexplore.Caption = transe("Mail-Details")
Command18.ToolTipText = transe("Hilfeseite öffnen")
Command19.ToolTipText = transe("Diese Mail im Explorer öffnen")
Command1.ToolTipText = transe("Formular schliessen")
fnam.Caption = transe("Datei:")
Show

End Sub
Private Sub Form_Unload(Cancel As Integer)

Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub

Private Sub Form_Resize()

axsResizer1.Resize
End Sub

Private Sub List2_Click()
Dim fn$, o%, i%, l$, bfn$, ct$, X, prtfl%, appt$, exe$, orgfile$

i% = List2.ListIndex
mtxt = ""
If i% < 0 Then Exit Sub
bfn$ = fnlist.List(i%)
bfn$ = List2.List(i%)
o% = InStr(bfn$, ":")
If o% > 0 Then
  ct$ = word1(LCase(Mid$(bfn$, o% + 1))): If Right$(ct$, 1) = ";" Then ct$ = Left$(ct$, Len(ct$) - 1)
  bfn$ = Left$(bfn$, o% - 1)
End If
fn$ = bd$ & "\" & bfn$
If Not nexist(fn$ + ".decoded.txt") Then fn$ = fn$ + ".decoded.txt"
prtfl% = 1
If bfn$ <> "" Then
  prtfl% = 0
End If
Select Case ct$
  Case "":
  Case "text/plain": If prtfl% = 0 Then prtfl% = 1
  Case "text/html":
      Unload frmBrowser
      DoEvents
      If Right$(fn$, 4) = ".txt" Then
        If nexist(fn$ + ".jpg") Then Call FileCopy(fn$, fn$ + ".htm")
        fn$ = fn$ + ".htm"
      End If
      frmBrowser.StartingAddress = "file:////" & fn$
      Load frmBrowser
  Case Else:
      prtfl% = 1
      exe$ = form1.getusersetting(LCase(ct$))
      If exe$ = "" Then exe$ = form1.getsystemsetting(LCase(ct$))
      If exe$ <> "" Then
        If nexist(fn$ + ".jpg") Then Call FileCopy(fn$, fn$ + ".jpg")
        If Not nexist(form1.fixfilename(fn$ + ".jpg")) Then X = Shell("""" & form1.fixfilename(exe$) & """ " & form1.fixfilename(fn$ + ".jpg"), 1)
      End If
End Select
'If prtfl% = 1 Then
  o% = FreeFile
  Open fn$ For Input As #o%: i% = 0
  While Not EOF(o%) And i% < 32000
    Line Input #o%, l$
    If mtxt <> "" Then mtxt = mtxt & vbCrLf
    mtxt = mtxt & l$
    i% = i% + Len(l$)
  Wend
  Close #o%
'End If

End Sub

Sub decodepending(code$, Src$, dst$)

'If exist(dst$) = 0 Then
  If LCase(code$) = "base64" Then
    usrfnlist.AddItem "decoded: " & FileName(dst$)
    Call DecodeFileB64(Src$, dst$)
    Exit Sub
  End If
  If code$ = "" Then
    Call FileCopy(Src$, dst$)
    usrfnlist.AddItem "erstellt: " & FileName(dst$)
    Exit Sub
  End If
  usrfnlist.AddItem "code unbekannt: " & code$ & " in " & FileName(Src$)
'Else
  usrfnlist.AddItem "im Cache: " & FileName(dst$)
'End If

End Sub

Private Sub usrfnlist_Click()
Call Command19_Click
End Sub
