VERSION 5.00
Begin VB.Form launch 
   Caption         =   "Launcher"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "launch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   120
      Picture         =   "launch.frx":0442
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "launch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wrt%, rrg$
Dim rgd$(3)

Private Sub Form_Load()
Dim rr1

If Command$ = "" Then
  rrg$ = "Agencyprof" & Trim(Str(App.Major)) & ".exe"
Else
  rrg$ = Command$
End If
s0d$ = CurDir
Show
smsg ("Agencyprof wird aktualisiert in " + s0d$ + Chr$(13) + Chr$(10))
d0 = Time: While d0 = Time: DoEvents: Wend
d0 = Time: While d0 = Time: DoEvents: Wend

If Not nexist(s0d$ & "\" & "neu.Agencyprof" & Trim(Str(App.Major)) & ".exe") Then
  smsg ("Agencyprof wird aktualisiert" + Chr$(13) + Chr$(10))
  d0 = Time: While d0 = Time: DoEvents: Wend
  d0 = Time: While d0 = Time: DoEvents: Wend

  fn$ = "alt-" & datum2sql(Date) & "-" & strrepl(Time, ":", "-") & "Agencyprof" & Trim(Str(App.Major)) & ".exe"
  On Error Resume Next
  Name s0d$ & "\" & "Agencyprof" & Trim(Str(App.Major)) & ".exe" As s0d$ & "\" & fn$
  rr1 = Err
  On Error GoTo 0
  If rr1 <> 0 Then
    MsgBox ("Update failed, cannot create backup. starting old version.")
  Else
    On Error Resume Next
    Name s0d$ & "\neu." & "Agencyprof" & Trim(Str(App.Major)) & ".exe" As s0d$ & "\Agencyprof" & Trim(Str(App.Major)) & ".exe"
    rr1 = Err
    On Error GoTo 0
    If rr1 <> 0 Then
      MsgBox ("error renaming neu.Agencyprof1.exe to Agencyprof1.exe." + vbCrLf + "That is fatal, cannot start Agencyprof.")
    End If
  End If
End If

If Not nexist(s0d$ & "\" & "neu.agencyproflib.dll") Then
  smsg ("agencyproflib.dll wird aktualisiert" + Chr$(13) + Chr$(10))
  d0 = Time: While d0 = Time: DoEvents: Wend
  d0 = Time: While d0 = Time: DoEvents: Wend

  fn$ = "alt-" & datum2sql(Date) & "-" & strrepl(Time, ":", "-") & "agencyproflib.dll"
  If Not nexist(s0d$ & "\" & "agencyproflib.dll") Then
    On Error Resume Next
    Name s0d$ & "\" & "agencyproflib.dll" As s0d$ & "\" & fn$
    rr1 = Err
    On Error GoTo 0
    If rr1 <> 0 Then
      MsgBox ("Update failed, cannot create backup of the library.")
    End If
  End If
  On Error Resume Next
    Name s0d$ & "\neu.agencyproflib.dll" As s0d$ & "\agencyproflib.dll"
    rr1 = Err
  On Error GoTo 0
  If rr1 <> 0 Then
    MsgBox ("error renaming neu.agencyproflib.dll to agencyproflib.dll." + vbCrLf + "That is fatal, cannot start Agencyprof.")
  End If
End If

smsg ("Starte:" + rrg$ + Chr$(13) + Chr$(10))
wrt% = 4
Timer1.Interval = 1000
Timer1.Enabled = True

End Sub
Sub smsg(l$)
Label1.Caption = Label1.Caption + l$
DoEvents
End Sub
Sub smsg2(l$)
Label2.Caption = l$
DoEvents
End Sub

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

Private Sub Timer1_Timer()
If wrt% > 0 Then
  wrt% = wrt% - 1
  smsg2 "" & wrt% & " "
Else
  Label1.Caption = "": DoEvents
  Timer1.Enabled = False
  On Error Resume Next
  x = Shell(rrg$, 1)
  On Error GoTo 0
  End
End If
End Sub

Public Function datum2sql(dtg) As String
Dim y$, rrr, M$, d$

datum2sql = ""
If Len(dtg) > 0 Then
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

