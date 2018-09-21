VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+nh"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "set"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Text            =   "XXXXX"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dtrg As Double, sd0 As String
Dim mplr$

Private Sub Command2_Click()
thn = CDate(Text4.Text)
'Text2.Text = thn
dtrg = thn
End Sub

Private Sub Command3_Click()
l = Val(Text5.Text)
'Text2.Text = Now + l / 24
Text4.Text = Now + l / 24
dtrg = Now + l / 24

End Sub

Private Sub Command4_Click()
'Text2.Text = Now
Text4.Text = Now
dtrg = Now

End Sub

Private Sub Form_Load()
Dim o%, l$, c$

sd0 = CurDir

Text5.Text = "5"
Call Command3_Click
Call Timer1_Timer
'mplr$ = "C:\Programme\Windows Media Player\wmplayer.exe"
mplr$ = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"
If nexist(mplr$) Then mplr$ = "C:\Program Files\Windows Media Player\wmplayer.exe"
o% = FreeFile
On Error Resume Next
Open sd0 + "\wecker.ini" For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Line Input #o%, l$
  Line Input #o%, c$
  Close #o%
  On Error Resume Next
  Kill sd0 + "\wecker.ini"
  On Error GoTo 0
  Me.Caption = c$
  Text4.Text = l$
  Call Command2_Click
End If
End Sub

Private Sub Timer1_Timer()
Dim dn As Double, difft As Double, delh As Long, delm As Long, dels As Integer
Dim dm$, ds$, ddt As Long

dn = Now
Text1.Text = Date & " " & Time
difft = dtrg - dn
ddt = Int(difft * 24 * 60 * 60)
Text3.Text = trm(ddt)
delh = Int(ddt / 3600)
delm = Int((ddt - delh * 3600) / 60)
dm$ = trm(delm): If Len(dm$) = 1 Then dm$ = "0" + dm$
dels = ddt Mod 60: ds$ = trm(dels): If Len(ds$) = 1 Then ds$ = "0" + ds$
Text2.Text = trm(delh) + ":" + dm$ + ":" + ds$
If difft < 0 Then
'  Text2.Text = Now + 1 / 144
  Text4.Text = Now + 1 / 144
  dtrg = Now + 1 / 144
  x = Shell(mplr$ + " " + sd0 + "\weck.mp3", 1)
  'x = Shell("d:\mm4\mmjb.exe weck.mp3", 1)
End If

End Sub

Public Function trm(l) As String
Dim rrr
On Error Resume Next
trm = Trim("" & l)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then trm = ""
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

