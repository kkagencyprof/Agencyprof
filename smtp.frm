VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form smtp 
   Caption         =   "Email senden"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9075
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   8640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Caption         =   "aus Datei"
      Height          =   255
      Left            =   7560
      TabIndex        =   80
      ToolTipText     =   "from file"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox tagchk 
      Height          =   255
      Left            =   4320
      TabIndex        =   78
      Top             =   4380
      Width           =   255
   End
   Begin VB.ComboBox marken 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   6720
      TabIndex        =   76
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   10
      Left            =   8280
      TabIndex        =   75
      Top             =   4080
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Left            =   480
      Top             =   1680
   End
   Begin VB.CommandButton prvw 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "smtp.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Vorschau"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   9
      Left            =   7560
      TabIndex        =   72
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   71
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   70
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   69
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   68
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   67
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   66
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   65
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   64
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   63
      Top             =   4080
      Width           =   615
   End
   Begin VB.CheckBox srcpt 
      Height          =   255
      Left            =   120
      TabIndex        =   61
      ToolTipText     =   $"smtp.frx":0532
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox merk0b 
      Height          =   315
      Left            =   9360
      TabIndex        =   60
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox merk0t 
      Height          =   735
      Left            =   9360
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox merk0 
      Height          =   1035
      IntegralHeight  =   0   'False
      Left            =   9360
      TabIndex        =   58
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   0
      MaskColor       =   &H00000000&
      Picture         =   "smtp.frx":0634
      Style           =   1  'Grafisch
      TabIndex        =   57
      ToolTipText     =   "Speichern"
      Top             =   3840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Testmail"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   56
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   600
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   1080
      MultiSelect     =   2  'Erweitert
      TabIndex        =   55
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   480
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   1080
      MultiSelect     =   2  'Erweitert
      TabIndex        =   54
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox List5 
      Height          =   645
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   53
      Top             =   840
      Width           =   615
   End
   Begin VB.ListBox yattach 
      Height          =   735
      IntegralHeight  =   0   'False
      Left            =   9360
      TabIndex        =   51
      Top             =   3840
      Width           =   1215
   End
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
      Height          =   495
      Left            =   720
      TabIndex        =   48
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   5940
      Width           =   255
   End
   Begin VB.CheckBox askb4send 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   46
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox eclient 
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3240
      Width           =   255
   End
   Begin VB.ListBox xattach 
      Height          =   615
      IntegralHeight  =   0   'False
      Left            =   9360
      TabIndex        =   43
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C0C0C0&
      Caption         =   "absenden"
      Height          =   615
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "smtp.frx":0D36
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Senden"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      MaskColor       =   &H00000000&
      Picture         =   "smtp.frx":0EC0
      Style           =   1  'Grafisch
      TabIndex        =   42
      ToolTipText     =   "Speichern"
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Picture         =   "smtp.frx":152C
      Style           =   1  'Grafisch
      TabIndex        =   41
      Top             =   5940
      Width           =   615
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   8280
      Top             =   4800
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command7 
      Caption         =   "all&e senden"
      Height          =   315
      Left            =   1080
      TabIndex        =   40
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&NEU"
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   0
      Top             =   1560
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
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
      Left            =   840
      TabIndex        =   37
      Top             =   4920
      Width           =   255
   End
   Begin VB.ListBox List4 
      Height          =   840
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   35
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   240
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   4320
      TabIndex        =   31
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox lblstatus 
      Height          =   1620
      Left            =   6240
      TabIndex        =   30
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   840
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   450
      Index           =   0
      Left            =   1080
      MultiSelect     =   2  'Erweitert
      TabIndex        =   27
      Top             =   720
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   4800
   End
   Begin VB.TextBox kid 
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox adrid 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3945
      TabIndex        =   22
      Top             =   7920
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
      Height          =   315
      Left            =   945
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   7320
      Width           =   2025
   End
   Begin VB.TextBox txtMailFrom 
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   7320
      Width           =   2235
   End
   Begin VB.TextBox txtBCC 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtCC 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   8520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtMessageHTML 
      Height          =   225
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   11
      ToolTipText     =   "Anhang entfernen"
      Top             =   4680
      Width           =   1065
   End
   Begin VB.ListBox lstAttachments 
      Height          =   1620
      Left            =   2760
      OLEDropMode     =   1  'Manuell
      TabIndex        =   9
      ToolTipText     =   "Sie können <Drag & Drop> benutzen um Dateien anzufügen"
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      ToolTipText     =   "Anhang hinzufügen"
      Top             =   4680
      Width           =   1065
   End
   Begin VB.TextBox txtMessageText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   1080
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      Top             =   1440
      Width           =   7815
   End
   Begin VB.TextBox txtMessageSubject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtSendTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   3105
   End
   Begin VB.TextBox replyto 
      Height          =   315
      Left            =   7080
      TabIndex        =   49
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "test HTML-Tags"
      Height          =   255
      Left            =   4560
      TabIndex        =   79
      Top             =   4380
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Textmarke:"
      Height          =   255
      Left            =   5880
      TabIndex        =   77
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Textfarbe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   74
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Quittung"
      Height          =   255
      Left            =   360
      TabIndex        =   62
      ToolTipText     =   "Bestätigt den Eingang der Mail im Postfach des Empfängers - NICHT das Lesen durch den Empfänger"
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label wrbytes 
      Alignment       =   1  'Rechts
      Height          =   255
      Left            =   7320
      TabIndex        =   52
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Reply-To:"
      Height          =   195
      Left            =   6360
      TabIndex        =   50
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Senden bestätigen"
      Height          =   195
      Left            =   1320
      TabIndex        =   47
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "ext. Client"
      Height          =   255
      Left            =   360
      TabIndex        =   45
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Rechts
      Height          =   255
      Left            =   2640
      TabIndex        =   38
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "gepeicherte Mail::"
      Height          =   195
      Left            =   1080
      TabIndex        =   36
      Top             =   4680
      Width           =   1260
   End
   Begin VB.Label Label15 
      Caption         =   "aus Adressen"
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Interne"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Meldungen"
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3135
      TabIndex        =   23
      Top             =   7920
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Enabled         =   0   'False
      Height          =   195
      Left            =   105
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "blinde Kopie an"
      Height          =   195
      Left            =   195
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Verteiler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "MessageHTML"
      Height          =   195
      Left            =   2385
      TabIndex        =   17
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Anhänge"
      Height          =   195
      Left            =   2760
      TabIndex        =   16
      Top             =   4680
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Betreff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   15
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "An"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Absender"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   7320
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Server"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   7320
      Width           =   465
   End
End
Attribute VB_Name = "smtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastsavename As String, mailstopped As Boolean
Dim tags$(19), tgcnt%
Dim tagsc$(19)

Private Function stripimgpart1(intxt As String) As String
Dim p%, rc As String, l$, fn$, infn$, o%

  p% = InStr(intxt, "<img src=")
  If p% = 1 Then
    l$ = Mid(intxt, 10)
    p% = InStr(l$, ">")
    If p% = 0 Then
      stripimgpart1 = Mid$(intxt, 2)
    Else
      infn$ = Left(l$, p% - 1)
      If Mid$(infn$, 2, 1) = ":" Or Left$(infn$, 2) = "\\" Then
        fn$ = form1.s0dir() & "\tmp\" & GUID() & ".b64"
Debug.Print infn$; " -> "; fn$
        Call EncodeFileB64(infn$, fn$)
        o% = FreeFile: rc = ""
        Open fn$ For Input As #o%
        While Not EOF(o%)
          Line Input #o%, l$
          rc = rc + l$
        Wend
        Close #o%
        stripimgpart1 = "<img src='data:image/" + FileExtension(infn$) + ";base64," + rc$ + "'>"
        On Error Resume Next
        Kill fn$
        On Error GoTo 0
      Else
        stripimgpart1 = "<img src=" + infn$ + ">"
      End If
    End If
    Exit Function
  End If
  If p% = 0 Then
    stripimgpart1 = intxt
  Else
    stripimgpart1 = Left$(intxt, p% - 1)
  End If

End Function

Private Function stripimgpart2(intxt As String) As String
Dim p%, rc As String, l$, picl As String

  p% = InStr(intxt, "<img src=")
  If p% = 1 Then
    l$ = Mid(intxt, 10)
    p% = InStr(l$, ">")
    If p% = 0 Then
      stripimgpart2 = ""
    Else
      stripimgpart2 = Mid$(l$, p% + 1)
    End If
    Exit Function
  End If
  If p% = 0 Then
    stripimgpart2 = ""
  Else
    stripimgpart2 = trm(Mid$(intxt, p%))
  End If

End Function

Private Sub imginclude()
Dim c As String, txg$
Dim f$, l$, s0 As Integer


    On Error Resume Next
    With cdlg1
    'Bei "Abbruch" Fehler raisen lassen:
    .CancelError = True
    'Suchpfad einstellen:
    .InitDir = ""
    .FileName = FileName("*.png")
    .DialogTitle = transe("Bild wählen ...")
    'und endlich den Dialog anzeigen:
    .ShowOpen

    'Auswertung:
    If Err = cdlCancel Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    c = .FileName

    End With
    On Error GoTo 0


l$ = txtMessageText.text: f$ = ""
s0 = txtMessageText.SelStart
If txtMessageText.SelStart > 0 Then
  f$ = Left(txtMessageText.text, txtMessageText.SelStart)
  l$ = Mid$(txtMessageText.text, txtMessageText.SelStart + 1)
End If
txg$ = "<img src=" + c + ">"
txtMessageText.text = f$ + txg$ + l$
txtMessageText.SelStart = s0 + Len(txg$)
Call txtMessageText.SetFocus
Call txtMessageTextchg

End Sub
Function bkmks(mtxt$, adrid$, kid$) As String
Dim rc$, rest$, bkn$, p%, cmd$, subst$, lbkn$, ths$
Dim land$, plz$, ort$

Debug.Print adrid$ + ", k=" + kid$
bkmks = mtxt$
rest$ = mtxt$
rc$ = ""
While Len(rest$) > 0
  p% = InStr(rest$, "<!--bkmkstart-->")
  If p% > 0 Then
    If p% > 1 Then rc$ = rc$ + Left(rest$, p% - 1)
    rest$ = Mid(rest$, p% + 16)
    p% = InStr(rest$, "<!--bkmkend-->")
    bkn$ = ""
    If p% > 0 Then
      If p% > 1 Then bkn$ = Left(rest$, p% - 1)
      rest$ = Mid(rest$, p% + 14)
    End If
    If bkn$ <> "" Then
      If InStr(bkn$, "{") > 0 Then bkn$ = Mid(bkn$, 2, Len(bkn$) - 2)
Debug.Print bkn$
      lbkn$ = LCase(bkn$): subst$ = ""
      If lbkn$ = "anrede" Then subst$ = cut_d1(trm(form1.getanabrede(adrid$, kid$)), "|")
      If lbkn$ = "abrede" Then subst$ = cut_d2bis(trm(form1.getanabrede(adrid$, kid$)), "|")
      If lbkn$ = "hinweise" And kid$ <> "" And kid$ <> "-1" Then subst$ = " "
      If lbkn$ = "plzort" Then
        If kid$ = "" Or kid$ = "-1" Then
          cmd$ = "select land as wert from adresse where id='" + adrid$ + "'": land$ = trm(form1.get1erg(cmd$))
          cmd$ = "select plz as wert from adresse where id='" + adrid$ + "'": plz$ = trm(form1.get1erg(cmd$))
          cmd$ = "select ort as wert from adresse where id='" + adrid$ + "'": ort$ = trm(form1.get1erg(cmd$))
          If LCase(land$) = LCase(form1.getusersetting("meinland", "")) Then land = ""
          subst$ = form1.getplzort(land$, plz$, ort$)
        Else
          cmd$ = "select lkz as wert from kontakt where id='" + kid$ + "'": land$ = trm(form1.get1erg(cmd$))
          If land$ = "" Then
            cmd$ = "select land as wert from adresse where id='" + adrid$ + "'": land$ = trm(form1.get1erg(cmd$))
          End If
          cmd$ = "select plz as wert from kontakt where id='" + kid$ + "'": plz$ = trm(form1.get1erg(cmd$))
          If plz$ = "" Then
            cmd$ = "select plz as wert from adresse where id='" + adrid$ + "'": plz$ = form1.get1erg(cmd$)
          End If
          cmd$ = "select ort as wert from kontakt where id='" + kid$ + "'": ort$ = form1.get1erg(cmd$)
          If ort$ = "" Then
            cmd$ = "select ort as wert from adresse where id='" + adrid$ + "'": ort$ = form1.get1erg(cmd$)
          End If
          subst$ = form1.getplzort(land$, plz$, ort$)
        End If
      End If
      If subst$ = "" Then
        If kid$ = "" Or kid$ = "-1" Then
          cmd$ = "select " + lbkn$ + " as wert from adresse where id='" + adrid$ + "'"
Debug.Print cmd$
        Else
          ths$ = lbkn$
          If lbkn$ = "land" Then ths$ = "lkz"
          cmd$ = "select " + ths$ + " as wert from kontakt where id='" + kid$ + "'"
Debug.Print cmd$
          subst$ = form1.get1erg(cmd$)
          If subst$ = "" Then
            cmd$ = "select " + lbkn$ + " as wert from adresse where id='" + adrid$ + "'"
Debug.Print cmd$
          End If
        End If
        subst$ = form1.get1erg(cmd$)
      End If
      rc$ = rc$ + trm(subst$)
    End If
  Else
    rc$ = rc$ + rest$: rest$ = ""
  End If
Wend
bkmks = rc$

End Function


Private Sub askb4send_Click()

'd2infile = "smtp": d2insub = "askb4send_Click"
Call form1.setmylastFormVar(Me.name, "ab4s", trm(askb4send.value))

End Sub

Public Sub cmdAdd_Click()
Dim rrr, o%, l$, p%, ifn$, ofn$, wr%

'd2infile = "smtp": d2insub = "cmdAdd_Click"
If xattach.ListCount > 0 Then
  While xattach.ListCount > 0
    ifn$ = xattach.List(0)
    ofn$ = form1.myuniquedocname("noask")
    o% = FreeFile
    Open ifn$ For Input As #o%
    p% = FreeFile
    Open ofn$ For Output As #p%
    wr% = 0
    While Not EOF(o%)
      Line Input #o%, l$
      If wr% = 0 And InStr(LCase(l$), "content-type:") = 1 Then wr% = 1
      If wr% = 1 Then Print #p%, l$
    Wend
    Close #p%
    Close #o%
    Call attachfile(xattach.List(0))
    xattach.RemoveItem 0
  Wend
  Exit Sub
End If
Load fselect
fselect.Visible = True
fselect.fqn.text = ""
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub cmdRemove_Click()
Dim i%
'd2infile = "smtp": d2insub = "cmdRemove_Click"
  i% = lstAttachments.ListIndex
  If i% < 0 Then Exit Sub
  Call detachfile(lstAttachments.List(i%))
  lstAttachments.RemoveItem i%

End Sub

Public Sub cmdSend_Click()
Dim strBuffer As String, fsiz, bndry$, o1%, ifn$, bcc$, c$, anadr$
Dim intPos As Integer, nureinmalhier As Integer, attl$, rrr, sndpart As String
Dim strMess As String, mboxfile As String, optfile As String
Dim volltext$, i%, t$, n%, id$, k$, eml$, o%, l$, ask%, mbox$, X, p%, lockfile
Dim txtMF$, cmdl$, mlto$, mlclf$, bag$, ucf$, dtg$, mailusefont$
Dim htmldraft As Boolean, txt2send As String, lcc$, addhd$, orgfn As String

'd2infile = "smtp": d2insub = "cmdSend_Click"
If txtMessageSubject.text = "" Then
  ask% = MsgBox(transe("Die Betreffzeile ist leer.") + vbCrLf + transe("Email ohne Betrff senden?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Kein Betreff?"))
  If ask% = vbNo Then
    Call txtMessageSubject.SetFocus
    Exit Sub
  End If
End If
mailstopped = False
Unload emailadrselect
htmldraft = False
If form1.getusersetting("htmldraft", "ja") = "ja" Then htmldraft = True
If eclient.value <> 0 Then
  t$ = form1.getusersetting("sendmail", "")
  If InStr(LCase(t$), "thunderbird") > 0 And InStr(LCase(t$), "-remote") > 0 Then
    txtServer = "thunderbirdremote"
  Else
    t$ = form1.getusersetting("mailclient")
    If InStr(LCase(t$), "netscape") > 0 Or LCase(form1.getusersetting("Mozillaclient")) = "ja" Then txtServer = "NETSCAPE47"
    If InStr(LCase(t$), "outlook") > 0 Then txtServer = "OUTLOOK"
  End If
End If
If trm(txtServer) = "" Or LCase(trm(txtServer)) = "dir:inbox" Then txtServer = "dir:Outbox"
If trm(txtMailFrom) = "" Then
  Beep
  MsgBox "Absender unbekannt"
  Exit Sub
End If
If txtServer <> "NETSCAPE47" And txtServer <> "thunderbirdremote" And txtServer <> "OUTLOOK" Then
  If askb4send.value = 1 Then
    ask% = MsgBox("Email abschicken?", vbYesNo + vbCritical + vbDefaultButton1, "Email senden")
    If ask% = vbNo Then Exit Sub
  End If
End If
Call form1.dbg2f("sending mail ...")
If List1(0).ListCount > 0 Then
  cmdStop.Top = cmdSend.Top
  cmdStop.Left = cmdSend.Left
  cmdStop.Visible = True
  cmdSend.Visible = False
End If
mailusefont$ = strrepl(form1.getusersetting("mailfont", ""), ":", "'")
While List1(0).ListCount > 0 Or Len(trm(txtSendTo.text)) <> 0
  MousePointer = 11
  
    anadr$ = ""
    txt2send = mailusefont$ + txtMessageText.text
    If mailusefont$ <> "" Then txt2send = txt2send + "</font>"
    nureinmalhier = 0
    If Len(txtSendTo.text) <> 0 Then nureinmalhier = 1
    While List1(0).ListCount > 0 And nureinmalhier = 0
      nureinmalhier = nureinmalhier + 1
      t$ = List1(0).List(0)
      List1(0).RemoveItem 0
      n% = InStr(t$, "|"): id$ = Left$(t$, n% - 1): t$ = Mid$(t$, n% + 1)
      n% = InStr(t$, "|")
      If n% > 0 Then
        k$ = Left$(t$, n% - 1): eml$ = Mid$(t$, n% + 1)
      Else
        k$ = "-1": eml$ = t$
      End If
      kid.text = k$
      adrid.text = id$
      If Len(txtSendTo.text) = 0 Then
        txtSendTo.text = "<" & eml$ & ">"
      Else
        txtSendTo.text = txtSendTo.text + ",<" + eml$ & ">"
      End If
      DoEvents
    Wend
      While List1(1).ListCount > 0
        nureinmalhier = nureinmalhier + 1
        t$ = List1(1).List(0)
        List1(1).RemoveItem 0
        n% = InStr(t$, "|"): id$ = Left$(t$, n% - 1): t$ = Mid$(t$, n% + 1)
        n% = InStr(t$, "|"): k$ = Left$(t$, n% - 1): eml$ = Mid$(t$, n% + 1)
        kid.text = k$
        adrid.text = id$
        If Len(txtSendTo.text) = 0 Then
          txtSendTo.text = "<" & eml$ & ">"
        Else
          txtSendTo.text = txtCC.text + ",<" + eml$ & ">"
        End If
        DoEvents
      Wend

'  If List1(0).ListCount > 0 Then
'    t$ = List1(0).List(0)
'    List1(0).RemoveItem 0
'    n% = InStr(t$, "|"): id$ = Left$(t$, n% - 1): t$ = Mid$(t$, n% + 1)
'    n% = InStr(t$, "|"): k$ = Left$(t$, n% - 1): eml$ = Mid$(t$, n% + 1)
'    txtSendTo.Text = eml$
'    kid.Text = k$
'    adrid.Text = id$
'    DoEvents
'  End If
  If trm(adrid.text) <> "" Then
    If adrid.text <> "-1" Then
      'bei faxan gehts nochmal in dochist
      'Form1.sqlqry "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff) values(NULL,'" & adrid.text & "','" & kid.text & "','EMail','" & Date & " " & Time & "','" & Form1.getuserid() & "','" + txtMessageSubject.text + "')"
      txt2send = bkmks(txtMessageText.text, adrid.text, kid.text)
      volltext$ = txt2send
      volltext$ = Chr$(13) + Chr$(10) + volltext$ + Chr$(13) + Chr$(10) + transe("Anhänge") + ":"
      If lstAttachments.ListCount > 0 Then
        For i% = 0 To lstAttachments.ListCount - 1
          volltext$ = volltext$ + Chr$(13) + Chr$(10) + strrepl(lstAttachments.List(i%), "\", "\\")
        Next i%
      End If
      volltext$ = Chr$(13) + Chr$(10) + volltext$ + Chr$(13) + Chr$(10) + transe("Dieses Memo bedeutet nicht, dass die Mail den Empfänger erreicht hat.")
      volltext$ = Chr$(13) + Chr$(10) + volltext$ + Chr$(13) + Chr$(10) + transe("Sendeprotokoll:")
      While lblStatus.ListCount > 0
        lblStatus.ListIndex = 0
        volltext$ = volltext$ + Chr$(13) + Chr$(10) + form1.repl1310rtf(lblStatus.List(i%))
        lblStatus.RemoveItem 0
      Wend
      Call form1.memonoshow
'      Call form1.faxan(adrid.text, kid.text, form1.meinememovorlage(), transe("Email geschrieben") + ": " + txtMessageSubject.text, volltext$, "", "defaultname")
    End If
  End If

  o% = FreeFile
  On Error Resume Next
  Open mbox$ + "\lock.lck" For Output As #o%
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then Close #o%
  While trm(txtSendTo) <> ""
  lblStatus.AddItem "Connecting ..."
  lblStatus.ListIndex = lblStatus.ListCount - 1
  Select Case txtServer
    Case "OUTLOOK"
' Start Outlook.
 ' If it is already running, you'll use the same instance...
   Dim olApp As Outlook.Application
   Set olApp = CreateObject("Outlook.Application")
 
 ' Logon. Doesn't hurt if you are already running and logged on...
   Dim olNs As Outlook.NameSpace
   Set olNs = olApp.GetNamespace("MAPI")
   olNs.Logon

 ' Send a message to your new contact.
   Dim olMail As Outlook.MailItem
   Set olMail = olApp.CreateItem(olMailItem)
 ' Fill out & send message...
   olMail.To = txtSendTo
   olMail.Subject = txtMessageSubject
   mailusefont$ = strrepl(form1.getusersetting("mailfont", ""), ":", "'")
   If trm(adrid.text) <> "" Then
     If adrid.text <> "-1" Then
       c$ = bkmks(txtMessageText.text, adrid.text, kid.text)
     End If
   End If
   txt2send = mailusefont$ + strrepl(c$, vbCrLf, "<br>")
   If mailusefont$ <> "" Then txt2send = txt2send + "</font>"
   Do
     sndpart = stripimgpart1(txt2send)
     txt2send = stripimgpart2(txt2send)
     If sndpart <> "" Then olMail.HTMLBody = olMail.HTMLBody + sndpart
   Loop While sndpart <> ""
   For i% = 0 To lstAttachments.ListCount - 1
     olMail.Attachments.add lstAttachments.List(i%)
   Next i%

   olMail.Display

 ' Clean up...
   olNs.Logoff
   Set olNs = Nothing
   Set olMail = Nothing
   Set olApp = Nothing

    Case "thunderbirdremote"
' Start Outlook.
 ' If it is already running, you'll use the same instance...
 ' Fill out & send message...
   'To = txtSendTo
   'Subject = txtMessageSubject
   mailusefont$ = strrepl(form1.getusersetting("mailfont", ""), ":", "'")
   If trm(adrid.text) <> "" Then
     If adrid.text <> "-1" Then
       c$ = bkmks(txtMessageText.text, adrid.text, kid.text)
     End If
   End If
   txt2send = mailusefont$ + strrepl(c$, vbCrLf, "<br>")
   If mailusefont$ <> "" Then txt2send = txt2send + "</font>"
   l$ = ""
   Do
     sndpart = stripimgpart1(txt2send)
     txt2send = stripimgpart2(txt2send)
     If sndpart <> "" Then l$ = l$ + sndpart
   Loop While sndpart <> ""
   t$ = ""
   For i% = 0 To lstAttachments.ListCount - 1
     t$ = t$ + lstAttachments.List(1) + " "
   Next i%
   c$ = "xfeDoCommand(composeMessage,subject='" + txtMessageSubject + "',to='" + txtSendTo + "',body='" + l$ + "'"
   If t$ <> "" Then c$ = c$ + ",attachment='" + t$ + "'"
   c$ = c$ + ")"
Debug.Print form1.getusersetting("sendmail", "") + " """ + c$ + """"
   X = Shell(form1.getusersetting("sendmail", "") + " """ + c$ + """", 1)

    Case "NETSCAPE47":
      o% = FreeFile
'      mbox$ = strrepl(form1.getusersetting("Mozillaprofile"), """", "")
'      mbox$ = mbox$ & "\Mail\Local Folders\Drafts"
      mbox$ = strrepl(DirName(form1.getusersetting("netscape47inbox", "")) & "\Drafts", """", "")
      On Error Resume Next: Kill mbox$ + ".msf": On Error GoTo 0
      Open mbox$ For Append As #o%
      Print #o%, "From - " & Date & " " & Time
      Print #o%, "Message-ID: <" + strrepl(datum2sql(Date), "-", "") + strrepl(Time, ":", "") + "." + strrepl(strrepl(trm(Rnd), ",", ""), ".", "") + ">"
      Print #o%, "Date: " & Date & " " & Time
      Print #o%, "From: " & form1.getusersetting("Name") & "<" & form1.getuseremail(form1.getuserid) & ">"
      Print #o%, "To: " & txtSendTo
      l$ = "CC: " & txtCC: Print #o%, l$
      l$ = "BCC: " & txtBCC: Print #o%, l$
      Print #o%, "Subject: " & txtMessageSubject
      If htmldraft Then
        Print #o%, "Content-Type: text/html; charset=ISO-8859-15"
        Print #o%, "Content-Transfer-Encoding: 7bit"
        Print #o%, "<html><body>"
      End If
      Print #o%,
      If htmldraft Then
        txt2send = strrepl(txt2send, vbCrLf, "<br>")
        Do
            sndpart = stripimgpart1(txt2send)
            txt2send = stripimgpart2(txt2send)
            If sndpart <> "" Then Print #o%, sndpart;
        Loop While sndpart <> ""
        Print #o%, "</body></html>"
      Else
        Print #o%, txt2send
      End If
      Print #o%,
      Print #o%,
      Print #o%, "."
      Close #o%
   
   Case Else
      txtServer.text = "dir:Outbox"
      mbox$ = form1.mylocaldatadir() + "\mail"
      On Error Resume Next
      MkDir mbox$
      On Error GoTo 0
      mbox$ = mbox$ + "\outbox"
      DoEvents
      On Error Resume Next
      MkDir mbox$
      On Error GoTo 0
      o% = FreeFile
      dtg$ = datum2sql(Date)
      p% = 0
      Do
        mboxfile = mbox$ + "\" + dtg$ + "-" + strrepl(trm(Time), ":", "") + "-" + trm(p%)
        optfile = mboxfile$ + ".aof"
        lockfile = mboxfile$ + ".amf.lck"
        mboxfile$ = mboxfile$ + ".amf"
        p% = p% + 1
      Loop Until nexist(mboxfile)
      o% = FreeFile: Open lockfile For Output As #o%: Close #o%
      o% = FreeFile
Call form1.dbg2f("quittung=" + trm(srcpt.value))
      Open optfile For Output As #o%
      Print #o%, "quittung="; trm(srcpt.value)
      Close #o%
Call form1.dbg2f("writing mailfile " + mboxfile)
      o% = FreeFile
      Open mboxfile For Output As #o%
      Print #o%, "From: " & form1.getusersetting("Name") & "<" & form1.getuseremail(form1.getuserid) & ">"
      Print #o%, "Message-ID: <" + strrepl(datum2sql(Date), "-", "") + strrepl(Time, ":", "") + "." + strrepl(strrepl(trm(Rnd), ",", ""), ".", "") + ">"
      Print #o%, "Date: " & GetFormattedTime()
      Print #o%, "To: " & txtSendTo
      bcc$ = form1.getusersetting("AutoBCC", "")
      If txtBCC.text <> "" Then bcc$ = txtBCC.text
      For i% = 0 To List1(2).ListCount - 1
        l$ = List1(2).List(i%)
        l$ = cut_d2bis(l$, "|")
        l$ = cut_d2bis(l$, "|")
        If bcc$ <> "" Then bcc$ = bcc$ + ","
        bcc$ = bcc$ + l$
      Next i%
      lcc$ = ""
      If txtCC.text <> "" Then lcc$ = txtCC.text
      For i% = 0 To List1(1).ListCount - 1
        l$ = List1(1).List(i%)
        l$ = cut_d2bis(l$, "|")
        l$ = cut_d2bis(l$, "|")
        If lcc$ <> "" Then lcc$ = lcc$ + ","
        lcc$ = lcc$ + l$
      Next i%
      If lcc$ <> "" Then Print #o%, "CC: " + lcc$
      Print #o%, "Subject: " & txtMessageSubject
        
        
        bndry$ = "000_" + GUID()
        Print #o%, "Content-Type: multipart/mixed; boundary=""" & bndry$ & """"
'        Print #o%, "Content -Disposition: InLine"
         Print #o%, "This is a multi-part message in MIME format.";
        Print #o%,
        Print #o%, "--" & bndry$
        If htmldraft Then
          Print #o%, "Content-Type: text/html; charset=ISO-8859-15"
          Print #o%, "Content-Transfer-Encoding: 7bit" & vbCrLf
          Print #o%,
          Print #o%, "<html><body>"
          txt2send = strrepl(txt2send, vbCrLf, "<br>")
          Do
            sndpart = stripimgpart1(txt2send)
            txt2send = stripimgpart2(txt2send)
            If sndpart <> "" Then Print #o%, sndpart;
          Loop While sndpart <> ""
          Print #o%, "</body></html>"
          Print #o%,
        Else
          Print #o%, "Content-Type: text/plain; charset=us-ascii"
          Print #o%, "Content -Disposition: InLine"
          Print #o%,
          Print #o%, txtMessageText
        End If
        Print #o%,
        For i% = 0 To yattach.ListCount - 1
          lblStatus.AddItem Date & " " & Time & " Attachment" & i% + 1 & " wird geschrieben"
          lblStatus.ListIndex = lblStatus.ListCount - 1
          DoEvents
          ifn$ = Mid$(yattach.List(i%), InStr(yattach.List(i%), "|") + 1)
Call form1.dbg2f("adding attachment " + ifn$)
          If LCase(FileExtension(ifn$)) <> "b64" Then
            Print #o%, "--" + bndry$
          End If
          addhd$ = ""
          orgfn = FileName(cut_d1(yattach.List(i%), "|"))
          If LCase(FileExtension(ifn$)) = "tmp" Then addhd$ = "Content-ID: <" + mkkey(5) + strrepl(orgfn, " ", "") + ">"
          o1% = FreeFile
          Open ifn$ For Input As #o1%
          While Not EOF(o1%)
            DoEvents
            Line Input #o1%, strMess
            If InStr(LCase(strMess), "content-disposition") > 0 Then
              strMess = strMess + " filename=""" + orgfn + """;"
            End If
            If strMess = "" And addhd$ <> "" Then
              Print #o%, addhd$
            End If
            Print #o%, strMess
          Wend
          Close #o1%
        Next i%
        Print #o%,
        Print #o%, "--" & bndry$ & "--"
        Print #o%,


      Close #o%
      If bcc$ <> "" Then
        o% = FreeFile
        Open mboxfile + ".bcc" For Output As #o%
        Print #o%, bcc$;
        Close #o%
      End If
      rrr = 0
      On Error Resume Next
      Kill lockfile
      rrr = Err
      On Error GoTo 0
Call form1.dbg2f("lock removed=" + trm(rrr))
      On Error Resume Next
  End Select
  If anadr$ <> "" Then anadr$ = anadr$ + ","
  anadr$ = anadr$ + trm(txtSendTo)
  txtSendTo = ""
  txtCC.text = ""
  DoEvents
  Wend
  On Error Resume Next
  Kill mbox$ + "\lock.lck"
  On Error GoTo 0
  If adrid.text <> "-1" Then Call form1.apmaillog(adrid.text, kid.text, anadr$, txtMessageSubject.text)
  anadr = ""
DoEvents
Wend
cmdStop.Visible = False
cmdSend.Visible = True
txtSendTo.text = ""
txtMessageSubject.text = ""
txt2send = ""
While lstAttachments.ListCount > 0
  Call detachfile(lstAttachments.List(0))
  lstAttachments.RemoveItem 0
Wend
yattach.Clear
Call form1.signaturinclude
MousePointer = 0
If Not mailstopped Then
  If List4.ListIndex >= 0 And Not mailstopped Then
    Call Command5_Click
  End If
Else
  While merk0.ListCount > 0
    List1(0).AddItem merk0.List(0)
    merk0.RemoveItem 0
  Wend
  txtMessageText.text = merk0t
  txtMessageSubject.text = merk0b
End If
If txtServer = "NETSCAPE47" Then
  On Error Resume Next
  X = Shell(form1.getusersetting("Mailclient"), 1)
  On Error GoTo 0
End If
lblStatus.AddItem "Bye."
lblStatus.ListIndex = lblStatus.ListCount - 1
xattach.Clear

End Sub

Private Sub cmdStop_Click()
'd2infile = "smtp": d2insub = "cmdStop_Click"
merk0.Clear
While List1(0).ListCount > 0
  merk0.AddItem List1(0).List(0)
  List1(0).RemoveItem 0
Wend
merk0t = txtMessageText.text
merk0b = txtMessageSubject.text
txtSendTo.text = ""
mailstopped = True

End Sub

Public Sub Command1_Click()
Dim i%
'd2infile = "smtp": d2insub = "Command1_Click"
txtSendTo.text = ""
txtMessageText = ""
Hide
For i% = 0 To 2
  While List1(i%).ListCount > 0
    List1(i).RemoveItem 0
  Wend
Next i%
Unload emailadrselect
Unload Me
End Sub

Private Sub Command10_Click()
Dim fname As String, o%, l$

    fname = form1.s0dir() + "\*.csv"
    On Error Resume Next
    With cdlg1
    'Bei "Abbruch" Fehler raisen lassen:
    .CancelError = True
    'Suchpfad einstellen:
    .InitDir = DirName(fname)
    .FileName = FileName(fname)
    .DialogTitle = "Open ..."
    'und endlich den Dialog anzeigen:
    .ShowOpen

    'Auswertung:
    If Err = cdlCancel Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    fname = .FileName

    End With
    On Error GoTo 0

List1(0).Clear
o% = FreeFile
Open fname For Input As #o%
While Not EOF(o%)
  Line Input #o%, l$
  List1(0).AddItem "-1|" + l$
Wend
Close #o%

End Sub

Private Sub Command18_Click()
'd2infile = "smtp": d2insub = "Command18_Click"
Call form1.handbuchcall("13-Email.htm")

End Sub

Public Sub Command2_Click()

'd2infile = "smtp": d2insub = "Command2_Click"
Load emailadrselect
emailadrselect.Text1.text = ""
Call emailadrselect.SetFocus
Call emailadrselect.callbackto("smtp")

End Sub

Private Sub Command3_Click()
Dim i As Integer, p As Integer

'd2infile = "smtp": d2insub = "Command3_Click"
p = List5.ListIndex
If p < 0 Then Exit Sub
If List1(p).ListCount <= 0 Then Exit Sub
For i = List1(p).ListCount - 1 To 0 Step -1
  If List1(p).Selected(i) Then List1(p).RemoveItem i
Next i

End Sub

Public Sub Command4_Click()
Dim fn$, o%, l$, rrr, j%

'd2infile = "smtp": d2insub = "Command4_Click"
fn$ = form1.myuniqueemlname()
o% = FreeFile
If trm(txtSendTo.text) <> "" Then
  List1(0).AddItem adrid.text & "|" & kid.text & "|" & txtSendTo.text
  txtSendTo.text = ""
End If
lastsavename = fn$
Open fn$ For Output As #o%
For j% = 0 To 2
  While List1(j%).ListCount > 0
    Print #o%, List1(j%).List(0)
    List1(j%).RemoveItem 0
    DoEvents
  Wend
  Print #o%, "."
Next j%
Print #o%, txtMessageSubject.text
txtMessageSubject.text = ""
While lstAttachments.ListCount > 0
  Print #o%, lstAttachments.List(0)
  Call detachfile(lstAttachments.List(0))
  lstAttachments.RemoveItem 0
  DoEvents
Wend
yattach.Clear
Print #o%, "."
Print #o%, txtMessageText.text
txtMessageText.text = ""
Close #o%
Call rlist4
Call form1.signaturinclude
End Sub

Private Sub Command5_Click()
Dim fn$

'd2infile = "smtp": d2insub = "Command5_Click"
If List4.ListIndex < 0 Then Exit Sub
fn$ = form1.mylocaldatadir() & "\" & List4.List(List4.ListIndex)
If exist(fn$) Then
  Kill fn$
  List4.RemoveItem List4.ListIndex
End If
End Sub

Private Sub Command6_Click()
Dim l$, o%, rrr, j%
'd2infile = "smtp": d2insub = "Command6_Click"
txtMessageText.text = ""
List4.ListIndex = -1
For j% = 0 To 2: List1(j%).Clear: Next j%
lstAttachments.Clear
yattach.Clear
txtMessageSubject.text = ""
l$ = " "
txtMessageText.text = ""
Call form1.signaturinclude
List4.ListIndex = -1
End Sub

Private Sub Command7_Click()
'd2infile = "smtp": d2insub = "Command7_Click"
While List4.ListCount > 0
  List4.ListIndex = 0
  DoEvents
  Call cmdSend_Click
  'List4.RemoveItem 0
  DoEvents
Wend
End Sub


Private Sub Command8_Click()
Dim j%, sbj As String

'd2infile = "smtp": d2insub = "Command8_Click"
'sbj = txtMessageSubject.Text
Call Command4_Click
DoEvents
For j% = 0 To List4.ListCount - 1
  If InStr(lastsavename, List4.List(j%)) > 0 Then
    List4.ListIndex = j%
    DoEvents
    List4.ListIndex = -1
    Exit For
  End If
Next j%
For j% = 0 To 2: List1(j%).Clear: Next j%
'txtMessageSubject.Text = sbj
'txtSendTo.Text = form1.getuseremail(form1.getuserid())
txtSendTo.text = "ping@i4f.de"
txtMessageSubject = "Testmail von " + form1.getuseremail(form1.getuserid())
Call cmdSend_Click
DoEvents
For j% = 0 To List4.ListCount - 1
  If InStr(lastsavename, List4.List(j%)) > 0 Then
    List4.ListIndex = j%
    DoEvents
    Call Command5_Click
    Exit Sub
  End If
Next j%
End Sub

Private Sub Command9_Click(Index As Integer)
Dim f$, r$, tg$, s0 As Integer, add2tag$

If Command9(Index).Caption = "im&g" Then
  Call imginclude
  Exit Sub
End If
If Left(Command9(Index).Caption, 3) = "fon" Then
  Call Label6_Click
  Exit Sub
End If
  
r$ = txtMessageText.text: f$ = ""
s0 = txtMessageText.SelStart
If txtMessageText.SelStart > 0 Then
  f$ = Left(txtMessageText.text, txtMessageText.SelStart)
  r$ = Mid$(txtMessageText.text, txtMessageText.SelStart + 1)
End If
tg$ = strrepl(Command9(Index).Caption, "&", "")
txtMessageText.text = f$ + "<" + tg$ + add2tag$ + ">" + r$
txtMessageText.SelStart = s0 + Len(tg$) + 2
Call txtMessageText.SetFocus
Call txtMessageTextchg

End Sub

Private Sub eclient_Click()
Call form1.setmylastFormVar(Me.name, "eclient", trm(eclient.value))
End Sub

Private Sub Form_Load()
Dim r As ADODB.Recordset, dbpara$, klrv%, s%, j%, rrr
Dim htmldraft As Boolean

htmldraft = False
If form1.getusersetting("htmldraft", "ja") = "ja" Then htmldraft = True

tgcnt% = 0:
tags$(tgcnt%) = "b": Command9(tgcnt%).Caption = "&" + tags$(tgcnt%): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "big": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "small": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "i": Command9(tgcnt%).Caption = "&" + tags$(tgcnt%): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "table": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "tr": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "th": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "td": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "font": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "h1": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tags$(tgcnt%) = "img": Command9(tgcnt%).Caption = Left(tags$(tgcnt%), Len(tags$(tgcnt%)) - 1) + "&" + Right(tags$(tgcnt%), 1): tagsc$(tgcnt%) = Command9(tgcnt%).Caption: tgcnt% = tgcnt% + 1
tagchk.value = 1
If form1.getusersetting("smtptagtest", "ja") = "nein" Then tagchk.value = 0
If Not htmldraft Then
  For j% = 0 To tgcnt% - 1
   Command9(j%).Enabled = False
  Next j%
End If
Dim d2infile As String, d2insub As String
d2infile = "smtp": d2insub = "Form_Load"
axsResizer1.SaveControlPositions

s% = form1.myfontsize()
List4.Font.Size = s%
For j% = 0 To 2
  List1(j%).Font.Size = s%
  List1(j%).Font.Size = s%
Next j%
txtSendTo.Font.Size = s%
txtMessageSubject.Font.Size = imax(10, s%)
txtMessageText.Font.Size = imax(10, s%)
lblStatus.Font.Size = s%
lstAttachments.Font.Size = s%
List5.Clear
List5.AddItem transe("An")
List5.AddItem transe("CC")
List5.AddItem transe("BCC")
List5.ListIndex = 0
xattach.Clear
yattach.Clear
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM gruppennamen", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
List2.Clear
While Not r.EOF
  List2.AddItem r!gid
  r.MoveNext
Wend
r.Close
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM adressgruppenindex", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

List3.Clear
While Not r.EOF
  List3.AddItem r!id
  r.MoveNext
Wend
eclient.value = 0
txtServer.text = form1.mymailoutserver()
If txtServer.text = "OUTLOOK" Or txtServer.text = "NETSCAPE" Then
  eclient.value = 1
Else
  txtServer.text = "dir:Outbox"
End If
klrv% = Val(form1.mylastFormVar(Me.name, "eclient", "0"))
If klrv% <> 0 Then klrv% = 1
eclient.value = klrv%

txtMailFrom.text = form1.mymailaddress()
klrv% = Val(form1.mylastFormVar(Me.name, "ab4s", "0"))
If klrv% <> 0 Then klrv% = 1
askb4send.value = klrv%
Call form1.dbg2f("smtp.Form_Load:setting replyto")
'replyto.Text = form1.getusersetting("Name") & "<" & form1.getuseremail(form1.getuserid()) & ">"
replyto.text = form1.getuseremail(form1.getuserid())
Call form1.dbg2f("smtp.Form_Load:setting replyto done")
smtp.Caption = transe("Email senden")
Command18.ToolTipText = transe("Hilfeseite öffnen")
cmdSend.ToolTipText = transe("Senden")
Command4.ToolTipText = transe("Speichern")
Command7.Caption = transe("all&e senden")
Command6.Caption = transe("&NEU")
cmdRemove.ToolTipText = transe("Anhang entfernen")
lstAttachments.ToolTipText = transe("Sie können <Drag & Drop> benutzen um Dateien anzufügen")
cmdAdd.ToolTipText = transe("Anhang hinzufügen")
Label20.Caption = transe("Reply-To:")
Label19.Caption = transe("Senden bestätigen")
Label18.Caption = transe("ext. Client")
Label16.Caption = transe("gepeicherte Mail:")
Label15.Caption = transe("aus Adressen")
Label14.Caption = transe("Interne")
Label13.Caption = transe("Meldungen")
Label12.Caption = transe("Password")
Label11.Caption = transe("Username")
Label10.Caption = transe("blinde Kopie an")
Label9.Caption = transe("Verteiler")
Label8.Caption = transe("MessageHTML")
Label7.Caption = transe("Anhänge")
cmdSend.Caption = transe("absenden")

Label4.Caption = transe("Betreff")
Label3.Caption = transe("An")
Label2.Caption = transe("Absender")
Label1.Caption = transe("Server")
Label6.ForeColor = RGB(0, 0, 0)
Call rmarken
Show

Call rlist4
If txtServer.text = "NETSCAPE47" Or txtServer.text = "OUTLOOK" Then
  cmdAdd.Enabled = False
End If
smtp.Top = form1.mylasttop(Me.name)
smtp.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
End Sub

Private Sub Form_Resize()
'd2infile = "smtp": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "smtp": d2insub = "Form_Unload"
Hide
On Error Resume Next
Kill form1.s0dir() & "\debug2file_" & form1.getuserid() & "_smtp.txt"
On Error GoTo 0

On Error GoTo exuldx1
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuldx1:
On Error GoTo 0
End Sub

Private Sub Label1_DblClick()
'd2infile = "smtp": d2insub = "Label1_DblClick"
Load einstellungen

End Sub

Private Sub Label19_Click()
'd2infile = "smtp": d2insub = "Label19_Click"
If askb4send.value = 0 Then
  askb4send.value = 1
Else
  askb4send.value = 0
End If

End Sub

Private Sub Label2_DblClick()
'd2infile = "smtp": d2insub = "Label2_DblClick"
Load einstellungen

End Sub

Private Sub Label22_Click()
Call tagchk_Click
End Sub

Private Sub Label6_Click()
Load colorsel
colorsel.SetFocus
colorsel.updc (Label6.BackColor)
Timer3.Enabled = True
Timer3.Interval = 500
While Timer3.Enabled: DoEvents: Wend

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'd2infile = "smtp": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then Call Command3_Click
End Sub

Private Sub List2_DblClick()
Dim eml$, g$, r As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "smtp": d2insub = "List2_DblClick"
g$ = List2.List(List2.ListIndex)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM benutzergruppen where groupid='" + g$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  eml$ = form1.getuseremail(r!userid)
  If eml$ <> "" Then List1(List5.ListIndex).AddItem "-1|-1|" + eml$
  r.MoveNext
Wend
End Sub

Public Sub List3_DblClick()
Dim eml$, g$, r As ADODB.Recordset, c As String, rrr

Dim d2infile As String, d2insub As String
d2infile = "smtp": d2insub = "List3_DblClick"
If List3.ListIndex < 0 Then Exit Sub
g$ = List3.List(List3.ListIndex)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
c = "SELECT * FROM adressgruppen where grpid='" + g$ + "'"
rrr = form1.adoopen(r, c, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  eml$ = ""
  If Not IsNull(r!kid) And r!kid <> "-1" Then eml$ = form1.getkontaktemailbyid(r!kid)
  If eml$ = "" Then eml$ = form1.getemailbyid(r!adressid)
  If eml$ <> "" Then List1(List5.ListIndex).AddItem r!adressid & "|" & r!kid & "|" & eml$
  r.MoveNext
Wend

End Sub

Private Sub List4_Click()
Dim fn$, o%, l$, j%

'd2infile = "smtp": d2insub = "List4_Click"
If List4.ListIndex < 0 Then Exit Sub
txtMessageText.text = ""
For j% = 0 To 2: List1(j%).Clear: Next j%
While lstAttachments.ListCount > 0
  Call smtp.detachfile(lstAttachments.List(0))
  lstAttachments.RemoveItem 0
Wend
yattach.Clear
fn$ = form1.mylocaldatadir() & "\" & List4.List(List4.ListIndex)
o% = FreeFile
Open fn$ For Input As #o%
For j% = 0 To 2
  l$ = "x"
  While Not EOF(o%) And trm(l$) <> "."
    Line Input #o%, l$
    If l$ <> "." Then List1(j%).AddItem l$
  Wend
Next j%
If Not EOF(o%) Then
  Line Input #o%, l$: txtMessageSubject.text = l$
  l$ = " "
  While Not EOF(o%) And l$ <> "."
    Line Input #o%, l$
    If l$ <> "." Then Call attachfile(l$)
  Wend
End If
l$ = " "
While Not EOF(o%)
  Line Input #o%, l$
  If l$ <> "." Then txtMessageText.text = txtMessageText.text & l$ & Chr$(13) & Chr$(10)
Wend
Close #o%

End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, fn$

If KeyCode = 8 Or KeyCode = 46 Then
  i% = List4.ListIndex
  If i% < 0 Then Exit Sub
  fn$ = form1.mylocaldatadir() + "\" + List4.List(i%)
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
  List4.RemoveItem i%
End If

End Sub

Private Sub List5_Click()
Dim i%, j%

'd2infile = "smtp": d2insub = "List5_Click"
i% = List5.ListIndex
If i% < 0 Then Exit Sub

Label9.Caption = List5.List(i%)
For j% = 0 To 2
  If j% = i% Then
    List1(j).Visible = True
  Else
    List1(j).Visible = False
  End If
Next j%
End Sub

Private Sub lstAttachments_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%
'd2infile = "smtp": d2insub = "lstAttachments_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  i% = lstAttachments.ListIndex
  If i% < 0 Then Exit Sub
  Call detachfile(lstAttachments.List(i%))
  lstAttachments.RemoveItem i%
End If

End Sub

Private Sub lstAttachments_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fn$, nf%, rrr

'd2infile = "smtp": d2insub = "lstAttachments_OLEDragDrop"
If Data.GetFormat(vbCFFiles) Then
  nf% = 1
  Do
    On Error Resume Next
    fn$ = Data.Files(nf%)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      Call attachfile(fn$)
      DoEvents
    End If
    nf% = nf% + 1
  Loop Until rrr <> 0
End If
End Sub

Private Sub marken_Click()
Dim M$, l$, s0 As Integer, f$, txg$

M$ = trm(marken.text)
If M$ <> "" Then
  l$ = txtMessageText.text: f$ = ""
  s0 = txtMessageText.SelStart
  If txtMessageText.SelStart > 0 Then
    f$ = Left(txtMessageText.text, txtMessageText.SelStart)
    l$ = Mid$(txtMessageText.text, txtMessageText.SelStart + 1)
  End If
  txg$ = "<!--bkmkstart-->{" + M$ + "}<!--bkmkend-->"
  txtMessageText.text = f$ + txg$ + l$
  txtMessageText.SelStart = s0 + Len(txg$)
  Call txtMessageText.SetFocus
  Call txtMessageTextchg
End If

End Sub

Private Sub prvw_Click()
Dim fn$, o%, brw$, X

fn$ = form1.mylocaldatadir() + "/mailpreview.html"
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, "<html><body>"
Print #o%, strrepl(txtMessageText, vbCrLf, "<br>")
Print #o%, "</body></html>"
Close #o%

      Unload frmBrowser
      DoEvents
      brw$ = form1.UseBrowser()
      If brw$ <> "" And Not nexist(brw$) Then
        X = Shell(brw$ & " file:///" + strrepl(fn$, "\", "/"), 1)
      Else
        frmBrowser.StartingAddress = "file:///" + strrepl(fn$, "\", "/")
        Load frmBrowser
        frmBrowser.cboAddress.Visible = False
        frmBrowser.lblAddress.Visible = False
      End If


End Sub

Private Sub tagchk_Click()

If tagchk.value = 0 Then
  Call form1.setusersetting("smtptagtest", "nein")
Else
  Call form1.setusersetting("smtptagtest", "ja")
End If
End Sub

Private Sub Timer1_Timer()
Dim tg$

'd2infile = "smtp": d2insub = "Timer1_Timer"
Call form1.dbg2f("smtp Timer1 start")
tg$ = fselect.fqn.text
If tg$ = "" Then Exit Sub
fselect.fqn.text = ""
Timer1.Enabled = False
Call attachfile(tg$)
Call form1.dbg2f("smtp Timer1 exit")

End Sub

Public Sub callback(vid$, kid$, eml$)

'd2infile = "smtp": d2insub = "callback"
List1(List5.ListIndex).AddItem vid$ & "|" & kid$ & "|" & eml$

End Sub

Sub rlist4()
Dim fn$
'd2infile = "smtp": d2insub = "rlist4"
List4.Clear

fn$ = Dir(form1.mylocaldatadir() & "\*.apm")
While fn$ <> ""
  List4.AddItem fn$
  fn$ = Dir
Wend

End Sub

Sub rmarken()

marken.Clear
marken.AddItem ""
marken.AddItem "Anrede"
marken.AddItem "Abrede"
marken.AddItem "Postanrede"
marken.AddItem "Name"
marken.AddItem "Strasse"
marken.AddItem "PLZORT"
marken.AddItem "PLZ"
marken.AddItem "Land"
marken.AddItem "Ort"
marken.AddItem "Tel"
marken.AddItem "Fax"
marken.AddItem "Handy"
marken.AddItem "Hinweise"

End Sub

Private Sub Timer2_Timer()
'd2infile = "smtp": d2insub = "Timer2_Timer"
If Val(Label17.Caption) = List1(List5.ListIndex).ListCount Then Exit Sub
Label17.Caption = List1(List5.ListIndex).ListCount & " Empfänger"
End Sub
Public Sub attachfile(tgin$)
Dim fn$, mt$, ext$, o%, d$, fin$, rrr, tg$

'd2infile = "smtp": d2insub = "attachfile"
tg$ = tgin$
If tg$ = "" Then Exit Sub
Call form1.dbg2f("attaching " + tg$)
fin$ = FileName(tg$)
o% = FreeFile
On Error Resume Next
Open tg$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Close #o%
Else
  tg$ = form1.filenamekurz(tg$)
End If
If tg$ = "" Then
  MsgBox "cannot attach " + tgin$ + vbCrLf + "possibly an umlaut that cannot be handled."
  Exit Sub
End If
lstAttachments.AddItem tg$
ext$ = FileExtension(tg$)
If InStr(LCase(form1.getusersetting("mailclient", "")), "outlook") > 0 And eclient.value = 1 Then
  Call form1.dbg2f("using outlook as client, not attaching here.")
  Exit Sub
End If
mt$ = form1.mimetype(ext$): If Right$(mt$, 1) <> ";" Then mt$ = mt$ & ";"
d$ = form1.s0dir() & "\tmp"
On Error Resume Next
MkDir d$
On Error GoTo 0
fn$ = form1.s0dir() & "\tmp\" & GUID() & ".tmp"
Call form1.dbg2f("converting to: " + fn$)
yattach.AddItem tg$ & "|" & fn$
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, "Content-Type: " & mt$
Print #o%, " name=""" & fin$ & """"
Print #o%, "Content-Transfer-Encoding: base64"
Print #o%, "Content-Disposition: inline;"
Print #o%, "  name=""" & fin$ & """"
Print #o%,
Close #o%
fn$ = form1.s0dir() & "\tmp\" & GUID() & ".b64"
yattach.AddItem tg$ & "|" & fn$
Call EncodeFileB64(tg$, fn$)
If Not nexist(fn$) Then
  Call form1.dbg2f("file exists: " + fn$)
Else
  Call form1.dbg2f("FILE DOES NOT EXIST: " + fn$)
End If
End Sub

Public Sub detachfile(tg$)
Dim i%

'd2infile = "smtp": d2insub = "detachfile"
For i% = yattach.ListCount - 1 To 0 Step -1
  If InStr(yattach.List(i%), tg$ & "|") = 1 Then
    On Error Resume Next
    Kill Mid$(yattach.List(i%), InStr(yattach.List(i%), "|") + 1)
    On Error GoTo 0
    yattach.RemoveItem i%
  End If
Next i%

End Sub

Private Sub Timer3_Timer()
Dim c As Long, tg$, txg$
Dim w As Long, r As Long, g As Long, b As Long, f$, l$, s0 As Integer

c = form1.getcolorselected()

If c < -10 Then Exit Sub
Timer3.Enabled = False
If c < 0 Then Exit Sub
Label6.ForeColor = c
Call form1.dbg2f("smtp Timer3 start")
b = c / 65536
w = c Mod 65536
g = w / 256
r = w Mod 256
  
l$ = txtMessageText.text: f$ = ""
s0 = txtMessageText.SelStart
If txtMessageText.SelStart > 0 Then
  f$ = Left(txtMessageText.text, txtMessageText.SelStart)
  l$ = Mid$(txtMessageText.text, txtMessageText.SelStart + 1)
End If
tg$ = strrepl(Command9(8).Caption, "&", "")
txg$ = "<font color=#" + hex2(r) + hex2(g) + hex2(b) + ">"
txtMessageText.text = f$ + txg$ + "</font>" + l$
txtMessageText.SelStart = s0 + Len(txg$)
Call txtMessageText.SetFocus
Call txtMessageTextchg
Call form1.dbg2f("smtp Timer3 exit")
End Sub

Private Sub txtMessageTextchg()
Dim tgflg%(19), i%, j%, hier$

If tagchk.value = 0 Then Exit Sub
'Debug.Print txtMessageText.SelStart, txtMessageText.SelLength
For i% = 1 To 9: tgflg(i%) = 0: Next i%
For i% = 1 To txtMessageText.SelStart - 1
  For j% = 0 To tgcnt% - 1
    If Len(tags$(j%)) + 2 <= txtMessageText.SelStart Then
      hier$ = Mid$(txtMessageText.text, i%, Len("<" + tags$(j%) + ">"))
      If hier$ = "<" + tags$(j%) + ">" Then
        tgflg%(j%) = tgflg%(j%) + 1
'Debug.Print tags$(j%) + ": " + trm(tgflg%(j%))
      End If
    End If
    If Len(tags$(j%)) + 3 < txtMessageText.SelStart Then
      hier$ = Mid$(txtMessageText.text, i%, Len("</" + tags$(j%) + ">"))
      If hier$ = "</" + tags$(j%) + ">" Then
        tgflg%(j%) = tgflg%(j%) - 1
'Debug.Print tags$(j%) + ": " + trm(tgflg%(j%))
      End If
    End If
  Next j%
Next i%
For j% = 0 To tgcnt% - 1
  If tgflg(j%) > 0 Then
    Command9(j%).Caption = "/" + tagsc$(j%)
  Else
    Command9(j%).Caption = tagsc$(j%)
  End If
Next j%
End Sub

Private Sub txtMessageText_Click()
Call txtMessageTextchg
End Sub

Private Sub txtMessageText_KeyUp(KeyCode As Integer, Shift As Integer)
Call txtMessageTextchg
End Sub
