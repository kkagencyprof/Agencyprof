VERSION 5.00
Begin VB.Form einstellungen 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Benutzer-Einstellungen AgencyProf"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "einstellungen.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "clear"
      Height          =   255
      Left            =   5280
      TabIndex        =   128
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox shbl 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   127
      ToolTipText     =   "schwarze Liste anzeigen"
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   125
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   124
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1200
      Picture         =   "einstellungen.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   122
      ToolTipText     =   "Schwarze Liste löschen"
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   121
      Top             =   8160
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   120
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&+"
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
      Left            =   9360
      TabIndex        =   118
      Top             =   7200
      Width           =   255
   End
   Begin VB.ListBox List5 
      Height          =   1035
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   117
      Top             =   7440
      Width           =   3615
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
      Left            =   120
      TabIndex        =   115
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   1680
      Width           =   495
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   113
      Top             =   7440
      Width           =   3015
   End
   Begin VB.TextBox pin 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8040
      PasswordChar    =   "*"
      TabIndex        =   110
      ToolTipText     =   "Passwort, mit dem die Mailpasswörter verschlüsselt werden"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Speichern als"
      Height          =   255
      Left            =   6000
      TabIndex        =   109
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox svas 
      Height          =   285
      Left            =   7560
      TabIndex        =   108
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9240
      Picture         =   "einstellungen.frx":069C
      Style           =   1  'Grafisch
      TabIndex        =   107
      Top             =   4920
      Width           =   375
   End
   Begin VB.ListBox popl 
      Height          =   1035
      Left            =   7560
      TabIndex        =   106
      Top             =   3810
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "einstellungen.frx":0B8C
      Style           =   1  'Grafisch
      TabIndex        =   104
      ToolTipText     =   "Dieses Formular schliessen"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   6000
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   28
      Left            =   6480
      TabIndex        =   102
      Text            =   "Text2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   27
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   27
      Left            =   6480
      TabIndex        =   100
      Text            =   "Text2"
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   6000
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   26
      Left            =   6480
      TabIndex        =   98
      Text            =   "Text2"
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   6000
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   25
      Left            =   6480
      TabIndex        =   96
      Text            =   "Text2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   8520
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   24
      Left            =   8400
      TabIndex        =   94
      Text            =   "Text2"
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   23
      Left            =   2520
      TabIndex        =   92
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   1560
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   7080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   1560
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   22
      Left            =   2520
      TabIndex        =   90
      Text            =   "Text2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Signatur"
      Height          =   255
      Left            =   3480
      Style           =   1  'Grafisch
      TabIndex        =   89
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   21
      Left            =   2880
      TabIndex        =   87
      Text            =   "Text2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   6240
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   20
      Left            =   6480
      TabIndex        =   85
      Text            =   "Text2"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "weitere Einstellungen"
      Height          =   495
      Left            =   8040
      TabIndex        =   83
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   19
      Left            =   8400
      TabIndex        =   80
      Text            =   "Text2"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   8520
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   6240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   18
      Left            =   6480
      TabIndex        =   78
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   17
      Left            =   2520
      TabIndex        =   76
      Text            =   "Text2"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   1560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   6360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   1560
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6000
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   16
      Left            =   2520
      TabIndex        =   74
      Text            =   "Text2"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   15
      Left            =   2520
      TabIndex        =   72
      Text            =   "Text2"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   1560
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   14
      Left            =   6480
      TabIndex        =   69
      Text            =   "Text2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   6240
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   1560
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   13
      Left            =   2520
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9000
      Picture         =   "einstellungen.frx":0DDC
      Style           =   1  'Grafisch
      TabIndex        =   66
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8040
      Picture         =   "einstellungen.frx":12CC
      Style           =   1  'Grafisch
      TabIndex        =   65
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
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
      Left            =   7560
      TabIndex        =   62
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
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
      Left            =   7560
      TabIndex        =   61
      Top             =   720
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   59
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   6240
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   12
      Left            =   6480
      TabIndex        =   57
      Text            =   "Text2"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   56
      Text            =   "Text2"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   1560
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   6240
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5040
      Picture         =   "einstellungen.frx":165E
      Style           =   1  'Grafisch
      TabIndex        =   49
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4560
      Picture         =   "einstellungen.frx":1B4E
      Style           =   1  'Grafisch
      TabIndex        =   48
      Top             =   1800
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   47
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   46
      Text            =   "Text2"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   44
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   42
      Text            =   "Text2"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   40
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "einstellungen.frx":1EE0
      Style           =   1  'Grafisch
      TabIndex        =   36
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnereinstellungen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   126
      ToolTipText     =   "Einstellungen dieses Computers ändern"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4440
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Importiere von:"
      Height          =   255
      Left            =   4560
      TabIndex        =   123
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Line bottomline 
      Visible         =   0   'False
      X1              =   120
      X2              =   9720
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Importiere von:"
      Height          =   255
      Left            =   360
      TabIndex        =   119
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "mehr Userdaten:"
      Height          =   255
      Left            =   7920
      TabIndex        =   116
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "schwarze Liste:"
      Height          =   255
      Left            =   360
      TabIndex        =   114
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label nserv 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4800
      TabIndex        =   112
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7560
      TabIndex        =   111
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzer-Daten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   105
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   28
      Left            =   4680
      TabIndex        =   103
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   27
      Left            =   4680
      TabIndex        =   101
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   26
      Left            =   4680
      TabIndex        =   99
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   25
      Left            =   4680
      TabIndex        =   97
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   24
      Left            =   7200
      TabIndex        =   95
      Top             =   6735
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   23
      Left            =   240
      TabIndex        =   93
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   91
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   21
      Left            =   840
      TabIndex        =   88
      Top             =   2295
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   20
      Left            =   4680
      TabIndex        =   86
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ja/nein"
      Height          =   255
      Left            =   6600
      TabIndex        =   84
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzerkennung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   82
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   19
      Left            =   7200
      TabIndex        =   81
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   18
      Left            =   4680
      TabIndex        =   79
      Top             =   7095
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   77
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   75
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   73
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   4680
      TabIndex        =   70
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   68
      Top             =   5295
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gruppen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7920
      TabIndex        =   64
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ist Mitglied von"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   63
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   58
      Top             =   6375
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   55
      Top             =   4935
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "msec"
      Height          =   255
      Left            =   7080
      TabIndex        =   54
      Top             =   5415
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   53
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   51
      Top             =   4575
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   37
      Top             =   4215
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   35
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   34
      Top             =   1935
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   33
      Top             =   3495
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   32
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   31
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   30
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   29
      Top             =   855
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   495
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2655
      Left            =   720
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2175
      Left            =   4440
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "einstellungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uId$, nflds As Integer
Dim poplistok%, grantlock As Boolean

Sub rlist23()
Dim r As ADODB.Recordset, usrid$

d2infile = "einstellungen": d2insub = "rlist23"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.rlist23:" + "SELECT * FROM gruppennamen")
On Error Resume Next
r.Open "SELECT * FROM gruppennamen", form1.adoc, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
List3.Clear
While Not r.EOF
  List3.AddItem r!gid
  r.MoveNext
Wend
usrid$ = Text1(0).text
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.rlist23:" + "SELECT * FROM benutzergruppen where userid='" + usrid$ + "'")
r.Open "SELECT * FROM benutzergruppen where userid='" + usrid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
List2.Clear
While Not r.EOF
  List2.AddItem r!groupid
  r.MoveNext
Wend

End Sub
Sub rlist1()
Dim rtmp As ADODB.Recordset, i As Integer

d2infile = "einstellungen": d2insub = "rlist1"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.rlist1:" + "SELECT * FROM benutzerdaten")
rtmp.Open "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly

List1.Clear
While Not rtmp.EOF
  If form1.getuserid() = "www" Or _
     form1.getuserid() = rtmp!id Or _
     form1.getusersettingfromuser(rtmp!id, "appasswort", "") = "" Then
    List1.AddItem rtmp!id
  End If
  rtmp.MoveNext
Wend
End Sub

Private Sub Combo1_DropDown()
Dim rtmp As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Combo1_DropDown"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Combo1_DropDown:" + "SELECT * FROM benutzerdaten")
rtmp.Open "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly

Combo1.Clear
While Not rtmp.EOF
  Combo1.AddItem rtmp!id
  rtmp.MoveNext
Wend

If Combo1.ListCount > 0 Then
  Command14.Enabled = True
Else
  Command14.Enabled = False
End If

End Sub

Private Sub Combo2_DropDown()
Dim rtmp As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Combo2_DropDown"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Combo2_DropDown:" + "SELECT * FROM benutzerdaten")
rtmp.Open "SELECT * FROM benutzerdaten", form1.adoc, adOpenDynamic, adLockReadOnly

Combo2.Clear
While Not rtmp.EOF
  Combo2.AddItem rtmp!id
  rtmp.MoveNext
Wend

If Combo2.ListCount > 0 Then
  Command16.Enabled = True
Else
  Command16.Enabled = False
End If

End Sub

Private Sub Command1_Click()
Dim i%, up$, cmd$, didsomething As Integer, fb$, fbu$, fbp$, fbd$, o%, rrr

d2infile = "einstellungen": d2insub = "Command1_Click"
id$ = Text1(0).text
If id$ = "" Then Exit Sub
didsomething = 0

fb$ = form1.getusersetting("fallbackserver", "")
If id$ = form1.getuserid() And fb$ <> "nein" And fb$ <> "none" And fb$ <> "" Then
  fbu$ = form1.getusersetting("fallbackserverusername", "")
  fbp$ = form1.getusersetting("fallbackserverpasswort", "")
  fbd$ = form1.getusersetting("fallbackserverdatenbank", "")
  If InStr(fbp$, "decrypt") = 0 Then
    fbp$ = "decrypt:" + encrypt(fbp$, form1.getinternalkey())
  End If
  If fb$ <> "" Then
    o% = FreeFile
    On Error Resume Next
    Open form1.hppth$ + "\aprepl.ini" For Output As #o%
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      Print #o%, "fallbackserver=" + fb$
      Print #o%, "fallbackserverdatenbank=" + fbd$
      Print #o%, "fallbackserverusername=" + fbu$
      Print #o%, "fallbackserverpasswort=" + fbp$
      Print #o%, "replid_" + form1.computername + "=" + form1.computername
      Print #o%, "replikant_" + form1.computername + "=" + form1.computername
      Print #o%, "replnode_" + form1.computername + "=" + form1.getusersetting("replnode_" + form1.computername, "")
      Close #o%
    End If
  End If
End If

For i% = 1 To nflds
  If Text1(i%).text <> Text2(i%).text Then
    If Len(Text1(i%).text) = 0 Then
      up$ = Label1(i%).Caption + "=NULL"
    Else
      'up$ = Label1(i%).Caption + "= '" + strrepl(Text1(i%).Text, "\", "\\") + "'"
      up$ = Label1(i%).Caption + "= '" + Text1(i%).text + "'"
    End If
    didsomething = 1
    cmd$ = "update benutzerdaten set " + up$ + " where id='" + id$ + "'"
    Call form1.sqlqry(cmd$)
  End If
Next i%

Call showrec(id$)
Call rlist1
einstellungen.BackColor = form1.cleancolor()
If id$ = form1.getuserid() And didsomething Then
  MsgBox transe("Sie haben Ihre eigenen Daten geändert. Das Programm wird neu gestartet.")
  Call form1.unloadall
  DoEvents
  X = Shell("zlauncher" + trm(App.Major) + ".exe", 1)
  End
End If
End Sub

Private Sub Command10_Click()

X = Shell("notepad.exe " + form1.s0dir() + "\" + form1.docs() + "\" + trm(Text1(0).text) + "\signatur.txt", 1)

End Sub

Private Sub Command11_Click()

Hide
Unload einstellungen

End Sub

Private Sub Command12_Click()
d2infile = "einstellungen": d2insub = "Command12_Click"
Call popl_KeyDown(46, 0)
End Sub

Private Sub Command14_Click()
Dim r As ADODB.Recordset, t As ADODB.Recordset, s As ADODB.Recordset, c$, ic%

d2infile = "einstellungen": d2insub = "Command14_Click"
MousePointer = 11
BackColor = form1.dirtycolor()
DoEvents
ic% = 0
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command14_Click:" + "SELECT wert FROM sysvars where owner='blacklist:" + trm(Combo1.text) + "'")
r.Open "SELECT wert FROM sysvars where owner='blacklist:" + trm(Combo1.text) + "'", form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  c$ = "select id from sysvars where owner='blacklist:" + trm(uId$) + "' and wert='" + r!wert + "'"
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  Call form1.dbg2f("einstellungen.Command14_Click:" + c$)
  s.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
  If s.EOF Then
    c$ = "insert into sysvars (id,owner,wert) values('" + _
                  form1.newid("sysvars", "id", 30) + "','blacklist:" + _
                  uId$ + "','" + _
                  trm(r!wert) + "')"
    Call form1.sqlqry(c$)
    ic% = ic% + 1
  End If
  r.MoveNext
Wend
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command14_Click:" + "SELECT wert FROM sysvars where owner='blacklistdom:" + trm(Combo1.text) + "'")
r.Open "SELECT wert FROM sysvars where owner='blacklistdom:" + trm(Combo1.text) + "'", form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  c$ = "select id from sysvars where owner='blacklistdom:" + trm(uId$) + "' and wert='" + trm(r!wert) + "'"
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  Call form1.dbg2f("einstellungen.Command14_Click:" + c$)
  s.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
  If s.EOF Then
    c$ = "insert into sysvars (id,owner,wert) values('" + _
                  form1.newid("sysvars", "id", 30) + "','blacklistdom:" + _
                  uId$ + "','" + _
                  trm(r!wert) + "')"
    Call form1.sqlqry(c$)
    ic% = ic% + 1
  End If
  r.MoveNext
Wend
Call MsgBox(trm(ic%) + " neue Sätze kopiert")
Call rlist4
MousePointer = 0
BackColor = form1.cleancolor()
End Sub

Private Sub Command15_Click()

d2infile = "einstellungen": d2insub = "Command15_Click"
MousePointer = 11: DoEvents
c$ = "delete from sysvars where owner='blacklist:" + trm(uId$) + "'"
form1.sqlqry (c$)
c$ = "delete from sysvars where owner='blacklistdom:" + trm(uId$) + "'"
form1.sqlqry (c$)
Call rlist4
MousePointer = 0
End Sub

Private Sub Command16_Click()
Dim r As ADODB.Recordset, c$, ic%, trg$, w$

d2infile = "einstellungen": d2insub = "Command16_Click"
MousePointer = 11: DoEvents
ic% = List1.ListIndex
If ic% < 0 Then Exit Sub
trg$ = List1.List(ic%)
If List5.ListCount > 0 Then
  Command16.Enabled = False
  Combo2.Enabled = False
End If
BackColor = form1.dirtycolor()
DoEvents
ic% = 0
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command16_Click:" + "SELECT * FROM sysvars where instr(owner,'sysvar_" + trm(Combo2.text) + "')>0")
r.Open "SELECT * FROM sysvars where instr(owner,'sysvar_" + trm(Combo2.text) + "')>0", form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  o$ = Mid$(r!Owner, 8)
  o$ = Mid$(o$, InStr(o$, "_") + 1)
  w$ = "sysvar_" + trg$ + "_" + o$
  c$ = "delete from sysvars where owner='" + w$ + "'"
  form1.sqlqry (c$)
  If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "einstellungen", "Command16_Click")
  c$ = "insert into sysvars (id,owner,wert) values(" + "'" + _
    form1.newid("sysvars", "id", 10) + "','" + _
    w$ + "','" + trm(r!wert) + "');"
  form1.sqlqry (c$)
  r.MoveNext
Wend
Call rlist5
MousePointer = 0
End Sub

Private Sub Command17_Click()
Dim i As Integer

For i = 25 To 28
  Text1(i).text = ""
Next i

End Sub

Private Sub Command18_Click()

d2infile = "einstellungen": d2insub = "Command18_Click"
Call form1.handbuchcall("03-Benutzereinstellungen.htm")

End Sub

Private Sub Command2_Click()
Dim aKey() As Byte

d2infile = "einstellungen": d2insub = "Command2_Click"
If poplistok% = 0 Then
  MsgBox "poplist fehlt, bitte kontaktieren Sie den Support"
  Exit Sub
End If

a$ = trm(svas.text)
If a$ = "" Then Exit Sub
If svas.text <> "PDFServer" Then
  enc$ = trm(pin.text)
Else
  enc$ = "hihallohuhu4716"
End If
If svas.text = "DEFAULT" Then
  enc$ = "hihallohuhu4716"
End If
If enc$ = "" Then
  MsgBox "Bitte wählen Sie ein Passwort um die Passwörter in der Datenbank zu verschüsseln."
  Call pin.SetFocus
  Exit Sub
End If
rc$ = encrypt(Text1(27).text, enc$)
c$ = "delete from poplist where id='" + uId$ + "_" + a$ + "'"
Call form1.sqlqry(c$)
c$ = "insert into poplist (id,server,user,psswd,port) values(" + _
                        "'" + uId$ + "_" + svas.text + "'," + _
                        "'" + Text1(25).text + "'," + _
                        "'" + Text1(26).text + "'," + _
                        "'" + rc$ + "'," + _
                        "'" + Text1(28).text + "'" + _
                        ")"
Call form1.sqlqry(c$)

Call Command9_Click
End Sub


Private Sub Command3_Click()
Dim neuid As String

d2infile = "einstellungen": d2insub = "Command3_Click"
neuid = trm(InputBox(transe("Neue Benutzer-ID"), transe("Neuen Benutzer anlegen")))
If neuid = "" Then Exit Sub

On Error Resume Next
MkDir form1.s0dir() + "\" + form1.docs() + ""
MkDir form1.s0dir() + "\" + form1.docs() + "\" + neuid
On Error GoTo 0

cmd$ = "INSERT INTO benutzerdaten (ID) VALUES('" + neuid + "')"
Call form1.sqlqry(cmd$)
Call showrec(neuid)
Call rlist1

End Sub

Private Sub Command4_Click()
Dim r As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Command4_Click"
id$ = Text1(0).text
If id$ = "" Then Exit Sub

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command4_Click:" + "SELECT * FROM benutzergruppen where userid='" + id$ + "'")
r.Open "SELECT * FROM benutzergruppen where userid='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  While List2.ListCount - 1 > 0
    List2.ListIndex = 0
    Call List2_DblClick
    List2.RemoveItem 0
  Wend
End If

form1.sqlqry ("delete from alarmliste where uid='" + id$ + "'")
form1.sqlqry ("delete from benutzerdaten where id='" + id$ + "'")
Call rlist1
If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub

Private Sub Command5_Click()
d2infile = "einstellungen": d2insub = "Command5_Click"
usrid$ = Text1(0).text
i% = List3.ListIndex
If i% < 0 Then Exit Sub
For j% = 0 To List2.ListCount
  If List2.List(j%) = List3.List(i%) Then
    List2.ListIndex = j%
    Exit Sub
  End If
Next j%
form1.sqlqry ("insert into benutzergruppen (id,groupid,userid) values('" + mkkey(30) + "','" + List3.List(i%) + "','" + usrid$ + "')")
List2.AddItem List3.List(i%)

End Sub

Private Sub Command6_Click()
d2infile = "einstellungen": d2insub = "Command6_Click"
usrid$ = Text1(0).text
i% = List2.ListIndex
If i% < 0 Then Exit Sub
gid$ = List2.List(i%)
form1.sqlqry ("delete from benutzergruppen where groupid='" + gid$ + "' and userid='" + usrid$ + "'")
List2.RemoveItem i%

End Sub

Private Sub Command7_Click()
Dim neuid As String

d2infile = "einstellungen": d2insub = "Command7_Click"
neuid = InputBox(transe("Neue Benutzergruppe"), transe("Neue Benutzergruppe"))
If trm(neuid) <> "" Then
  Call form1.sqlqry("INSERT INTO gruppennamen (gid) VALUES('" + neuid + "')")
  List3.AddItem neuid
End If

End Sub

Private Sub Command8_Click()
Dim r As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Command8_Click"
i% = List3.ListIndex
If i% < 0 Then Exit Sub

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command8_Click:" + "SELECT * FROM benutzergruppen where groupid='" + List3.List(i%) + "'")
r.Open "SELECT * FROM benutzergruppen where groupid='" + List3.List(i%) + "'", form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  MsgBox "Diese Gruppe enthält Benutzer. Löschen nicht möglich."
  Exit Sub
End If

form1.sqlqry ("delete from gruppennamen where gid='" + List3.List(i%) + "'")
List3.RemoveItem i%

End Sub

Private Sub Command9_Click()
Dim r As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Command9_Click"
popl.Clear
Height = bottomline.Y1
If poplistok% = 0 Then
  Exit Sub
End If
If form1.pin.Visible = True And trm(form1.pin.text) <> "" Then pin.text = form1.pin.text
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Command9_Click:" + "SELECT id FROM poplist where instr(id,'" + uId$ + "_')=1")
r.Open "SELECT id FROM poplist where instr(id,'" + uId$ + "_')=1", form1.adoc, adOpenDynamic, adLockReadOnly
While Not r.EOF
  popl.AddItem Mid$(r!id, Len(uId$) + 2)
  r.MoveNext
Wend
Call rlist4
Call rlist5

End Sub

Sub rlist4()
Dim r As ADODB.Recordset, o As Integer

d2infile = "einstellungen": d2insub = "rlist4"
List4.Clear
If shbl.value = 1 Then
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  Call form1.dbg2f("einstellungen.rlist4:" + "SELECT id,wert FROM sysvars where (owner='blacklist:" + uId$ + "' or owner='blacklistdom:" + uId$ + "')")
  r.Open "SELECT id,wert FROM sysvars where (owner='blacklist:" + uId$ + "' or owner='blacklistdom:" + uId$ + "')", form1.adoc, adOpenDynamic, adLockReadOnly
  o = FreeFile
  Open form1.s0dir() + "\" + form1.docs() + "\" + form1.getuserid() + "_spammers.txt" For Output As #o
  While Not r.EOF
    w$ = trm(r!wert)
    If InStr(w$, "@") = 0 Then w$ = "@" + w$
    List4.AddItem w$ + Space$(80) + "(ID:" + trm(r!id)
    Print #o, w$
    r.MoveNext
  Wend
  Close #o
End If
End Sub
Private Sub Form_Load()
Dim r As ADODB.Recordset, cf$, i As Integer

d2infile = "einstellungen": d2insub = "Form_Load"
uId$ = form1.getuserid()
poplistok% = 0
grantlock = False
einstellungen.Caption = transe("Benutzer-Einstellungen AgencyProf")
Command15.ToolTipText = transe("Schwarze Liste löschen")
Command14.Caption = transe("Import")
Command18.ToolTipText = transe("Hilfeseite öffnen")
pin.ToolTipText = transe("Passwort, mit dem die Mailpasswörter verschlüsselt werden")
Command2.Caption = transe("Speichern als")
Command11.ToolTipText = transe("Dieses Formular schliessen")
Command10.Caption = transe("Signatur")
Command9.Caption = transe("weitere Einstellungen")
Label13.Caption = transe("Rechnereinstellungen")
Label12.Caption = transe("Importiere von:")
Label11.Caption = transe("Importiere von:")
Label10.Caption = transe("mehr Userdaten:")
Label9.Caption = transe("schwarze Liste:")
Label8.Caption = transe("PIN:")
Label7.Caption = transe("Benutzer-Daten")
Label6.Caption = transe("ja/nein")
Label5.Caption = transe("Benutzerkennung")
Label4.Caption = transe("Gruppen")
Label3.Caption = transe("ist Mitglied von")
Label2.Caption = transe("msec")
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.Form_Load:" + "SELECT id FROM poplist where instr(id,'" + uId$ + "_')=1")
On Error Resume Next
r.Open "SELECT id FROM poplist where instr(id,'" + uId$ + "_')=1", form1.adoc, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  r.Close
  poplistok% = 1
End If

tp = form1.mylasttop(Me.name)
tl = form1.mylastleft(Me.name)
If tl = 20 And tp = 20 Then
  tl = form1.Left + form1.Width
  tp = form1.Top + form1.Height / 2
End If
Me.Top = tp
Me.Left = tl
Call form1.formpos(Me)
Label13.Caption = form1.computername
Shape2.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Shape3.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Shape1.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents

r.Open "SELECT id FROM sysvars where instr(owner,'sysvar_" + uId$ + "_grant_')=1", form1.adoc, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  If Not r.EOF Then
    grantlock = True
  End If
End If
Show

nflds = 28

Call showrec(uId$)
Call rlist1
'Call rlist4
'Call rlist5
BackColor = form1.cleancolor()
For i = 0 To List1.ListCount - 1
  If List1.List(i) = uId$ Then
    List1.ListIndex = i
    DoEvents
    Me.BackColor = form1.cleancolor
    Exit For
  End If
Next i
BackColor = form1.cleancolor()

End Sub

Private Sub Form_Unload(Cancel As Integer)
d2infile = "einstellungen": d2insub = "Form_Unload"
Call savecheck
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
exuld:
On Error GoTo 0

End Sub



Private Sub Label1_DblClick(Index As Integer)

d2infile = "einstellungen": d2insub = "Label1_DblClick"
If Label1(Index).Caption = "editor" Then
  gwp$ = form1.fixfilename(GetWordPath())
  If gwp$ <> "" Then Text1(Index).text = gwp$
  Exit Sub
End If
If Label1(Index).Caption = "Mailclient" Then
  gwp$ = GetOutlookPath()
  If gwp$ <> "" Then Text1(Index).text = gwp$ + " /c ipm.note"
  Exit Sub
End If

End Sub

Private Sub Label13_DblClick()
d2infile = "einstellungen": d2insub = "Label13_DblClick"
'MsgBox "sorry, das ist leider noch science fiction."
End Sub

Private Sub List1_Click()

d2infile = "einstellungen": d2insub = "List1_Click"
Call savecheck

Call showrec(List1.List(List1.ListIndex))
DoEvents
Call rlist4
Call rlist5
einstellungen.BackColor = form1.cleancolor()
End Sub

Private Sub List2_DblClick()
d2infile = "einstellungen": d2insub = "List2_DblClick"
Call Command6_Click
End Sub

Private Sub List3_DblClick()
d2infile = "einstellungen": d2insub = "List3_DblClick"
Call Command5_Click

End Sub


Private Sub List4_DblClick()
Dim i%, l$, ask, c$

d2infile = "einstellungen": d2insub = "List4_DblClick"
i% = List4.ListIndex
If i% < 0 Then Exit Sub
l$ = List4.List(i%)
l$ = trm(Left(l$, InStr(l$, "(ID:") - 1))
l$ = domainofemail(l$)
If l$ = "" Then Exit Sub
ask = MsgBox("Sollen alle Adressen der Domain " + l$ + " zur schwarzen Liste?", vbYesNo + vbCritical + vbDefaultButton2, "Domain sperren?")
If ask = vbNo Then Exit Sub
c$ = "delete from sysvars where ((owner='blacklist:" + form1.getuserid() + "') and (instr(wert,'" + l$ + "')>0));"
Call form1.sqlqry(c$)
c$ = "insert into sysvars (id,owner,wert) values('" + form1.newid("sysvars", "id", 36) + _
         "','blacklistdom:" + form1.getuserid() + _
         "','" + l$ + "')"
Call form1.sqlqry(c$)
Call rlist4
i% = i% + 1
If i% >= List4.ListCount Then i% = List4.ListCount - 1
List4.ListIndex = i%
j% = i% + 3
If j% >= List4.ListCount Then j% = List4.ListCount - 1
List4.ListIndex = j%
'DoEvents
List4.ListIndex = i%
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
d2infile = "einstellungen": d2insub = "List4_KeyDown"
If KeyCode = 46 Or KeyCode = 8 Then
  i% = List4.ListIndex
  If i% < 0 Then Exit Sub
  id$ = List4.List(i%)
  id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
  c$ = "delete FROM sysvars where id='" + id$ + "'"
  Call form1.sqlqry(c$)
  List4.RemoveItem i%
End If

End Sub


Private Sub List5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%
Dim n$, wert$, p%, c$

d2infile = "einstellungen": d2insub = "List5_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  i% = List5.ListIndex
  If i% >= 0 Then
    wert$ = List5.List(i%)
    p% = InStr(wert$, "=")
    If p% > 0 Then
      wert$ = trm(Left$(wert$, p% - 1))
      wert$ = "sysvar_" + uId$ + "_" + wert$
      c$ = "delete from sysvars where owner='" + wert$ + "'"
      Call form1.sqlqry(c$)
      If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "einstellungen", "List5_Keydown")
      Call rlist5
      If i% > List5.ListCount - 1 Then i% = List5.ListCount - 1
      List5.ListIndex = i%
    End If
  End If
End If

End Sub

Private Sub popl_Click()
Dim r As ADODB.Recordset

d2infile = "einstellungen": d2insub = "popl_Click"
enc$ = trm(pin.text)
If enc$ = "" Then
  MsgBox "Bitte geben Sie das Passwort ein mit dem die Passwörter in der Datenbank entschüsselt werden."
  Call pin.SetFocus
  Exit Sub
End If

id$ = uId$ + "_" + popl.List(popl.ListIndex)
svas.text = popl.List(popl.ListIndex)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.popl_Click:" + "SELECT * FROM poplist where id='" + id$ + "'")
r.Open "SELECT * FROM poplist where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  rc$ = decrypt(r!psswd, enc$)
  nserv.Caption = "Server: " + trm(r!server) + ", Benutzer: " + trm(r!user)
End If

End Sub

Private Sub popl_DblClick()
Dim r As ADODB.Recordset
Dim aKey() As Byte

d2infile = "einstellungen": d2insub = "popl_DblClick"
enc$ = trm(pin.text)
If enc$ = "" Then
  MsgBox "Bitte geben Sie das Passwort ein mit dem die Passwörter in der Datenbank entschüsselt werden."
  Call pin.SetFocus
  Exit Sub
End If

id$ = uId$ + "_" + popl.List(popl.ListIndex)

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.popl_DblClick:" + "SELECT * FROM poplist where id='" + id$ + "'")
r.Open "SELECT * FROM poplist where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then
  rc$ = decrypt(r!psswd, enc$)
  Text1(25).text = r!server
  Text1(26).text = r!user
  Text1(27).text = rc$
  Text1(28).text = r!Port
End If


End Sub

Private Sub popl_KeyDown(KeyCode As Integer, Shift As Integer)

d2infile = "einstellungen": d2insub = "popl_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  id$ = uId$ + "_" + popl.List(popl.ListIndex)
  c$ = "delete FROM poplist where id='" + id$ + "'"
  Call form1.sqlqry(c$)
  Call Command9_Click
End If

End Sub

Private Sub Text1_Change(Index As Integer)

d2infile = "einstellungen": d2insub = "Text1_Change"
Command1.Enabled = True
einstellungen.BackColor = form1.dirtycolor()

End Sub
Public Sub showrec(nuId$)
Dim rtmp As ADODB.Recordset

d2infile = "einstellungen": d2insub = "showrec"
uId$ = nuId$
i% = 0

For i% = 0 To nflds
  Label1(i%).Caption = transe(form1.sqla.TableDefs("benutzerdaten").Fields(i%).name)
  Text1(i%).text = ""
  Text2(i%).text = ""
Next i%
Command1.Enabled = False

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.showrec:" + "SELECT * FROM benutzerdaten where id ='" + uId$ + "'")
rtmp.Open "SELECT * FROM benutzerdaten where id ='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then
  For i% = 0 To nflds
    If Not IsNull(rtmp.Fields(i%)) Then
      Text1(i%).text = rtmp.Fields(i%)
      Text2(i%).text = rtmp.Fields(i%)
    End If
    If Label1(i%).Caption = "mysql" Then
      Text1(i%).ToolTipText = "Voller Name von mysql.exe. z.B. c:\mysql\bin\mysql.exe"
      Label1(i%).ToolTipText = Text1(i%).ToolTipText
    End If
    If Label1(i%).Caption = "mysqlhost" Then
      Text1(i%).ToolTipText = "IP-Nummer oder Name des Servers, auf dem die Datenbank läuft. Der eigene Rechner ist localhost oder 127.0.0.1"
      Label1(i%).ToolTipText = Text1(i%).ToolTipText
    End If
  Next i%
End If
Call rlist23

End Sub

Sub savecheck()
d2infile = "einstellungen": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then Call Command1_Click
End If
BackColor = form1.cleancolor()
End Sub
Sub rlist5()
Dim rtmp As ADODB.Recordset
Dim o$, myuID As String

d2infile = "einstellungen": d2insub = "rlist5"
'o$="SELECT * FROM sysvars where owner like 'sysvar_" + uId$ + "_*'"
o$ = "SELECT * FROM sysvars where instr(owner,'sysvar_" + uId$ + "_')>0"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
Call form1.dbg2f("einstellungen.rlist5:" + o$)
rtmp.Open o$, form1.adoc, adOpenDynamic, adLockReadOnly

List5.Clear
myuID = form1.getuserid()
While Not rtmp.EOF
  o$ = rtmp!Owner
  o$ = Mid$(o$, InStr(o$, uId$) + Len(uId$) + 1)
  If LCase(o$) <> "rechnr" Or myuID = "www" Or myuID = "administrator" Then
    If InStr(o$, "grant_") = 0 Or (InStr(o$, "grant_") = 1 And Not grantlock) Then List5.AddItem o$ + "=" + trm(rtmp!wert)
  End If
  rtmp.MoveNext
Wend
Label12.Caption = transe("Importiere von")
Command16.Enabled = True
Combo2.Enabled = True
If List5.ListCount > 0 Then
  Command16.Enabled = False
  Combo2.Enabled = False
  Label12.Caption = transe("Kein Import möglich")
End If

End Sub
Private Sub Command13_Click()
Dim n$, wert$, p%, c$, r As ADODB.Recordset

d2infile = "einstellungen": d2insub = "Command13_Click"
wert$ = "wert="
wert$ = InputBox(transe("Neue Benutzereinstellung:"), transe("Neue Einstellung"), wert$)
If InStr(LCase(wert$), "grant") = 1 And grantlock Then Exit Sub
p% = InStr(wert$, "=")
If p% > 0 Then
  n$ = Mid$(wert$, p% + 1)
  wert$ = trm(Left$(wert$, p% - 1))
  If Left(n$, 8) = "encrypt:" Then
    n$ = "decrypt:" + encrypt(Mid$(n$, 9), form1.getinternalkey())
  End If
  If (LCase(uId$) <> "system" And LCase(wert$) <> "rechnr") Or (LCase(uId$) = "system") Then
    wert$ = "sysvar_" + uId$ + "_" + wert$
    c$ = "select * from sysvars where lcase(owner)='" + LCase(wert) + "';"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    Call form1.dbg2f("einstellungen.newsysvar:" + c$)
    r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
    If r.EOF Then
      c$ = "insert into sysvars (id,owner,wert) values('" + _
                 form1.newid("sysvars", "id", 30) + "','" + _
                 wert$ + "','" + _
                 n$ + "')"
      Call form1.sqlqry(c$)
      Call form1.rereadsomesysvars
      Call rlist5
    Else
      MsgBox transe("Einstellung existiert bereits.")
    End If
  Else
    MsgBox transe("Dieser Einstellungsname ist reserviert.")
  End If
End If

End Sub


Private Sub List5_dblClick()
Dim i%, mi%
Dim n$, wert$, p%, c$

d2infile = "einstellungen": d2insub = "List5_dblClick"
i% = List5.ListIndex
If i% >= 0 Then
mi% = i%
wert$ = List5.List(i%)
p% = InStr(wert$, "=")
If p% > 0 Then
  n$ = trm(Mid$(wert$, p% + 1))
  wert$ = trm(Left$(wert$, p% - 1))
  n$ = InputBox(transe("Neue Benutzereinstellung:") + vbCrLf + wert$, "Neue Einstellung", n$)
  If n$ = "" Then
    ask = MsgBox("Der Eintrag soll gelöscht werden?", vbYesNo + vbCritical + vbDefaultButton2, "Wert löschen?")
    If ask = vbNo Then Exit Sub
  Else
    If Left(n$, 8) = "encrypt:" Then
      n$ = "decrypt:" + encrypt(Mid$(n$, 9), form1.getinternalkey())
    End If
  End If
  wert$ = "sysvar_" + uId$ + "_" + wert$
  c$ = "delete from sysvars where owner='" + wert$ + "'"
  Call form1.sqlqry(c$)
  If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "einstellungen", "dblclick")
  c$ = "insert into sysvars (id,owner,wert) values('" + _
        form1.newid("sysvars", "id", 30) + "','" + _
        wert$ + "','" + _
        n$ + "')"
  Call form1.sqlqry(c$)
  Call form1.rereadsomesysvars
  Call rlist5
  If mi% >= 0 And mi% < List5.ListCount Then List5.ListIndex = mi%
End If
End If

End Sub


