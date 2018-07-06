VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form auftritt 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Termine"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox fromtpwernoch 
      Height          =   1035
      Left            =   7200
      TabIndex        =   182
      Top             =   8880
      Width           =   2415
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2280
      Picture         =   "auftritt.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   181
      ToolTipText     =   "Adressen im Umkreis um eine Postleitzahl suchen"
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "auftritt.frx":0672
      Style           =   1  'Grafisch
      TabIndex        =   180
      ToolTipText     =   "Formular schiessen"
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command22 
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
      Height          =   195
      Left            =   3360
      TabIndex        =   179
      ToolTipText     =   "show also in project ..."
      Top             =   480
      Width           =   375
   End
   Begin VB.ListBox tmplst 
      Height          =   1425
      Index           =   1
      Left            =   11040
      TabIndex        =   178
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox tmplst 
      Height          =   1425
      Index           =   0
      Left            =   10800
      TabIndex        =   177
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton btnTopic 
      Caption         =   "Topic"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   176
      Top             =   8040
      Width           =   615
   End
   Begin VB.CommandButton chkedt 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   8880
      Picture         =   "auftritt.frx":08C2
      Style           =   1  'Grafisch
      TabIndex        =   175
      ToolTipText     =   "gewählte Nachricht bearbeiten"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton chksve 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   8880
      Picture         =   "auftritt.frx":1A14
      Style           =   1  'Grafisch
      TabIndex        =   174
      ToolTipText     =   "Auftritt speichern"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox uselct 
      Height          =   1425
      Left            =   9000
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   172
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton chkown 
      Caption         =   "Besitzer"
      Height          =   495
      Left            =   8880
      TabIndex        =   173
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton chkeclse 
      Caption         =   "(x)"
      Height          =   495
      Left            =   8880
      TabIndex        =   171
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton chkedone 
      Caption         =   "Neuer Checkpoint"
      Height          =   495
      Left            =   8880
      TabIndex        =   170
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton chkudall 
      Caption         =   "Alle wiederherstellen"
      Height          =   495
      Left            =   8880
      TabIndex        =   169
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton chkdlall 
      Caption         =   "Alle löschen"
      Height          =   495
      Left            =   8880
      TabIndex        =   168
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton chkokall 
      Caption         =   "Alle ok"
      Height          =   495
      Left            =   8880
      TabIndex        =   164
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView listMessages 
      Height          =   1455
      Left            =   9000
      TabIndex        =   167
      Top             =   2160
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2566
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "auftritt.frx":1DBB
      Style           =   1  'Grafisch
      TabIndex        =   166
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "auftritt.frx":293D
      Style           =   1  'Grafisch
      TabIndex        =   165
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox pstt 
      Height          =   255
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   99
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   600
      Picture         =   "auftritt.frx":34BF
      ScaleHeight     =   300
      ScaleWidth      =   285
      TabIndex        =   163
      Top             =   8400
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   960
      Picture         =   "auftritt.frx":39B1
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   162
      Top             =   8400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox prio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9000
      TabIndex        =   161
      ToolTipText     =   "Priorität"
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox tmpsort 
      Height          =   645
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   160
      Top             =   9120
      Width           =   975
   End
   Begin MSComctlLib.ListView gd1 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   158
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton vwopts 
      Height          =   375
      Left            =   6240
      Picture         =   "auftritt.frx":3EA3
      Style           =   1  'Grafisch
      TabIndex        =   157
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      Picture         =   "auftritt.frx":501D
      Style           =   1  'Grafisch
      TabIndex        =   156
      ToolTipText     =   "Kalender öffnen"
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Left            =   8880
      Top             =   1320
   End
   Begin VB.PictureBox calcol 
      Height          =   255
      Left            =   8520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   155
      ToolTipText     =   "Farbe im Kalender (Doppelklick zum ändern)"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1080
      Picture         =   "auftritt.frx":51A7
      Style           =   1  'Grafisch
      TabIndex        =   154
      ToolTipText     =   "Alle Tabellen und Kalkulationen neu berechnen"
      Top             =   8040
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8880
      Picture         =   "auftritt.frx":5819
      Style           =   1  'Grafisch
      TabIndex        =   153
      ToolTipText     =   "Neue Vorlage für diesen Termintyp erstellen"
      Top             =   360
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "auftritt.frx":5BAB
      Left            =   7680
      List            =   "auftritt.frx":5BAD
      TabIndex        =   152
      ToolTipText     =   "Dateinamen erstellen aus:"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6720
      Picture         =   "auftritt.frx":5BAF
      Style           =   1  'Grafisch
      TabIndex        =   149
      ToolTipText     =   "Auftritt als Mail versenden"
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton opendir 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   720
      Picture         =   "auftritt.frx":5D39
      Style           =   1  'Grafisch
      TabIndex        =   148
      ToolTipText     =   "Terminverzeichnis im Explorer öffnen"
      Top             =   8040
      Width           =   375
   End
   Begin VB.Timer Timer_dtst 
      Left            =   3960
      Top             =   8160
   End
   Begin VB.CommandButton dtst 
      Caption         =   "Daten- test"
      Height          =   495
      Left            =   5400
      TabIndex        =   147
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   33
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   145
      Text            =   "auftritt.frx":6363
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   33
      Left            =   8640
      TabIndex        =   144
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   32
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   142
      Text            =   "auftritt.frx":6369
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   32
      Left            =   8640
      TabIndex        =   141
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   31
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   139
      Text            =   "auftritt.frx":636F
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   31
      Left            =   8640
      TabIndex        =   138
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   30
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   136
      Text            =   "auftritt.frx":6375
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   30
      Left            =   8640
      TabIndex        =   135
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   29
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   133
      Text            =   "auftritt.frx":637B
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   29
      Left            =   8640
      TabIndex        =   132
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   28
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   130
      Text            =   "auftritt.frx":6381
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   28
      Left            =   8640
      TabIndex        =   129
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   27
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   127
      Text            =   "auftritt.frx":6387
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   27
      Left            =   8640
      TabIndex        =   126
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   26
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   124
      Text            =   "auftritt.frx":638D
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   26
      Left            =   8640
      TabIndex        =   123
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   25
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   121
      Text            =   "auftritt.frx":6393
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   25
      Left            =   8640
      TabIndex        =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   24
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   118
      Text            =   "auftritt.frx":6399
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   24
      Left            =   8640
      TabIndex        =   117
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   23
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   115
      Text            =   "auftritt.frx":639F
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   23
      Left            =   8640
      TabIndex        =   114
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   22
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   112
      Text            =   "auftritt.frx":63A5
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   22
      Left            =   8640
      TabIndex        =   111
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   109
      Text            =   "auftritt.frx":63AB
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   21
      Left            =   8640
      TabIndex        =   108
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   20
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   106
      Text            =   "auftritt.frx":63B1
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   20
      Left            =   8640
      TabIndex        =   105
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
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
      Height          =   375
      Left            =   480
      TabIndex        =   104
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox kalimmer 
      Height          =   255
      Left            =   6000
      TabIndex        =   102
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox kalres 
      Height          =   255
      Left            =   6000
      TabIndex        =   100
      Top             =   8280
      Width           =   255
   End
   Begin VB.ComboBox astatcmb 
      Height          =   315
      Left            =   6720
      TabIndex        =   98
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox mwst 
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "€"
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
      Left            =   4080
      TabIndex        =   96
      Top             =   8040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "->"
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
      Left            =   3360
      TabIndex        =   95
      ToolTipText     =   "In ein anderes Projekt verschieben"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "auftritt.frx":63B7
      Style           =   1  'Grafisch
      TabIndex        =   94
      ToolTipText     =   "Auftritt speichern"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton wvl 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   7440
      Picture         =   "auftritt.frx":675E
      Style           =   1  'Grafisch
      TabIndex        =   93
      ToolTipText     =   "Wiedervorlage"
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5280
      Picture         =   "auftritt.frx":6DCA
      Style           =   1  'Grafisch
      TabIndex        =   92
      ToolTipText     =   "Diesen Auftritt löschen"
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "wiederholen"
      Height          =   255
      Left            =   6840
      TabIndex        =   91
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   ">>"
      Height          =   495
      Left            =   8640
      TabIndex        =   90
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<<"
      Height          =   495
      Left            =   1440
      TabIndex        =   89
      Top             =   8040
      Width           =   375
   End
   Begin VB.ListBox chgs 
      Height          =   450
      Left            =   3120
      TabIndex        =   88
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Alarme"
      Height          =   255
      Left            =   4680
      TabIndex        =   87
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ausfül&len"
      Height          =   255
      Left            =   5640
      TabIndex        =   86
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "kopieren"
      Height          =   255
      Left            =   8160
      TabIndex        =   85
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   84
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Drucken:"
      Height          =   255
      Left            =   6720
      TabIndex        =   83
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&>"
      Height          =   495
      Left            =   8040
      TabIndex        =   81
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&<"
      Height          =   495
      Left            =   1920
      TabIndex        =   80
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   19
      Left            =   8640
      TabIndex        =   77
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   18
      Left            =   8640
      TabIndex        =   76
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   17
      Left            =   8640
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   16
      Left            =   4200
      TabIndex        =   74
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   15
      Left            =   4200
      TabIndex        =   73
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   14
      Left            =   4200
      TabIndex        =   72
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   13
      Left            =   4200
      TabIndex        =   71
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   12
      Left            =   4200
      TabIndex        =   70
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   11
      Left            =   4200
      TabIndex        =   69
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   68
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   67
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   66
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   65
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   64
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   63
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   61
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   59
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   58
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   19
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   57
      Text            =   "auftritt.frx":80A0
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   18
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   55
      Text            =   "auftritt.frx":80A6
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   17
      Left            =   6000
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   53
      Text            =   "auftritt.frx":80AC
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   51
      Text            =   "auftritt.frx":80B2
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   15
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   49
      Text            =   "auftritt.frx":80B8
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   47
      Text            =   "auftritt.frx":80BE
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   45
      Text            =   "auftritt.frx":80C4
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   43
      Text            =   "auftritt.frx":80CA
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   41
      Text            =   "auftritt.frx":80D0
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   39
      Text            =   "auftritt.frx":80D6
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   37
      Text            =   "auftritt.frx":80DC
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   35
      Text            =   "auftritt.frx":80E2
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   33
      Text            =   "auftritt.frx":80E8
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   31
      Text            =   "auftritt.frx":80EE
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   29
      Text            =   "auftritt.frx":80F4
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   27
      Text            =   "auftritt.frx":80FA
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   25
      Text            =   "auftritt.frx":8100
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   23
      Text            =   "auftritt.frx":8106
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   21
      Text            =   "auftritt.frx":810C
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1560
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatisch
      ScrollBars      =   2  'Vertikal
      TabIndex        =   7
      Text            =   "auftritt.frx":8112
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1560
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   5055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "auftritt.frx":8118
      Left            =   3720
      List            =   "auftritt.frx":811A
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1560
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1065
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4560
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   5280
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3360
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1560
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   345
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox findfeld 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   151
      Text            =   "Gehe zu"
      ToolTipText     =   "Feld nach Namen suchen"
      Top             =   1080
      Width           =   390
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      Picture         =   "auftritt.frx":811C
      Style           =   1  'Grafisch
      TabIndex        =   150
      ToolTipText     =   "Kalender öffnen"
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label basemerk 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   2760
      TabIndex        =   159
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   33
      Left            =   4560
      TabIndex        =   146
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   32
      Left            =   4560
      TabIndex        =   143
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   31
      Left            =   4560
      TabIndex        =   140
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   30
      Left            =   4560
      TabIndex        =   137
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   29
      Left            =   4560
      TabIndex        =   134
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   28
      Left            =   4560
      TabIndex        =   131
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   27
      Left            =   4560
      TabIndex        =   128
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   26
      Left            =   4560
      TabIndex        =   125
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   25
      Left            =   4560
      TabIndex        =   122
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   24
      Left            =   4560
      TabIndex        =   119
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   23
      Left            =   4560
      TabIndex        =   116
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   22
      Left            =   4560
      TabIndex        =   113
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   21
      Left            =   4560
      TabIndex        =   110
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   20
      Left            =   4560
      TabIndex        =   107
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kalender öffnen"
      Height          =   255
      Left            =   6240
      TabIndex        =   103
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Kalender nur für Beteiligte"
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
      Left            =   6240
      TabIndex        =   101
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "% MwSt"
      Height          =   255
      Left            =   4080
      TabIndex        =   97
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   2040
      TabIndex        =   82
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   2280
      TabIndex        =   79
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   78
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   19
      Left            =   4560
      TabIndex        =   56
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   18
      Left            =   4560
      TabIndex        =   54
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   17
      Left            =   4560
      TabIndex        =   52
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   50
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   48
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   46
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   42
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   40
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   38
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   36
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   34
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   15
      ToolTipText     =   "Doppelklick zum Ändern des Termintyps"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6495
      Left            =   0
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1440
      Width           =   9255
   End
   Begin VB.Menu dat_fkts 
      Caption         =   "&Datei"
      Visible         =   0   'False
      Begin VB.Menu dat_open 
         Caption         =   "... ö&ffnen"
         Shortcut        =   ^O
      End
      Begin VB.Menu ruler 
         Caption         =   "----------------"
      End
      Begin VB.Menu dat_close 
         Caption         =   "&schließen"
      End
   End
   Begin VB.Menu prn_fkts 
      Caption         =   "D&rucken"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "auftritt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nbase%, setast
Dim nflds As Integer, prv$, prvd$, prvt3$
Dim clickgetsfromtable(34)
Dim clickgetsfromfield(34)
Dim neuprogid$, v0base%, prvtt$, adrfldlist As String
Dim gotnetto$, kalkdirty As Boolean
Dim kalrestrict$(99), krcount%, angezeigtefelder%, formtt(99) As String, formttp As Integer
Dim gdmx As Integer, fl_showrec%
Dim delmode As Boolean, fpp%, chgsx(99) As String, recalclist As String, recalcplease As Boolean

Private Function chgsread(n As Integer)

If chgs.List(n) = "substituted" Then
  chgsread = chgsx(n)
  Exit Function
End If
chgsread = chgs.List(n)
End Function

Private Function chgswrite(s As String)
Dim i As Integer, t As String
i = chgs.ListCount
If Len(s) < 1000 Then
  chgs.AddItem s
Else
  chgs.AddItem "substituted"
  chgsx(i) = s
End If
End Function

Public Sub calccallback(l$)
gotnetto$ = l$
End Sub

Public Sub callback(p$)
neuprogid$ = p$
End Sub

Private Sub astatcmb_Click()
Dim ast As Integer

'd2infile = "auftritt": d2insub = "astatcmb_Click"
id$ = Text1(0).text
ast = astatcmb.ListIndex
astatcmb.Visible = False
pstt.Visible = True
pstt.BackColor = form1.get_eventstatuscolor(ast)
pstt.Cls
pstt.Print form1.get_eventstatusname(ast)

c$ = "update auftritt set astatus=" & trm(ast) & " where id='" & id$ & "'"
Call form1.sqlqry(c$)
If form1.cloud Then Call makemedirty
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub btnTopic_Click()

If form1.isfieldmissing("opt_topics", "id") Then Exit Sub
Load dochist2
DoEvents
dochist2.topics.Clear
dochist2.topics.AddItem Text1(1).text
dochist2.topics.Selected(0) = True
'Call dochist2.topics_Click

End Sub

Private Sub calcol_DblClick()
'd2infile = "auftritt": d2insub = "calcol_DblClick"
If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
  Load colorsel
  colorsel.SetFocus
  colorsel.updc (calcol.BackColor)
  Timer2.Enabled = True
  Timer2.Interval = 1000
Else
  MsgBox ("Feld ""optkalcolor"" in Tabelle ""auftritt"" fehlt. Funktion ist abgeschaltet.")
End If

End Sub


Private Sub chkdlall_Click()
Dim i%

For i% = 1 To listMessages.ListItems.Count
  listMessages.ListItems(i%).Selected = True
Next i%
Call listMessages_KeyDown(46, 0)

End Sub

Private Sub chkeclse_Click()
Call Command21_Click(1)
End Sub

Private Sub chkedone_Click()
Dim lvitem, hdm As Boolean, strMessageHeader As String
Dim r As ADODB.Recordset, c$, n%
Dim id$, pos%, cnf$, cf$

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "listMessages_DblClick"
On Error Resume Next

rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
aid$ = Text1(0).text
If aid$ = "" Then Exit Sub
p% = 0: n% = 0
cf$ = "dd.mm.yyyy:Text"
  wert$ = InputBox("edit: " + cf$ + vbCrLf + transe("Syntax: Datum der Erinnerung:Text der Frage") + vbCrLf + vbCrLf + "'" + trm(Date) + transe(":hallo?' --> Reminder mit der Frage 'hallo?' heute."), "Reminder", cf$)
  If wert$ <> cf$ Then
    cnf$ = cut_d1(wert$, ":")
    cf$ = cut_d2bis(wert$, ":")
    If cf$ <> "" Then
      cnf$ = datum2sql(cnf$)
      cid$ = form1.newidbase("opt_checks", "id", 5, "aid_", "_" + aid$)
      c$ = "insert into opt_checks (id,auftrittsid,dtg,checkpoint) values("
      c$ = c$ + "'" + cid$ + "',"
      c$ = c$ + "'" + aid$ + "',"
      c$ = c$ + "'" + cnf$ + "',"
      c$ = c$ + "'" + cf$ + "')"
      Call form1.sqlqry(c$)
    End If
    Call Command21_Click(999)
  Else
    MsgBox ("Syntx-Error")
  End If

End Sub

Private Sub chkedt_Click()
Dim i%, id$

For i% = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i%).Selected = True Then
    id$ = listMessages.ListItems(i%).SubItems(4)
    If id$ <> "" Then
      Load remedit
      remedit.remid = id$
      Exit Sub
    End If
  End If
Next i%

End Sub

Private Sub chkokall_Click()
Dim frm$, p%, rrr, i%, o%, l$, sbf$, sbj$, trg$, msgid$
Dim lvitem, hdm As Boolean, strMessageHeader As String
Dim r As ADODB.Recordset, c$, n%
Dim id$, pos%, cnf$, cf$

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "chkokall_Click"
On Error Resume Next
frm$ = listMessages.SelectedItem
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

p% = 0: n% = 0
For i% = 1 To listMessages.ListItems.Count

cnf$ = listMessages.ListItems(i%).SubItems(2)
If cnf$ = "" Then
  id$ = listMessages.ListItems(i%).SubItems(4)
  wert$ = "ok, " + trm(Date) + " " + trm(Time) + " " + form1.getuserid()
  listMessages.ListItems(i%).SubItems(2) = wert$
  DoEvents
  cf$ = "update opt_checks set confirmed='" + wert$ + "' where id='" + id$ + "'"
  Call form1.sqlqry(cf$)
End If

Next i%
listMessages.Visible = False
Call Command21_Click(1)
id$ = Text1(0).text
If id$ <> "" Then Call achktst(id$)

End Sub

Private Sub chkown_Click()
Dim r As ADODB.Recordset, c$, rrr, d$

uselct.Top = chkown.Top + chkown.Height - uselct.Height
uselct.Left = chkown.Left
uselct.Clear
c$ = "select ID from benutzerdaten order by id"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
While Not r.EOF
  uselct.AddItem trm(r!id)
  r.MoveNext
Wend
d$ = adrfldlist: c$ = "x"
While d$ <> ""
  c$ = cut_d1(d$, "|"): d$ = cut_d2bis(d$, "|")
  If c$ <> "" Then
    uselct.AddItem "{" + c$
  End If
Wend
uselct.Visible = True
chksve.Visible = True
chkedt.Visible = True
End Sub

Private Sub chksve_Click()
Dim i%, r$, id$, c$

rc$ = ""
uselct.Visible = False
chksve.Visible = False

For i% = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i%).Selected = True Then Exit For
Next i%
If i% <= listMessages.ListItems.Count Then
  id$ = listMessages.ListItems(i%).SubItems(4)

  For i% = 1 To uselct.ListCount - 1
    If uselct.Selected(i%) = True Then
      r$ = uselct.List(i%)
      If Left$(r$, 1) = "{" Then
        c$ = "select FeldDaten as wert from auftritthigru where auftrittsid='" + Text1(0).text + "' and FeldName='" + Mid$(r$, 2) + "'"
        r$ = form1.get1erg(c$)
        If r$ <> "" Then r$ = form1.APUsernameByAddressID(r$)
      End If
      If r$ <> "" Then
        If rc$ <> "" Then rc$ = rc$ + "|"
        rc$ = rc$ + r$
      End If
    End If
  Next i%
  If rc$ <> "" Then c$ = "update opt_checks set ownr='" + rc$ + "' where id='" + id$ + "'"
  Call form1.sqlqry(c$)
End If
Call shw_reminders
End Sub

Private Sub chkudall_Click()
Dim cf$, id$

id$ = Text1(0).text
If id$ = "" Then Exit Sub
cf$ = "update opt_checks set confirmed='' where auftrittsid='" + id$ + "' and confirmed like 'ok, delete%'"
Call form1.sqlqry(cf$)
Call Command21_Click(999)

End Sub

Private Sub combo1_Change()
Dim id$

'd2infile = "auftritt": d2insub = "combo1_Change"
id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  Exit Sub
End If
p$ = transo(Text1(5).text)
If Combo1.ListIndex >= 0 Then Text1(5).text = Combo1.List(Combo1.ListIndex)
If Len(Text1(5).text) > 0 Then
  If p$ = "Neuer Auftritt" Then
    chgs.AddItem "update auftritt set Auftrittstyp='" + p$ + "' where id='" + id$ + "'"
    Call showrec(id$, 0)
    Call makemedirty
  End If
End If

End Sub

Public Sub Combo1_Click()
Dim p$, ntyp$, tabf$, o%, unl As Boolean
Dim r As ADODB.Recordset, c$, rrr, fn As String, ffn$
Dim s As ADODB.Recordset, d$, fldn$, l$, tr, M$, w$, f$
Dim vonlist(199) As String, nachlist(199) As String, vptr%, nptr%, i%, j%

'd2infile = "auftritt": d2insub = "Combo1_Click"
id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  Exit Sub
End If
unl = False
p$ = transo(Text1(5).text)
Text1(5).text = Combo1.List(Combo1.ListIndex)
If Len(Text1(5).text) > 0 Then
  If p$ = "Neuer Auftritt" Then
    chgs.AddItem "update auftritt set Auftrittstyp='" & transo(Text1(5).text) & "' where id='" + id$ + "'"
    Call form1.sqlqry("update auftritthigru set auftrittstyp='" + transo(Text1(5).text) + "' where auftrittsid='" + id$ + "' and feldname='zzzsysez' and auftrittstyp='Neuer Auftritt';")
    Call showrec(id$, 1)
    Call makemedirty
    l$ = form1.getusersetting("eventimport_" + LCase(transo(Text1(5).text)), "")
    If l$ <> "" Then
      fn = Dir(l$ + "\*.csv")
      If fn <> "" Then
        o% = FreeFile
        M$ = form1.vorlagenverzeichnis() + "\apcsvimport_" & LCase(transo(Text1(5).text)) & ".imp"
        If nexist(M$) Then
          MsgBox "vorlage fehlt: " + M$
        Else
          Open M$ For Input As #o%
          While Not EOF(o%)
            Line Input #o%, M$
            M$ = trm(M$)
            If M$ <> "" And Left$(M$, 1) <> ";" Then
              tmplst(0).AddItem trm(cut_d1(M$, ";")): M$ = cut_d2bis(M$, ";")
              tmplst(1).AddItem trm(cut_d1(M$, ";"))
            End If
          Wend
          Close #o%
          o% = FreeFile
          ffn$ = l$ + "\" + fn
          Open ffn$ For Input As #o%
          While Not EOF(o%)
            Line Input #o%, l$
            For i% = 0 To tmplst(0).ListCount - 1
              If tmplst(0).List(i%) = cut_d1(l$, ";") Then
                w$ = trm(strrepl(cut_d1(cut_d2bis(l$, ";"), ";"), ";", " "))
                f$ = tmplst(1).List(i%)
                'MsgBox w$ + "->" + f$
                For j% = 0 To 10
                  If Label2(j%).Caption = f$ Then
                    Call Text2_GotFocus(j%): Text2(j%) = w$: Call Text2_LostFocus(j%)
                    Exit For
                  End If
                Next j%
                Exit For
              End If
            Next i%
          Wend
          Close #o%
          On Error Resume Next
          Kill ffn$
          On Error GoTo 0
        End If
      End If
    End If
  Else
    If p$ <> transo(Combo1.List(Combo1.ListIndex)) And Combo1.List(Combo1.ListIndex) <> "" Then
      i% = Combo1.ListIndex
      If i% >= 0 Then
        ntyp$ = transo(Combo1.List(i%))
        tabf$ = form1.s0dir() + "\" + form1.docs() + "\" + strrepl(p$, " ", "_") + "--" + strrepl(ntyp$, " ", "_") + ".txt"
        If nexist(tabf$) Then
          o% = FreeFile
          Open tabf$ For Output As #o%
          c$ = "select FeldName from auftrittsfelder where typ='" + p$ + "' order by position"
          Set r = New ADODB.Recordset
          r.CursorLocation = adUseServer
          rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          d$ = "select FeldName from auftrittsfelder where typ='" + ntyp$ + "' order by position"
          Set s = New ADODB.Recordset
          s.CursorLocation = adUseServer
          rrr = form1.adoopen(s, d$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          vptr% = 0
          While Not r.EOF
            vonlist(vptr%) = trm(r!feldname)
            If InStr(vonlist(vptr%), ".") > 0 Then
              vonlist(vptr%) = cut_d2bis(vonlist(vptr%), ".")
              vonlist(vptr%) = cut_d1(vonlist(vptr%), ".")
            End If
            vptr% = vptr% + 1
            r.MoveNext
          Wend
          nptr% = 0
          While Not s.EOF
            nachlist(nptr%) = trm(s!feldname)
            If InStr(nachlist(nptr%), ".") > 0 Then
              nachlist(nptr%) = cut_d2bis(nachlist(nptr%), ".")
              nachlist(nptr%) = cut_d1(nachlist(nptr%), ".")
            End If
            nptr% = nptr% + 1
            s.MoveNext
          Wend
          For i% = 0 To vptr% - 1
            Print #o%, vonlist(i%); " -> ";
            For j% = 0 To nptr% - 1
              If vonlist(i%) = nachlist(j%) Then
                Print #o%, nachlist(j%)
                nachlist(j%) = ""
                j% = 9999
              End If
            Next j%
            If j% < 9999 Then Print #o%, "_drop_data_"
          Next i%
          For i% = 0 To nptr% - 1
            If nachlist(i%) <> "" Then Print #o%, "_unavailable_ -> "; nachlist(i%)
          Next i%
          Print #o%,
          Print #o%, "; you may enter comments. they start with ;"
          Print #o%, "; lines starting with _unavailable_ are also ignored"
          Print #o%, "; _drop_data_ will not take over the field (data is lost)."
          Print #o%, "; data of unmentioned sourcefields will also be lost."
          Print #o%, "; an incorrect table will result in loosing data."
          Close #o%
          MsgBox "Unknown conversion. Please check " + tabf$ + vbCrLf + "then retry."
          rrr = Shell("notepad.exe " + tabf$, 1)
          Text1(5).text = p$
          BackColor = form1.cleancolor()
          MousePointer = 0
          Exit Sub
        Else
          o% = FreeFile
          c$ = "insert into usr_" & utabn(ntyp$) + " (id) values('" + id$ + "')"
          Call form1.sqlqry(c$)
          Open tabf$ For Input As #o%
          While Not EOF(o%)
            Line Input #o%, d$
            d$ = trm(d$)
            If d$ <> "" Then
            If Left$(d$, 1) <> ";" Then
              c$ = cut_d1(d$, " ")
              d$ = trm(cut_d2bis(d$, ">"))
              If c$ <> "_unavailable_" Then
'Debug.Print c$; "->"; d$
                If d$ = "_drop_data_" Or d$ = "" Then
                  c$ = "delete from auftritthigru where auftrittstyp='" + p$ + "' and auftrittsid='" + id$ + "' and FeldName='" + c$ + "'"
                Else
                  wert$ = form1.get1erg("select FeldDaten as wert from auftritthigru where auftrittsid='" + id$ + "' and FeldName='" + c$ + "'")
                  l$ = "update usr_" & utabn(ntyp$) + " set " + d$ + "='" + wert$ + "' where id='" + id$ + "'"
                  Call form1.sqlqry(l$)
                  If c$ <> d$ Then
                    c$ = "update auftritthigru set FeldName='" + d$ + "' where auftrittstyp='" + p$ + "' and auftrittsid='" + id$ + "' and FeldName='" + c$ + "'"
                  Else
                    c$ = ""
                  End If
                End If
                If c$ <> "" Then Call form1.sqlqry(c$)
              End If
            End If
            End If
          Wend
          Close #o%
          c$ = "update auftritthigru set auftrittstyp='" + ntyp$ + "' where auftrittstyp='" + p$ + "' and auftrittsid='" + id$ + "'"
          Call form1.sqlqry(c$)
          If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
            calcol.BackColor = form1.get_eventcolor(ntyp$)
            cmd$ = "update auftritt set optkalcolor='" + trm(calcol.BackColor) + "' where id='" + id$ + "';"
            Call form1.sqlqry(cmd$)
          End If
          c$ = "update auftritt set auftrittstyp='" + ntyp$ + "' where id='" + id$ + "'"
          Call form1.sqlqry(c$)
          c$ = "delete from usr_" & utabn(p$) + " where id='" + id$ + "'"
          Call form1.sqlqry(c$)
          unl = True
          BackColor = form1.cleancolor()
          If form1.kalopen Then Call kc.Command1_Click
        End If
      End If
    End If
  End If
End If
MousePointer = 0
If unl Then Call showrec(id$, 0)
End Sub

Private Sub Combo2_Change()
Dim i%
'd2infile = "auftritt": d2insub = "Combo2_Change"
For i% = 0 To Combo2.ListCount - 1
  If Combo2.List(i%) = Combo2.text Then
    Call form1.setusersetting("TerminDokumentname", Combo2.text)
    Exit For
  End If
Next i%
End Sub

Private Sub Combo2_Click()
'd2infile = "auftritt": d2insub = "Combo2_Click"
DoEvents
Call Combo2_Change
End Sub

Private Sub Combo2_DropDown()
'd2infile = "auftritt": d2insub = "Combo2_DropDown"
Combo2.Clear
Combo2.AddItem "Typ-Termindatum-User"
Combo2.AddItem formtranso(Label2(0).Caption) & transe("-Typ-Termindatum-User")
End Sub

Private Sub Command1_Click()
'd2infile = "auftritt": d2insub = "Command1_Click"
Hide
Unload auftritt

End Sub

Public Sub Command10_Click()
Dim r As ADODB.Recordset, ra As ADODB.Recordset, mws As Integer, fi As ADODB.Recordset
Dim ftst As ADODB.Recordset
Dim danz As Double, dnet As Double, d1 As Double, i As Integer, j As Integer, p As Integer, lblno As Integer
Dim gdcmd As String, viz As Boolean, rrr, c$, typ$, allgdcmd$, jj As Integer, fldlist As String
Dim t1$, t2$, t3$, plist As String, elist As String, aelist As String, wlist As String, evrd$
Dim dtg$, usr$, desc$, aalist$

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Command10_Click"
MousePointer = 11
evrd$ = LCase(form1.getusersetting("eventrequiresdate", "ja"))
If evrd$ <> "no" And evrd$ <> "nein" Then
  dtg$ = trm(Text1(2).text)
  If datum2sql(dtg$) < "1000" Or datum2sql(dtg$) > "3000" Then
    Text1(2).text = datum2sql(Date)
    Call Text1_LostFocus(2)
    DoEvents
  End If
End If
id$ = Text1(0).text
If recalcplease Then
  recalcplease = False
  Call recalc
End If
For i% = 0 To chgs.ListCount - 1
  If Left(chgs.List(i%), 8) <> "iCalUpd " Then
Debug.Print chgs.List(i%)
    Call form1.sqlqry(chgsread(i%))
  End If
Next i%
chgs.Clear: For i = 0 To 99: chgsx(i) = "": Next i
form1.sqlqry ("update auftritt set stand='" + datum2sql(Date) + " " + trm(Time) + "' where id='" + id$ + "'")
If delmode Then
  delmode = False
  Exit Sub
End If
typ$ = transo(Text1(5).text)
i = 1
Do
  On Error Resume Next
  viz = gd1(i).Visible
  rrr = Err
  On Error GoTo 0
  If rrr = 0 And viz Then
    lblno = labelnumbylistnum(i)
    c$ = "delete from auftritthigru where auftrittsid='" + id$ + "' and auftrittstyp='" + typ$ + "' and feldname='" + formtranso(Label2(lblno).Caption) + "'"
    Call form1.sqlqry(c$)
    allgdcmd$ = ""
    For j = 1 To gd1(i).ListItems.Count
      gd1(i).SelectedItem = gd1(i).ListItems(j)
      If gd1(i).SelectedItem <> "d" And gd1(i).SelectedItem <> "delete" Then
        gdcmd = ""
        For p = 1 To gd1(i).ColumnHeaders.Count - 1
          If gdcmd <> "" Then gdcmd = gdcmd + "|"
          gdcmd = gdcmd + gd1(i).SelectedItem.SubItems(p)
        Next p
        If allgdcmd$ <> "" Then allgdcmd$ = allgdcmd$ + vbCrLf
        allgdcmd$ = allgdcmd$ + gdcmd$
      End If
    Next j
    c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + _
    form1.newid("auftritthigru", "id", 60) + "','" + _
    id$ + "','" + _
    typ$ + "','" + _
    formtranso(Label2(lblno).Caption) + "','" + _
    allgdcmd$ + "');"
    Call form1.sqlqry(c$)
  End If
  i = i + 1
Loop Until rrr <> 0

c$ = Text3.text
s$ = "delete from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysez' and auftrittstyp='" + typ$ + "';"
Call form1.sqlqry(s$)
If c$ <> "" Then
  s$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values(" & _
          "'" + form1.newid("auftritthigru", "id", 50) + "','" + _
          id$ + "','" + typ$ + "','zzzsysez','" + _
          trm(c$) + "');"
  Call form1.sqlqry(s$)
End If
bez$ = Text1(6).text

If trm(mwst.text) = "" Then
  mws = form1.getusersetting("auftrittsmwst", "1900")
Else
  mws = Int(var2dbl(strrepl(mwst.text, ",", ".")) * 100)
  c$ = "select mwst from finanzen where id='" & id$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If r.EOF Then
    c$ = "insert into finanzen (id,mwst) values('" & id$ & "'," & trm(mws) & ")"
  Else
    c$ = "update finanzen set mwst=" & trm(mws) & " where id='" & id$ & "'"
  End If
  Call form1.sqlqry(c$)
End If

anz$ = "1"
an$ = ""
von$ = ""
net$ = ""
wae$ = ""
tut$ = ""
c$ = "select * from finanzen where id='" & id$ & "'"
Set fi = New ADODB.Recordset
fi.CursorLocation = adUseServer
rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If fi.EOF Then
  c$ = "insert into finanzen (id,mwst,anz) values('" & id$ & "'," & trm(form1.getusersetting("auftrittsmwst", 1900)) & ",1)"
  Call form1.sqlqry(c$)
  c$ = "select * from finanzen where id='" & id$ & "'"
  Set fi = New ADODB.Recordset
  fi.CursorLocation = adUseServer
rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
End If
If InStr(typ$, " ") = 0 And typ$ <> "" Then
c$ = "select * from usr_" & utabn(typ$) + " where id='" + id$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
On Error GoTo 0
If rrr <> 0 Then
  MousePointer = 0
  Exit Sub
End If
If r.EOF Then
  MousePointer = 0
  Exit Sub                                          'cannot be error
End If
'erstmal iCalupdate?
For i% = 1 To r.Fields.Count - 1
  '2b faster: wir nehmen an es sei eine adresse und testen lediglich die existenz des ikalenders
  On Error Resume Next
  c$ = trm(r.Fields(i%).value)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then c$ = ""
  If c$ <> "" Then Call form1.iCalUpdate(c$, id$)
Next i%

prototyp$ = typ$
If LCase(prototyp$) = "perfartist" Then prototyp$ = "künstlerauftritt"
If LCase(prototyp$) = "promo" Then prototyp$ = "künstlerauftritt"
Select Case LCase(prototyp$)
  Case "orchesterauftritt"
    If trm(fi!an) = "" Then an$ = trm(r!orchester)
    If trm(fi!von) = "" Then von$ = trm(r!veranstalter)
    net$ = form1.ohnewaehrung(trm(r!Honorar))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    tut$ = "Honorar " & trm(bez$)
    an2$ = ""
    von2$ = ""
    net2$ = ""
    wae2$ = ""
    tut2$ = ""
    an2$ = trm(r!künstler)
    von2$ = trm(r!veranstalter)
    On Error Resume Next
    net2$ = form1.ohnewaehrung(strrepl(trm(r!honorarkünstler), ".", ""))
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      wae2$ = form1.nurdiewaehrung(trm(r!honorarkünstler))
      tut2$ = "Honorar " & trm(bez$)
      ad$ = " where id='HonorarKünstler(ID:" & id$ & "'"
      Set ftst = New ADODB.Recordset
      ftst.CursorLocation = adUseServer
rrr = form1.adoopen(ftst, "select id from finanzen " & ad$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If ftst.EOF Then
        c$ = "insert into finanzen (id,mwst,anz) values('HonorarKünstler(ID:" & id$ & "'," & trm(form1.getusersetting("auftrittsmwst", 1900)) & ",1)":
        Call form1.sqlqry(c$)
      End If
      c$ = "update finanzen set an='" & Left(txt2db(an2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set von='" & Left(txt2db(von2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set netto=" & d2db(net2$) & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set waehrung='" & wae2$ & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set bezeichnung='" & tut2$ & "'" & ad$: Call form1.sqlqry(c$)
    End If
    an2$ = ""
    von2$ = ""
    net2$ = ""
    wae2$ = ""
    tut2$ = ""
    On Error Resume Next
    an2$ = trm(r!künstler2)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      von2$ = trm(r!veranstalter)
      net2$ = form1.ohnewaehrung(strrepl(trm(r!honorarkünstler2), ".", ""))
      wae2$ = form1.nurdiewaehrung(trm(r!honorarkünstler2))
      tut2$ = "Honorar " & trm(bez$)
      ad$ = " where id='HonorarKünstler2(ID:" & id$ & "'"
      Set ftst = New ADODB.Recordset
      ftst.CursorLocation = adUseServer
rrr = form1.adoopen(ftst, "select id from finanzen " & ad$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If ftst.EOF Then
        c$ = "insert into finanzen (id,mwst,anz) values('HonorarKünstler2(ID:" & id$ & "'," & trm(form1.getusersetting("auftrittsmwst", 1900)) & ",1)":
        Call form1.sqlqry(c$)
      End If
      c$ = "update finanzen set an='" & Left(txt2db(an2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set von='" & Left(txt2db(von2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set netto=" & d2db(net2$) & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set waehrung='" & wae2$ & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set bezeichnung='" & tut2$ & "'" & ad$: Call form1.sqlqry(c$)
    End If
    an2$ = ""
    von2$ = ""
    net2$ = ""
    wae2$ = ""
    tut2$ = ""
    On Error Resume Next
    an2$ = trm(r!künstler3)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      von2$ = trm(r!veranstalter)
      net2$ = form1.ohnewaehrung(strrepl(trm(r!honorarkünstler3), ".", ""))
      wae2$ = form1.nurdiewaehrung(strrepl(trm(r!honorarkünstler3), ",", ""))
      tut2$ = "Honorar " & trm(bez$)
      ad$ = " where id='HonorarKünstler3(ID:" & id$ & "'"
      Set ftst = New ADODB.Recordset
      ftst.CursorLocation = adUseServer
rrr = form1.adoopen(ftst, "select id from finanzen " & ad$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If ftst.EOF Then
        c$ = "insert into finanzen (id,mwst,anz) values('HonorarKünstler3(ID:" & id$ & "'," & trm(form1.getusersetting("auftrittsmwst", 1900)) & ",1)":
        Call form1.sqlqry(c$)
      End If
      c$ = "update finanzen set an='" & Left(txt2db(an2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set von='" & Left(txt2db(von2$), 70) & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set netto=" & d2db(net2$) & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set waehrung='" & wae2$ & "'" & ad$: Call form1.sqlqry(c$)
      c$ = "update finanzen set bezeichnung='" & tut2$ & "'" & ad$: Call form1.sqlqry(c$)
    End If
  Case "komposition"
    If trm(fi!an) = "" Then an$ = trm(r!Komponist)
    If trm(fi!von) = "" Then von$ = trm(r!auftraggeber)
    net$ = form1.ohnewaehrung(trm(r!Honorar))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    tut$ = "Honorar " & trm(bez$)
  Case "deal"
    If trm(fi!an) = "" Then an$ = trm(r!Lieferant)
    If trm(fi!von) = "" Then von$ = trm(r!Kunde)
    net$ = form1.ohnewaehrung(strrepl(trm(r!Honorar), ".", ""))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    tut$ = "Honorar " & trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & ",1)": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case "künstlerauftritt"
    If trm(fi!an) = "" Then
        If LCase(typ$) = "promo" Then
          an$ = trm(r!wer)
        Else
          an$ = trm(r!künstler)
        End If
    End If
    If trm(fi!von) = "" Then
      On Error Resume Next
      von$ = trm(r!veranstalter)
      On Error GoTo 0
    End If
    net$ = form1.ohnewaehrung(strrepl(trm(r!Honorar), ".", ""))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    tut$ = "Honorar " & trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & ",1)": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case "dirigentenauftritt"
    If trm(fi!an) = "" Then an$ = trm(r!dirigent)
    If trm(fi!von) = "" Then von$ = trm(r!veranstalter)
    net$ = form1.ohnewaehrung(strrepl(trm(r!Honorar), ".", ""))
    wae$ = form1.nurdiewaehrung(trm(r!Honorar))
    tut$ = "Honorar " & trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & ",1)": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case "hotelaufenthalt"
    If trm(fi!an) = "" Then an$ = trm(r!hotel)
    If trm(fi!von) = "" Then
      On Error Resume Next
      von$ = trm(r!wer)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then
        On Error Resume Next
        von$ = trm(r!künstler)
        On Error GoTo 0
      End If
    End If
    tut$ = "Hotelkosten " & trm(bez$)

    On Error Resume Next
    anz$ = trm(r!Einzelzimmer)
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      an$ = "": von$ = ""
    Else
      net$ = form1.ohnewaehrung(trm(r!EZ_Preis))
      wae$ = form1.nurdiewaehrung(trm(r!EZ_Preis))
      danz = var2dbl(word1(strrepl(anz$, ".", "")))
      dnet = var2dbl(word1(strrepl(net$, ".", "")))
      d1 = danz * dnet
      anz$ = trm(r!Doppelzimmer)
      net$ = form1.ohnewaehrung(trm(r!DZ_Preis))
      If wae$ = "" Then wae$ = form1.nurdiewaehrung(trm(r!DZ_Preis))
      danz = var2dbl(word1(strrepl(anz$, ".", "")))
      dnet = var2dbl(word1(strrepl(net$, ".", "")))
      d1 = d1 + danz * dnet
      anz$ = trm(r!suiten)
      net$ = form1.ohnewaehrung(trm(r!SU_Preis))
      If wae$ = "" Then wae$ = form1.nurdiewaehrung(trm(r!SU_Preis))
      danz = var2dbl(word1(strrepl(anz$, ".", "")))
      dnet = var2dbl(word1(strrepl(net$, ".", "")))
      d1 = d1 + danz * dnet
      anz$ = "1"
      net$ = fixeur(d1)
    End If
  Case "dienstleistung"
    an$ = trm(r!wer)
    von$ = trm(r!Kunde)
    net$ = form1.ohnewaehrung(strrepl(trm(r!betrag_pro_stunde), ".", ""))
    wae$ = form1.nurdiewaehrung(trm(r!betrag_pro_stunde))
    anz$ = word1(trm(r!Dauer)): If anz$ = "" Then anz$ = "0"
    tut$ = trm(bez$)
    ad$ = " where id='Honorar(ID:" & id$ & "'"
    c$ = "delete from finanzen " + ad$: Call form1.sqlqry(c$)
    c$ = "insert into finanzen (id,mwst,anz) values('Honorar(ID:" & id$ & "'," & trm(fi!mwst) & "," & strrepl(anz$, ",", ".") & ")": Call form1.sqlqry(c$)
    c$ = "update finanzen set an='" & Left(txt2db(an$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set von='" & Left(txt2db(von$), 70) & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set netto=" & d2db(net$) & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set waehrung='" & wae$ & "'" & ad$: Call form1.sqlqry(c$)
    c$ = "update finanzen set bezeichnung='" & tut$ & "'" & ad$: Call form1.sqlqry(c$)
  Case "verkauf"
    an$ = trm(r!verkäufer)
    von$ = trm(r!käufer)
    net$ = form1.ohnewaehrung(trm(r!einzelpreis))
    wae$ = form1.nurdiewaehrung(trm(r!einzelpreis))
    anz$ = word1(trm(r!anzahl))
    tut$ = trm(bez$)
  Case Else
    anz$ = "0"
End Select
ad$ = " where id='" & id$ & "'": ad1$ = "'" + ad$
w$ = trm(an$)
If w$ <> "" Then
  c$ = "update finanzen set an='" & Left(txt2db(w$), 70) & ad1$
  Call form1.sqlqry(c$)
End If
w$ = trm(von$)
If w$ <> "" Then
  c$ = "update finanzen set von='" & Left(txt2db(w$), 70) & ad1$
  Call form1.sqlqry(c$)
End If
w$ = trm(net$)
If w$ <> "" Then
  c$ = "update finanzen set netto=" & d2db(word1(strrepl(strrepl(w$, ".", ""), ",", "."))) & ad$
  Call form1.sqlqry(c$)
End If
w$ = trm(wae$)
If w$ <> "" Then
  c$ = "update finanzen set waehrung='" & trm(w$) & ad1$
  Call form1.sqlqry(c$)
End If
w$ = trm(anz$): w$ = strrepl(strrepl(w$, ".", ""), ",", ".")
If w$ <> "" Then
  c$ = "update finanzen set anz=" & d2db(w$) & ad$
  Call form1.sqlqry(c$)
End If
w$ = trm(tut$)
If w$ <> "" Then
  c$ = "update finanzen set bezeichnung='" & txt2db(w$) & ad1$
  Call form1.sqlqry(c$)
End If
w$ = trm(typ$)
If w$ <> "" Then
  c$ = "update finanzen set typ='" & trm(w$) & ad1$
  Call form1.sqlqry(c$)
End If
End If
Call initfields(transo(Text1(5).text), Val(basemerk.Caption), 0)
BackColor = form1.cleancolor()
Command10.Enabled = False
Command13.Enabled = True
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click
If form1.priosopen Then Call prios.Command20_Click
If form1.check_tst(id$) Then
  Command21(0).Visible = False
  Command21(1).Visible = True
Else
  Command21(1).Visible = False
  Command21(0).Visible = True
End If

If form1.getusersetting("checkdates", "nein") = "ja" Then
  c$ = form1.checkdates(id$, trm(Text1(3).text), trm(Text3.text))
  If c$ <> "" Then MsgBox (transe("Es bestehen Terminkonflikte:") + vbCrLf + c$)
End If

form1.hordexlock = True
Call form1.event2cloud(id$)
form1.cldpusher.Interval = 1000
form1.hordexlock = False

'repertoirecheck
If Not form1.isfieldmissing("opt_repertoire", "id") Then
c$ = "select id from auftrittsfelder where typ='" + typ$ + "' and FeldName like '%Programm%'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
On Error GoTo 0
If rrr = 0 And (Not r.EOF) Then
  c$ = "select FeldName from auftrittsfelder where typ='" + typ$ + "'"
  Set ftst = New ADODB.Recordset
  ftst.CursorLocation = adUseServer
  rrr = form1.adoopen(ftst, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 And (Not ftst.EOF) Then
    'programme und Personen sammeln
    plist = "": elist = "": aelist = "": fldlist = ""
    While Not ftst.EOF
      t1$ = cut_d1(trm(ftst!feldname), ".")
      t2$ = cut_d2bis(trm(ftst!feldname), ".")
      t3$ = cut_d2bis(t2$, ".")
      t2$ = cut_d1(t2$, ".")
      If t1$ = "programm" Then plist = plist + "|" + t2$
      If t1$ = "adrselect" Then
        aelist = aelist + "|" + t2$
        If InStr(LCase(t3$), "künstler") > 0 Or InStr(LCase(t3$), "dirigent") > 0 Or InStr(LCase(t3$), "orchester") > 0 Then elist = elist + "|" + t2$
      End If
      ftst.MoveNext
    Wend
    If plist <> "" And elist <> "" Then     'timesaver
      'werke sammeln
      While plist <> ""
        t1$ = cut_d1(plist, "|"): plist = cut_d2bis(plist, "|")
        If t1$ <> "" Then
          c$ = "select FeldDaten from auftritthigru where FeldName='" + t1$ + "' and auftrittsid='" + id$ + "'"
          Set fi = New ADODB.Recordset
          fi.CursorLocation = adUseServer
          rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If rrr = 0 And (Not fi.EOF) Then
            wlist = wlist + "|" + form1.getwerkids(trm(fi!felddaten))
          End If
        End If
      Wend
      
      If wlist <> "" Then
      If Not form1.isfieldmissing("opt_repertoire", "id") Then
      Load repertoire: DoEvents
      repertoire.artid.Caption = "repertoire_addmode": DoEvents
      While elist <> ""
        t1$ = cut_d1(elist, "|"): elist = cut_d2bis(elist, "|")
        If t1$ <> "" Then
          c$ = "select FeldDaten from auftritthigru where FeldName='" + t1$ + "' and auftrittsid='" + id$ + "'"
          Set fi = New ADODB.Recordset
          fi.CursorLocation = adUseServer
          rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          If rrr = 0 And (Not fi.EOF) Then
            c$ = trm(fi!felddaten)
            If InStr(c$, "{") > 0 Then
              c$ = cut_d2bis(c$, "{")
              c$ = cut_d1(c$, "}")
            End If
            repertoire.List2.AddItem c$: DoEvents
            t2$ = wlist
            While t2$ <> ""
              t3$ = cut_d1(t2$, "|"): t2$ = cut_d2bis(t2$, "|")
              If t3$ <> "" Then
                i% = form1.reptest(c$, t3$)
                If i% < 0 Then
                  For jj = 0 To repertoire.List3.ListCount - 1
                    If InStr(repertoire.List3.List(jj), c$ + "|" + t3$) = 1 Then Exit For
                  Next jj
                  If jj >= repertoire.List3.ListCount Then repertoire.List3.AddItem c$ + "|" + t3$ + "|0": DoEvents
                End If
              End If
            Wend
          End If
        End If
      Wend
      If repertoire.List2.ListCount > 0 Then
        On Error Resume Next
        Call repertoire.SetFocus
        On Error GoTo 0
        repertoire.List2.ListIndex = 0
      Else
        Unload repertoire
      End If
      
      End If          'wlist<>""
      End If
    End If
  End If
End If
End If
If Not form1.isfieldmissing("opt_checklists", "id") Then Call shw_reminders
MousePointer = 0
End Sub

Private Sub Command11_Click()
'd2infile = "auftritt": d2insub = "Command11_Click"
Call savecheck
nbase% = 0

Call initfields(transo(Text1(5).text), nbase%, 0)

End Sub

Private Sub Command12_Click()

'd2infile = "auftritt": d2insub = "Command12_Click"
Call savecheck
nbase% = form1.sqla.TableDefs("usr_" & utabn(transo(Text1(5).text))).Fields.Count - 17

Call initfields(transo(Text1(5).text), nbase%, 0)
End Sub


Private Sub Command13_Click()

'd2infile = "auftritt": d2insub = "Command13_Click"
Call savecheck
id$ = Text1(0).text
If id$ = "" Then Exit Sub
Load fdet
Call fdet.SetFocus
fdet.fid = id$

End Sub

Private Sub Command14_Click()
'd2infile = "auftritt": d2insub = "Command14_Click"
Load auftrittrepeat
Call auftrittrepeat.SetFocus

End Sub

Private Sub Command15_Click()
Dim stmp As ADODB.Recordset, r1 As ADODB.Recordset, c$, c1$, id$, i%, q As Integer
Dim rrr

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Command15_Click"
id$ = Text1(0).text
If id$ = "" Then Exit Sub
  MousePointer = 11: DoEvents
  On Error Resume Next
  Kill form1.mydatadir() & "\*.sql"
  On Error GoTo 0
  tb$ = "auftritt"
  Call form1.sqlex_adresse("auftritt", "id", id$)
  Call form1.sqlex_adresse("finanzen", "id", id$)
  Call form1.sqlex_adresse("usr_" & utabn(trm(transo(Text1(5).text))), "id", id$)
  tb$ = "auftritthigru"
  fn$ = form1.mydatadir() & "\" & tb$ & "_" & form1.mkfn(id$) & ".sql"
  If exist(fn$) = 0 Then
    o% = FreeFile
    Open fn$ For Output As #o%
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT * FROM auftritthigru where auftrittsid ='" + id$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      c$ = "insert into auftritthigru (id) values('" & stmp.Fields(0).value & "');"
      Print #o%, c$
      If Len(trm(stmp!felddaten)) < 80 Then
        c$ = "select id from adresse where id='" & trm(stmp!felddaten) & "'"
        Set r1 = New ADODB.Recordset
        r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r1.EOF Then
          Call form1.sqlex_adresse("adresse", "id", trm(stmp!felddaten))
        End If
      End If
      c$ = trm(stmp!felddaten)
      If InStr(c$, "|") > 0 Then
        For i% = 1 To linesof(c$)
          c1$ = lineof(i%, c$)
          q = InStr(c1$, "|")
          If q > 1 Then
            c1$ = Left(c1$, q - 1)
            Call form1.sqlex_adresse("adresse", "id", c1$)
          End If
        Next i%
      End If
      For i% = 1 To stmp.Fields.Count - 1
        If trm(stmp.Fields(i%).value) <> "" Then
        If LCase(stmp.Fields(i%).name) <> "tourneeplaniD" And _
           LCase(stmp.Fields(i%).name) <> "bezeichnung" And _
           LCase(stmp.Fields(i%).name) <> "datum" And _
           LCase(stmp.Fields(i%).name) <> "ort" And _
           LCase(stmp.Fields(i%).name) <> "tourneeplanid" And _
           LCase(stmp.Fields(i%).name) <> "zeit" And _
           LCase(stmp.Fields(i%).name) <> "stand" And _
           LCase(stmp.Fields(i%).name) <> "astatus" Then
           c$ = form1.mkupdcmd(tb$, "id", stmp.Fields(0).value, stmp.Fields(i%).name, stmp.Fields(i%).Type, stmp.Fields(i%).value) & ";"
           Print #o%, c$
        End If
        End If
      Next i%
      stmp.MoveNext
    Wend
    Close #o%
  End If
  MousePointer = 0: DoEvents
  smtp.txtMessageSubject = "Agencyprof Datenpakete Saalplan " & h$ & " " & p$
  smtp.txtMessageText = "Speichern Sie das Attachment in Ihrem Agencyprof-Verzeichnis"
  tg$ = Dir(form1.mydatadir() & "\*.sql")
  While tg$ <> ""
    Call smtp.attachfile(form1.mydatadir() & "\" & tg$)
    tg$ = Dir
  Wend
End Sub

Private Sub Command16_Click()
'd2infile = "auftritt": d2insub = "Command16_Click"
  dtg$ = Text1(2).text
  If dtg$ <> "" Then
    Load kc
'select: nur Kalendereinträge für vorkommende
    If krcount% > 0 And kalres.value <> 0 Then
      kc.selct(2).Clear
      For i% = 0 To krcount% - 1
        kc.selct(2).AddItem kalrestrict$(i%)
      Next i%
    End If
    Call kc.settag0(dtg$)
    'If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then kc.Text1.Text = 1
    On Error Resume Next
    Call kc.SetFocus
    Call k3.SetFocus
    On Error GoTo 0
  End If

End Sub

Private Sub Command17_Click()
Dim vorlage$, o%, nf$, p%, l$, c$, fldn$, q%, f0$, dne As Boolean
Dim rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Command17_Click"
vorlage$ = form1.vorlagenverzeichnis() + "\v0__auftritt.rtf"
If nexist(vorlage$) Then
  MsgBox transe("Die Vorlage") + " " + vorlage$ + " " + transe("existiert nicht") + "."
  Exit Sub
End If
MousePointer = 11: DoEvents
nf$ = form1.vorlagenverzeichnis() + "\" & transo(Text1(5).text) & "_neue_vorlage.rtf"
o% = FreeFile
Open vorlage$ For Input As #o%
p% = FreeFile
Open nf$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  If InStr(l$, "feldliste") > 0 Then
    For i% = 1 To 7
      fldn$ = Label1(i%).Caption
      Print #p%, "\par " & fldn$ & ":{\*\bkmkstart " & fldn$ & "}" & fldn$ & "{\*\bkmkend " & fldn$ & "} "
    Next i%
    Print #p%, "\par "
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT feldname,zeilen FROM auftrittsfelder where typ='" & transo(Text1(5).text) & "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rtmp.EOF
      fldn$ = trm(rtmp!feldname)
      f0$ = fldn$
      q% = InStr(fldn$, ".")
      If q% > 0 Then
        fldn$ = Mid$(fldn$, q% + 1)
        fldn$ = cut_d1(fldn$, ".")
      End If
      dne = False
      If InStr(LCase(f0$), "adrselect.") = 1 Then
        dne = True
        w$ = "Name": Print #p%, "\par " & fldn$ & " \par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        w$ = "Strasse": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        Print #p%, "\par \tab PLZ/Ort:{\*\bkmkstart fx__plzort__" & fldn$ & "}fx__plzort__" & fldn$ & "{\*\bkmkend fx__plzort__" & fldn$ & "} "
        w$ = "Tel": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        w$ = "Handy": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        w$ = "Fax": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        w$ = "Email": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
        w$ = "URL": Print #p%, "\par \tab " & w$ & ":{\*\bkmkstart " & fldn$ & "__" & w$ & "}" & fldn$ & "__" & w$ & "{\*\bkmkend " & fldn$ & "__" & w$ & "} "
      End If
      If InStr(LCase(f0$), "programm.programm") = 1 Then
        dne = True
        Print #p%, "\par Programm:{\*\bkmkstart programm__text}Programm{\*\bkmkend programm__text} "
      End If
      If Not dne Then Print #p%, "\par " & fldn$ & ":{\*\bkmkstart " & fldn$ & "}" & fldn$ & "{\*\bkmkend " & fldn$ & "} "
      rtmp.MoveNext
    Wend
  Else
    Print #p%, l$
  End If
Wend
Close #o%
Close #p%
MousePointer = 0: DoEvents
Call showrec(Text1(0).text, 0)
Call form1.openthisdoc(nf$, "")

End Sub

Private Sub Command18_Click()

Call form1.handbuchcall("10-Termine.htm")

End Sub

Private Sub Command19_Click()

Load dayvw
On Error Resume Next
Call dayvw.SetFocus
On Error GoTo 0
dayvw.Text1.text = Text1(2).text

End Sub

Private Sub Command2_Click()
Dim cmd$, r As ADODB.Recordset, wert$, id$, Index As Integer, rrr

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Command2_Click"
id$ = Text1(0).text
If id$ = "" Then
  Exit Sub
End If

Index = 1
wert$ = Text1(Index).text
wert$ = trm(InputBox(transe("Projekt:") + " " + wert$, transe("In anderes Projekt verschieben"), wert$))
If wert$ = "" Then Exit Sub
If wert = "-1" Then
  cmd$ = "update auftritt set TourneeplanID='" + wert$ + "' where id='" + id$ + "'"
  Call form1.sqlqry(cmd$)
  Text1(Index).text = wert$
  Call Command10_Click
  Exit Sub
End If
cmd$ = "SELECT id FROM tplan where id='" & wert$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then
  MsgBox transe("Das Projekt") + " " + wert$ + " " + transe("existiert nicht") + "."
  Exit Sub
Else
  cmd$ = "update auftritt set TourneeplanID='" + wert$ + "' where id='" + id$ + "'"
  Call form1.sqlqry(cmd$)
  Text1(Index).text = wert$
  Call Command10_Click
End If

End Sub

Private Sub Command20_Click()

Call recalc

End Sub

Public Sub Command21_Click(Index As Integer)
Dim c$, id$

If form1.isfieldmissing("opt_checklists", "id") Then Exit Sub
If Index = 999 Then listMessages.Visible = False
DoEvents
id$ = Text1(0).text
If id$ <> "" Then Call achktst(id$)
If listMessages.Visible Then
  listMessages.Visible = False
  chkokall.Visible = False
  chkdlall.Visible = False
  chkudall.Visible = False
  chkedone.Visible = False
  chkeclse.Visible = False
  chkown.Visible = False
  chksve.Visible = False
  chkedt.Visible = False
  uselct.Visible = False
Else
  If id$ = "" Then Exit Sub
  listMessages.Top = 1080
  listMessages.Height = Me.Height - 1200 - 1080
  listMessages.Left = Command21(0).Left + Command21(0).Width + 60
  listMessages.ListItems.Clear
  listMessages.Visible = True
  chkeclse.Top = listMessages.Top + listMessages.Height - chkokall.Height - 120
  chkeclse.Left = listMessages.Left + 120
  chkeclse.Visible = True
  chkokall.Top = listMessages.Top + listMessages.Height - chkokall.Height - 120
  chkokall.Left = chkeclse.Left + 120 + chkeclse.Width
  chkokall.Visible = True
  chkdlall.Top = chkokall.Top
  chkdlall.Left = chkokall.Left + 120 + chkokall.Width
  chkdlall.Visible = True
  chkudall.Top = chkdlall.Top
  chkudall.Left = chkdlall.Left + 120 + chkdlall.Width
  chkudall.Visible = True
  chkedone.Top = chkudall.Top
  chkedone.Left = chkudall.Left + 120 + chkudall.Width
  chkedone.Visible = True
  chkown.Left = chkedone.Left + 120 + chkedone.Width
  chkown.Top = chkudall.Top
  chksve.Top = chkudall.Top
  chksve.Left = chkown.Left + 60 + chkown.Width
  chkedt.Top = chksve.Top
  chkedt.Left = chksve.Left + 60 + chksve.Width
  chkedt.Visible = True
  Call shw_reminders
End If
End Sub

Public Sub shw_reminders()
Dim c$, id$, i%, sownr$, chgs1 As Boolean
Dim rrr, dto, dtg$, r$
Dim s As ADODB.Recordset
Dim d2infile As String, d2insub As String
  
id$ = Text1(0).text
chgs1 = False
If id$ <> "" Then Call achktst(id$)
  listMessages.ListItems.Clear
  c$ = "SELECT checkpoint as chkp, checkid,auftrittsid, confirmed, id, dtg, ownr FROM opt_checks WHERE auftrittsid='" + id$ + "' order by dtg"
  Set s = New ADODB.Recordset
Debug.Print c$
  s.CursorLocation = adUseServer
  rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not s.EOF
    chktxt$ = trm(s!chkp)
    If chktxt = "" Then
      chktxt = form1.get_defaultchecklisttext(trm(s!checkid))
    End If
'Debug.Print s!dtg + " " + chktxt$ + " Confirmed: " + trm(s!confirmed) + Space$(160) + "(ID:" + s!id
    If InStr(trm(s!confirmed), "ok, deleted") <> 1 Then
      Set lvitem = listMessages.ListItems.add(, , datfromsql(s!dtg))
      lvitem.Selected = False
      lvitem.SubItems(1) = chktxt$
      lvitem.SubItems(2) = trm(s!confirmed)
      sownr$ = trm(strrepl(trm(s!ownr), "|", " "))
      If Left$(sownr$, 1) = "{" Then
        c$ = "select FeldDaten as wert from auftritthigru where auftrittsid='" + Text1(0).text + "' and FeldName='" + Mid$(sownr$, 2) + "'"
        r$ = form1.get1erg(c$)
        If r$ <> "" Then r$ = form1.APUsernameByAddressID(r$)
      End If
      If r$ <> "" Then
        sownr$ = r$
        chgs1 = True
      End If
      lvitem.SubItems(3) = sownr$
      lvitem.SubItems(4) = s!id
      If Left(trm(s!confirmed), 2) <> "ok" And datum2sql(trm(Date)) >= s!dtg Then
        listMessages.ListItems.Item(listMessages.ListItems.Count).Bold = True
      End If
      If chgs1 Then
        c$ = "update opt_checks set ownr='" + lvitem.SubItems(3) + "' where id='" + lvitem.SubItems(4) + "'"
        Call form1.sqlqry(c$)
      End If
    End If
    DoEvents
    s.MoveNext
  Wend

End Sub

Private Sub Command22_Click()
Dim add$

If form1.isfieldmissing("opt_othertplans", "id") Then
  add$ = "CREATE TABLE `opt_othertplans` (" + vbCrLf
  add$ = add$ + "`id` INT(11) NOT NULL AUTO_INCREMENT," + vbCrLf
  add$ = add$ + "`tplanid` VARCHAR(50) NOT NULL DEFAULT '0'," + vbCrLf
  add$ = add$ + "`aid` VARCHAR(120) NOT NULL DEFAULT '0'," + vbCrLf
  add$ = add$ + "PRIMARY KEY (`id`)," + vbCrLf
  add$ = add$ + "INDEX `tplanid` (`tplanid`)," + vbCrLf
  add$ = add$ + "INDEX `aid` (`aid`)" + vbCrLf
  add$ = add$ + ")"
  MsgBox "Table opt_othertplans is missing." + vbCrLf + "Contact support if needed."
  Exit Sub
End If
Load tpsel
Call tpsel.init(Text1(0).text)
End Sub

Private Sub Command3_Click(Index As Integer)
Dim sid$, sidk$, sida$, p%, X, fn$

'd2infile = "auftritt": d2insub = "Command3_Click"
i% = Index
If clickgetsfromtable(i%) = "" Then Exit Sub
Call form1.dbg2f("auftritt:cmd3_click:" & Index)
If LCase(clickgetsfromtable(i%)) = "finanzen" Then
  s$ = clickgetsfromfield(i%)
  id$ = Text1(0).text
  If id$ = "" Then Exit Sub
  Call savecheck
  Load fdet
  Call fdet.SetFocus
  fdet.fid = s$ & "(ID:" & id$
  DoEvents
  fdet.xpara.Caption = Text2(i%).text
End If
If LCase(clickgetsfromtable(i%)) = "vertragsnummer" Then
  Call Label2_DblClick(i%)
End If
If LCase(clickgetsfromtable(i%)) = "adrselect" Then
  sid$ = Text2(i%).text
  p% = InStr(sid$, "{")
  sida$ = sid$: sidk$ = ""
  If p% > 0 Then
    sidk$ = trm(Left(sid$, p% - 1))
    sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If Len(sida$) > 0 Then
    If sidk$ = "" And InStr(LCase$(transo(formtranso(Label2(i% + 1).Caption))), "kontakt") = 1 Then sidk$ = Text2(i% + 1).text
    Load shwAdrDetail
    Call shwAdrDetail.refreshadrdetail(sida$, sidk$)
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
  Else
    Call Label2_DblClick(i%)
  End If
End If
If LCase(clickgetsfromtable(i%)) = "tabelle" Then Call Label2_DblClick(i%)
If LCase(clickgetsfromtable(i%)) = "programm" Then
  sid$ = Text2(i%).text
  s0$ = sid$
  If LCase(Left$(s0$, 6)) = "datei:" Or LCase(Left$(s0$, 3)) = "fn:" Then
    p% = InStr(s0$, ":")
    fn$ = ""
    If p% > 0 Then
      fn$ = trm(Mid$(s0$, p% + 1))
      If Left$(fn$, 1) = "(" Then fn = cut_d2bis(fn$, ")")
      If nexist(fn$) Then
Call form1.dbg2f("openrequest: (not foumd) " + fn$)
        If Not nexist(form1.filenamekurz(fn$)) Then fn$ = form1.filenamekurz(fn$)
      End If
Call form1.dbg2f("openrequest: " + fn$)
      If fn$ = "" Or nexist(fn$) Then
        fn$ = form1.myuniquedocname("noask") & ".txt"
        fn$ = form1.filenamekurz(form1.saveasBox(fn$))
      End If
      If fn$ <> "" Then
        Call form1.openthisdoc(fn$, "")
      End If
    End If
  Else
    If Len(sid$) > 0 Then
      Load prog
      Call prog.SetFocus
      Call prog.selectone(sid$)
    Else
      Call Label2_DblClick(i%)
    End If
  End If
End If
End Sub

Private Sub Command4_Click()
'd2infile = "auftritt": d2insub = "Command4_Click"
Call savecheck
nbase% = Val(Label4.Caption)

Call initfields(transo(Text1(5).text), nbase%, 0)

End Sub

Private Sub Command42_Click()
Dim r As ADODB.Recordset, c$, rrr, dst$, difft$

Command42.BackColor = &HC0C0C0
form1.tlopen = False
c$ = "SELECT * FROM auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='umkrchk' and auftrittstyp='" & transo(Text1(5).text) & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not r.EOF Then
    c$ = "delete from auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='umkrchk' and auftrittstyp='" & transo(Text1(5).text) & "'"
    Call form1.sqlqry(c$)
  Else
    dst$ = form1.mylastFormVar(Me.name, "umkrsuch", "30")
    dst$ = trm(InputBox(transe("Suche im Umkreis (km):"), transe("Umkreissuche"), dst$))
    c$ = "delete from auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='umkrchk' and auftrittstyp='" & transo(Text1(5).text) & "'"
    Call form1.sqlqry(c$)
    c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
       form1.newid("auftritthigru", "id", 50) & "','" & _
       Text1(0).text & "','" & _
       transo(Text1(5).text) & "','umkrchk','" & _
       dst$ & "')"
    Call form1.sqlqry(c$)
    Call umkrtest(Text1(0).text, dst$)
  End If
End If
r.Close

End Sub

Private Sub Command42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
  If Command42.BackColor = &HFF& And Not form1.tlopen Then
    form1.tlopen = True
    tlform.Top = Command42.Top + Me.Top + 200
    tlform.Left = Command42.Left + Me.Left
    tlform.Show
  End If
End Sub

Private Sub Command5_Click()
'd2infile = "auftritt": d2insub = "Command5_Click"
Call savecheck
nbase% = Val(Label3.Caption)

Call initfields(transo(Text1(5).text), nbase%, 0)

End Sub

Private Sub Command6_Click()
Dim v0$, o%

'd2infile = "auftritt": d2insub = "Command6_Click"
Call form1.setAuftrittsdruckFuerAdresse("")
Call savecheck
auftritt.MousePointer = 11
id$ = Text1(0).text
v0$ = transo(Text1(5).text)
Call form1.delalias
listenhauptperson = ""
If id$ <> "" Then
  i% = List1.ListIndex
  If i% >= 0 Then
    voralias$ = form1.vorlagenverzeichnis() + "\" + v0$ + "_" + List1.List(List1.ListIndex) + ".alias"
    If nexist(voralias$) Then
      vorlage$ = form1.vorlagenverzeichnis() + "\" + List1.List(List1.ListIndex) + ".rtf"
      If InStr(List1.List(List1.ListIndex), v0$) = 0 Then
        vorlage$ = form1.vorlagenverzeichnis() + "\" + v0$ + "_" + List1.List(List1.ListIndex) + ".rtf"
      End If
    Else
      f$ = form1.readalias(voralias$)
      If f$ <> "" Then vorlage$ = form1.vorlagenverzeichnis() + "\" + v0$ + "_" + f$ + ".rtf"
    End If
  End If
  If nexist(vorlage$) Then
    vorlage$ = form1.vorlagenverzeichnis() + "\" + v0$ + ".rtf"
  End If
  form1.honorarlcount% = 0
  If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then
    Load dbupgrade
    dbupgrade.Caption = "Dokument wird erstellt ..."
    Call dbupgrade.SetFocus
  End If
  form1.skip1del = True
  Call form1.dbg2f("calling auftrttsdruck(" + id$ + "," + vorlage$ + ",...)")
  Call form1.auftrittsdruck(id$, vorlage$, "auftritt", "")
  Call form1.dbg2f("returned from auftrttsdruck(" + id$ + "," + vorlage$ + ",...)")
  form1.skip1del = False
End If
Call form1.dbg2f("unloading tplan")
On Error Resume Next
Unload tplan
On Error GoTo 0
Call form1.dbg2f("unloaded tplan")
auftritt.MousePointer = 0
listenhauptperson = ""
Call form1.dbg2f("auftritt:command6-click done")
End Sub

Public Sub Command7_Click()
Dim r As ADODB.Recordset, rrr
Dim ast As Integer, thiscolor

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Command7_Click"
oldid$ = trm(Text1(0).text)
If oldid$ = "" Then Exit Sub
typ$ = transo(Text1(5).text)
If LCase(typ$) = "neuer auftritt" Then Exit Sub
id$ = form1.newid("auftritt", "id", 20)
ast = form1.auftrittsstatus(oldid$)
form1.sqlqry ("INSERT INTO usr_" & utabn(typ$) + " (id) VALUES ('" + id$ + "')")
form1.sqlqry ("INSERT INTO auftritt (id) VALUES ('" + id$ + "')")
If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
  Call form1.sqlqry("update auftritt set optkalcolor='" & trm(calcol.BackColor) & "' where id='" & id$ & "'")
End If
Call form1.sqlqry("update auftritt set astatus=" & trm(ast) & " where id='" & id$ & "'")
For Index = 1 To 7
If Index <> 4 Then
nwert$ = trm(Text1(Index).text)
'If nwert$ <> prv$ Then
  fld$ = transo(Label1(Index).Caption)
  If LCase(fld$) = "auftrittstyp" Then nwert$ = transo(nwert$)
  If Index = 2 Then nwert$ = datum2sql(nwert$)
  If Index = 6 Then nwert$ = nwert$
  If nwert$ = "" Then
    nwert$ = "NULL"
  Else
    nwert$ = "'" + nwert$ + "'"
  End If
  Call makemedirty
  Call chgswrite("update auftritt set " + fld$ + "=" + nwert$ + " where id='" + id$ + "'")
'End If
End If
Next Index
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM usr_" & utabn(typ$) + " where id='" + oldid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not r.EOF Then
  anz = form1.sqla.TableDefs("usr_" & utabn(typ$)).Fields.Count - 1
  For i% = 1 To anz
    On Error Resume Next
    wert$ = r.Fields(i%).value
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      cmd$ = "update usr_" & utabn(typ$) + " set " + r.Fields(i%).name + "='" + wert$ + "' where id='" + id$ + "'"
      Call makemedirty
      Call chgswrite(cmd$)
    End If
  Next i%
End If

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM auftritthigru where auftrittstyp='" & typ$ & "' and auftrittsid='" + oldid$ + "'"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If trm(r!felddaten) <> "" And trm(r!feldname) <> "" Then
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
       form1.newid("auftritthigru", "id", 50) & "','" & _
       id$ & "','" & _
       r!auftrittstyp & "','" & _
       r!feldname & "','" & _
       r!felddaten & "')"
    Call form1.sqlqry(cmd$)
  End If
  r.MoveNext
Wend

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM auftritthigru where auftrittstyp='kalku_" & typ$ + "' and auftrittsid='" + oldid$ + "'"
'cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + oldid$ + "'"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If trm(r!felddaten) <> "" And trm(r!feldname) <> "" Then
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
       form1.newid("auftritthigru", "id", 50) & "','" & _
       id$ & "','" & _
       r!auftrittstyp & "','" & _
       r!feldname & "','" & _
       r!felddaten & "')"
    Call form1.sqlqry(cmd$)
  End If
  r.MoveNext
Wend
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM auftritthigru where instr(auftrittstyp,'tabkalk_" & typ$ & "_')=1 and auftrittsid='" + oldid$ + "'"
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If trm(r!felddaten) <> "" And trm(r!feldname) <> "" Then
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
       form1.newid("auftritthigru", "id", 50) & "','" & _
       id$ & "','" & _
       r!auftrittstyp & "','" & _
       r!feldname & "','" & _
       r!felddaten & "')"
    Call form1.sqlqry(cmd$)
  End If
  r.MoveNext
Wend
form1.fastsave_copy = True
Call showrec(id$, 0)
form1.fastsave_copy = False
Call Command10_Click

End Sub

Public Sub Command8_Click()

'd2infile = "auftritt": d2insub = "Command8_Click"
For i% = 0 To fpp%

If Text2(i%).Enabled = False Then
  Text2(i%).Enabled = True
  Call Text2(i%).SetFocus
  prvd$ = "grmblfzzqwerty"
  Text2(i%).text = Text2(i%).text
  Call Text2_LostFocus(i%)
End If

Next i%

End Sub

Private Sub Command9_Click()
'd2infile = "auftritt": d2insub = "Command9_Click"
Load alarmlist
Call alarmlist.settab("usr_" & utabn(transo(Text1(5).text)))
alarmlist.Caption = "Auftritt-ID:" + Text1(0).text

End Sub

Private Sub dat_open_Click()
Call opendir_Click
End Sub

Private Sub delme_Click()
Dim id$, c$, r As ADODB.Recordset
Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "delme_Click"
id$ = Text1(0).text
If id$ = "" Then Exit Sub

If Not form1.isfieldmissing("opt_vnr", "id") Then
  c$ = "select id from opt_vnr where aid='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    MsgBox (transe("Dieser Auftritt ist verbunden mit der Vertragsnummer " + trm(r!id) + " und kann nicht gelöscht werden."))
    Exit Sub
  End If
End If
antw = MsgBox(transe("Diesen Auftritt löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Daten löschen?"))
If antw = vbYes Then
  If form1.cloud Then Call form1.event2cloudremove(id$)
  Call makemedirty
  If Not form1.isfieldmissing("opt_checks", "id") Then
    c$ = "delete from opt_checks where auftrittsid='" + id$ + "'"
    Call form1.sqlqry(c$)
  End If
  If Not form1.isfieldmissing("opt_prios", "id") Then
    c$ = "delete from opt_prios where evnt='E:" + id$ + "'"
    Call form1.sqlqry(c$)
  End If
  typ$ = transo(Text1(5).text)
  c$ = "delete from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysez' and auftrittstyp='" + typ$ + "';": Call form1.sqlqry(c$)
  c$ = "delete from b_loc where wid='T:" + id$ + "';": Call form1.sqlqry(c$)
  c$ = "select * from usr_" & utabn(typ$) + " where id='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
  If Not r.EOF Then
    For i% = 1 To r.Fields.Count - 1
      On Error Resume Next
      c$ = trm(r.Fields(i%).value)
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        If c$ <> "" Then Call form1.iCalDelTermin(c$, id$)
      End If
    Next i%
  End If
  End If
  chgs.AddItem "delete from auftritt where id='" & id$ & "'"
  chgs.AddItem "delete from todolist where Betreff='[Wiedervorlage] AT:" + id$ + "'"
  chgs.AddItem "delete from finanzen where id='" & id$ & "'"
  chgs.AddItem "delete from auftritthigru where auftrittsid='" & id$ & "'"
  If transo(Text1(5).text) <> "Neuer Auftritt" Then
    chgs.AddItem "delete from usr_" & utabn(transo(Text1(5).text)) + " where id='" & id$ & "'"
  End If
  delmode = True
  Unload Me
End If
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub dtst_Click()
Dim s$, r As ADODB.Recordset, nwert$, Index As Integer, rwert$, rrr

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "dtst_Click"
id$ = Text1(0).text
typ$ = transo(Text1(5).text)
For Index = 0 To angezeigtefelder% - 1
  fld$ = transo(formtranso(Label2(Index).Caption))
  nwert$ = trm(Text2(Index).text)
  If nwert$ <> "" Then nwert$ = "'" + nwert$ + "'"
  cmd$ = "select " & LCase(fld$) & " as usrwert from usr_" & utabn(typ$) & " where id='" & id$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
  If Not r.EOF Then
    rwert$ = trm(r!usrwert)
    If rwert$ <> "" Then rwert$ = "'" + rwert$ + "'"
    If rwert$ <> nwert$ Then
      cmd$ = "update usr_" & utabn(typ$) & " set " + fld$ + "=" & nwert$ & " where id='" & id$ & "'"
      Text2(Index).ForeColor = RGB(255, 0, 0)
    Else
      Text2(Index).ForeColor = RGB(0, 255, 0)
    End If
    DoEvents
  End If
  End If
Next Index
Timer_dtst.Interval = 1000
Timer_dtst.Enabled = True

End Sub

Private Sub findfeld_Click()
Dim i%, sfld$, sfld0$, j%

'd2infile = "auftritt": d2insub = "findfeld_Click"
findfeld.Width = 390
i% = findfeld.ListIndex
If i% < 0 Then Exit Sub
sfld$ = findfeld.List(i%)
If InStr(sfld$, "ID") = 1 Then
  Clipboard.Clear
  Clipboard.settext Text1(0).text
  MsgBox "EventID " + Text1(0).text + vbCrLf + "was copied to clipboard."
  Exit Sub
End If
i% = InStr(sfld$, "(POS=")
sfld0$ = trm(Left(sfld$, i% - 1))
If i% = 0 Then Exit Sub
i% = Val(Mid$(sfld$, i% + 5)) - 2
If i% < 0 Then i% = 0
Call savecheck
nbase% = Val(Label4.Caption)
For j% = 0 To fpp%
  Label2(j%).ForeColor = &H80000012
  Label2(j%).Visible = False
  Text2(j%).Visible = False
  Command3(j%).Visible = False
Next j%
Call initfields(transo(Text1(5).text), i%, 0)
nbase% = i%
For j% = 0 To 9
  If LCase(sfld0$) = LCase(formtranso(Label2(j%).Caption)) Then
    On Error Resume Next
    Call Text2(j%).SetFocus
    On Error GoTo 0
    Exit Sub
  End If
Next j%
End Sub

Private Sub findfeld_DropDown()
Dim r As ADODB.Recordset, lfdn As Integer, rrr

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "findfeld_DropDown"
findfeld.Width = 2655
findfeld.Clear
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT feldname,position,zeilen FROM auftrittsfelder where typ='" & transo(Text1(5).text) & "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then Exit Sub
lfdn = 0
While Not r.EOF
  lfdn = lfdn + 1
  If r!zeilen > 0 Then
    fn$ = trm(r!feldname)
    p% = InStr(fn$, ".")
    If p% > 0 Then
      fn$ = Mid$(fn$, p% + 1)
      p% = InStr(fn$, ".")
      If p% > 0 Then
        fn$ = Left$(fn$, p% - 1)
      End If
    End If
    findfeld.AddItem transe(fn$) & Space$(40) & "(POS=" & trm(lfdn)
  End If
  r.MoveNext
Wend
findfeld.AddItem "ID" & Space$(40) & "(POS=-1"
End Sub

Private Sub findfeld_GotFocus()
'd2infile = "auftritt": d2insub = "findfeld_GotFocus"
Call findfeld_DropDown
End Sub

Private Sub findfeld_LostFocus()
'd2infile = "auftritt": d2insub = "findfeld_LostFocus"
findfeld.Width = 390

End Sub

Private Sub Form_Load()
Dim rtmp As ADODB.Recordset, klrv%, s%, dbpara$, i%, ast1$, rrr
Dim colHeader

Dim ctrl As Control
Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Form_Load"

listMessages.View = lvwReport
Set colHeader = listMessages.ColumnHeaders.add(, , transe("Datum"), 1400)
Set colHeader = listMessages.ColumnHeaders.add(, , transe("Check"), 3500)
Set colHeader = listMessages.ColumnHeaders.add(, , transe("Bestätigt"), 2500)
Set colHeader = listMessages.ColumnHeaders.add(, , transe("Benutzer"), 1000)
Set colHeader = listMessages.ColumnHeaders.add(, , "Message-ID", 2)
chkokall.Caption = transe("Alle ok")
chkdlall.Caption = transe("Alle löschen")
chkudall.Caption = transe("Alle wiederherstellen")
chkedone.Caption = transe("Neuer Checkpoint")
chkown.Caption = transe("Besitzer")
chkedt.ToolTipText = transe("gewählte Nachricht bearbeiten")
Command42.ToolTipText = transe("Adressen im Umkreis um eine Postleitzahl suchen")
ttlptr% = -1
form1.fastsave_copy = False
Call form1.dbg2f("auftritt:load")
dbpara$ = form1.getconnstr()
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)

nbase% = 0
nflds = 7
fpp% = 33
s% = form1.myfontsize()
listMessages.Font.Size = s%
For i% = 0 To nflds: Text1(i%).Font.Size = s%: Next i%
For i% = 0 To fpp%: Text2(i%).Font.Size = s%: Next i%
fl_showrec% = 0
i = 0
Do
  ast1$ = form1.get_eventstatusname(i%)
  If ast1$ <> "" Then astatcmb.AddItem ast1$
  i% = i% + 1
Loop Until ast1$ = transe("kein Status")

If form1.immerkalender() = "ja" Then
  kalimmer.value = 1
  klrv% = Val(form1.mylastFormVar(Me.name, "kalres", "0"))
  If klrv% <> 0 Then klrv% = 1
  kalres.value = klrv%
Else
  kalres.value = 0
  kalres.Enabled = False
End If
delmode = False
Combo2.text = form1.getusersetting("TerminDokumentname", "Typ-Termin-User")
Shape1.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Label5.Caption = transe("% MwSt")
Command6.Caption = transe("Drucken")
Command7.Caption = transe("kopieren")
Command8.Caption = transe("ausfüllen")
Command9.Caption = transe("Alarme")
Command14.Caption = transe("wiederholen")
Command17.ToolTipText = transe("Neue Vorlage für diesen Termintyp erstellen")
Command19.ToolTipText = transe("Tageskalender öffnen")
Command16.ToolTipText = transe("Kalender öffnen")
vwopts.ToolTipText = transe("Kalenderanzeigeoptionen")
Command21(0).Visible = False
Command21(1).Visible = True
Load ttform
Show
form1.auftrittisopen = True
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM auftrittstypen order by sortierung", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Call nulldsp

While Not rtmp.EOF
  Combo1.AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend
Combo1.text = ""
Command42.Enabled = False
If form1.geodbok Then Command42.Enabled = True

BackColor = form1.cleancolor()
kalkdirty = False
Command10.Enabled = False
Command13.Enabled = True
If nexist(form1.vorlagenverzeichnis() + "\v0__auftritt.rtf") Then Command17.Enabled = False

End Sub

Sub nulldsp()
Dim i%

'd2infile = "auftritt": d2insub = "nulldsp"
delmode = False
Command42.BackColor = &HC0C0C0
For i% = 0 To 99: chgsx(i%) = "": Next i%
Unload termviewlist
listMessages.Visible = False
btnTopic.Enabled = False
For i% = 0 To fpp%
  clickgetsfromtable(i%) = ""
  clickgetsfromfield(i%) = ""
  Label2(i%).Visible = False
  Text2(i%).Visible = False
  Command3(i%).Visible = False
Next i%

On Error Resume Next
For i% = 0 To nflds
  Label1(i%).Caption = transe(form1.sqla.TableDefs("auftritt").Fields(i%).name)
  Text1(i%).text = ""
Next i%
On Error GoTo 0
For i% = 0 To fpp%
  Label2(i%).ForeColor = &H80000012
Next i%

End Sub
Sub showrec(id$, initmode%)
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, pr As ADODB.Recordset, stmp As ADODB.Recordset, rr As ADODB.Recordset
Dim astatus As ComboBox, rrr, cmd$, tr As String, pnm$, anm$, cktst As Boolean, tpid$, pos As Integer
Dim c$, f$

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "showrec"
If fl_showrec% = 1 Then Exit Sub
Unload termviewlist

recalclist = ""
recalcplease = False
fl_showrec% = 1
On Error Resume Next
Unload tlform
On Error GoTo 0
Call savecheck
Call nulldsp
List1.Clear
delmode = False
Text1(0).text = id$
prio.text = ""
astatcmb.Visible = False
cmd$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and auftrittstyp='kalku_Künstlerauftritt'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("auftritt.showrec:" & cmd$)
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  While Not r.EOF
    c$ = trm(r!felddaten): pos = 1
    While pos > 0
      pos = InStr(c$, "auftritthigru__")
      If pos > 0 Then
        c$ = Mid$(c$, pos)
        c$ = cut_d2bis(cut_d2bis(c$, "_"), "_")
        f$ = LCase(cut_d1(c$, Chr$(13)))
        If InStr(recalclist, "|" + f$ + "|") = 0 Then recalclist = recalclist + "|" + f$ + "|"
      End If
    Wend
    r.MoveNext
  Wend
  recalclist = strrepl(recalclist, "||", "|")
End If
cmd$ = "SELECT * FROM auftritt where id= '" & id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
Call form1.dbg2f("auftritt.showrec:" & cmd$)
On Error Resume Next
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
If Not r.EOF Then
  If Not form1.isfieldmissing("opt_prios", "id") Then
    prio.Enabled = True
    prio.Visible = True
    cmd$ = "SELECT * FROM opt_prios where evnt= 'E:" & id$ & "' and userid='" + form1.getuserid() + "'"
    Set pr = New ADODB.Recordset
    pr.CursorLocation = adUseServer
    Call form1.dbg2f("auftritt.showrec:" & cmd$)
    On Error Resume Next
    rrr = form1.adoopen(pr, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then Exit Sub
    If Not pr.EOF Then prio.text = trmx1(pr!prio)
  Else
    prio.Enabled = False
    prio.Visible = False
  End If
  On Error Resume Next
  ast = r!astatus: setast = -1
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    pstt.BackColor = form1.cleancolor()
    pstt.Cls
    pstt.Print "kein Status"
  Else
    pstt.BackColor = form1.get_eventstatuscolor(ast)
    pstt.Cls
    pstt.Print form1.get_eventstatusname(ast)
  End If
  calcol.BackColor = 0
  If Not form1.isfieldmissing("auftritt", "optkalcolor") Then
    calcol.BackColor = Val(trm0(r!optkalcolor))
  End If
  If trm(r!auftrittstyp) = "" Then Exit Sub
  If calcol.BackColor <= 0 Then
    calcol.BackColor = form1.get_eventcolor(r!auftrittstyp)
  End If
  For i% = 1 To nflds
    If Not IsNull(r.Fields(i%)) Then
      If i% = 2 Then
        Text1(i%).text = datfromsql(trm(r.Fields(i%)))
      Else
        Text1(i%).text = trm(r.Fields(i%).value)
        If i% = 3 Then
          'Text1(3).ToolTipText = transe("Ende") + ":" + form1.auftrittsende(id$, "")
          Text3.text = form1.auftrittsende(id$, "")
          Text3.ToolTipText = transe("Ende")
'          Label1(3).ToolTipText = transe("Doppelklick um die Endezeit zu setzen")
        End If
      End If
      If i% = 5 Then
        Text1(i%).text = transe(trm(Text1(i%).text))
        Combo1.text = Text1(i%).text
      End If
      If i% = 1 Or i = 3 Then Label1(i%).ForeColor = form1.lnkcolor
    End If
  Next i%
End If

On Error Resume Next
p$ = transo(Combo1.text)
If p$ <> "Neuer Auftritt" Then
  Combo1.Visible = False
  Text1(5).Enabled = False
  Command9.Visible = True
  mwst.Visible = True
  Label5.Visible = True
Else
  mwst.Visible = False
  Label5.Visible = False
  Combo1.Visible = True
  Text1(5).Enabled = True
  Command9.Visible = False
End If
auftritt.Caption = transe(p$)
On Error Resume Next
tr$ = Dir(form1.vorlagenverzeichnis() + "\" & trm(r!auftrittstyp) & "*.rtf")
rrr = Err
On Error GoTo 0
While tr$ <> "" And rrr = 0
  If trm(r!auftrittstyp) = Left$(tr$, InStr(tr$, ".") - 1) Then
    List1.AddItem basename(r!auftrittstyp, ".rtf")
  Else
    ffn$ = basename(Mid$(tr$, InStr(tr$, "_") + 1), ".rtf")
    If Left(ffn$, 4) <> "zsys" Then List1.AddItem ffn$
  End If
  tr$ = Dir
Wend
On Error Resume Next
tr$ = Dir(form1.vorlagenverzeichnis() + "\" & trm(r!auftrittstyp) + "*.alias")
rrr = Err
On Error GoTo 0
While tr$ <> "" And rrr = 0
  ffn$ = basename(Mid$(tr$, InStr(tr$, "_") + 1), ".alias")
  If Left(ffn$, 4) <> "zsys" Then List1.AddItem ffn$
  tr$ = Dir
Wend
If Text1(0).text = "" Then Text1(0).text = id$
If initmode% = 0 And trm(r!tstamp) = "" Then initmode% = 1
Call initfields("" & p$ & "", 0, initmode%)

c$ = "select * from usr_" & utabn(transe(p$)) & " where id='" & id$ & "'"
Set rr = New ADODB.Recordset
rr.CursorLocation = adUseServer
rrr = form1.adoopen(rr, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If rr.EOF Then Call Command10_Click
Else
  Call Command10_Click
End If
vwopts.ToolTipText = "sichtbar für: " + form1.terminvizlist(id$)
vwopts.ToolTipText = vwopts.ToolTipText + " unsichtbar für: " + form1.termininvizlist(id$)

If form1.immerkalender() = "ja" Then
  dtg$ = Text1(2).text
  If dtg$ <> "" Then
    d0 = Time
    Load kc
'select: nur Kalendereinträge für vorkommende
    If krcount% > 0 And kalres.value <> 0 Then
      kc.selct(2).Clear
      For i% = 0 To krcount% - 1
        kc.selct(2).AddItem kalrestrict$(i%)
      Next i%
    End If
    Call kc.settag0(dtg$)
    If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then kc.Text1.text = 1
  End If
End If
fl_showrec% = 0
kalkdirty = False
tpid$ = Text1(1).text
pnm$ = form1.medienname(Text1(1).text)
anm$ = form1.medienname(form1.get_atabkz(trm(transe(transo(Text1(5).text)) & "_" & Text1(0).text)))
On Error Resume Next
tr = Dir(form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + pnm$ & "\" & anm$ + "\*.*")
rrr = Err
On Error GoTo 0
If tr$ <> "" And rrr = 0 Then opendir.Picture = Picture3(1).Picture
Command10.Enabled = False
Command13.Enabled = True
If form1.getusersetting("TerminDatentest") = "ja" Then Call dtst_Click

If Not form1.isfieldmissing("opt_topics", "id") Then
  c$ = "select * from opt_topics where topicid='" + tpid$ + "'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    If Not rtmp.EOF Then btnTopic.Enabled = True
  End If
End If
On Error Resume Next
Call Text1(6).SetFocus
On Error GoTo 0
Call Text1_GotFocus(6)
Call achktst(id$)

c$ = "SELECT * FROM auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='umkrchk' and auftrittstyp='" & transo(Text1(5).text) & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If Not r.EOF Then
    Call umkrtest(Text1(0).text, trm(r!felddaten))
  Else
    Command42.BackColor = &HC0C0C0
  End If
End If
r.Close

BackColor = form1.cleancolor()

fromtpwernoch.Clear
tpid$ = Text1(1).text
If tpid$ <> "" And tpid$ <> "-1" Then
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
  rrr = form1.adoopen(stmp, "SELECT funktion,kid FROM tpwernoch where tpid='" & tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    While Not stmp.EOF
      fromtpwernoch.AddItem formtranse(transe(trm(stmp!funktion))) + "|" + trm(stmp!kid)
      stmp.MoveNext
    Wend
  End If
  For i% = 0 To 33
    If Label2(i%).Visible And Text2(i%).text = "" Then
      For fri% = 0 To fromtpwernoch.ListCount - 1
        c$ = cut_d1(fromtpwernoch.List(fri%), "|")
        If c$ = Label2(i%).Caption Then
          wert$ = cut_d2bis(fromtpwernoch.List(fri%), "|")
          If wert$ <> "" Then
            Call Text2_GotFocus(i%)
            Text2(i%).text = wert$
            Call Text2_LostFocus(i%)
          End If
          Exit For
        End If
      Next fri%
    End If
  Next i%
End If

End Sub

Private Sub achktst(aid$)
If form1.check_tst(aid$) Then
  Command21(0).Visible = False
  Command21(1).Visible = True
Else
  Command21(1).Visible = False
  Command21(0).Visible = True
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ttform.Hide
tlform.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "auftritt": d2insub = "Form_Unload"
form1.fastsave_copy = False
Unload fdet
Unload kalku
Unload termviewlist
Unload auftrittslisten
Unload ttform
Unload tlform
Call savecheck
Hide
form1.auftrittisopen = False
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub

Private Sub gd1_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)

'd2infile = "auftritt": d2insub = "gd1_AfterLabelEdit"
i = Index
gd1(i).SelectedItem.text = NewString
Call gd1after(Index, "numsort")
Call makemedirty
End Sub

Private Sub gd1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim neuid As String, snum As String, lblnum As Integer, i As Integer, ort As String

'd2infile = "auftritt": d2insub = "gd1_ColumnClick"
snum = trm(ColumnHeader)
lblnum = labelnumbylistnum(Index)
If snum = "?" Then
  On Error Resume Next
  Unload auftrittslisten
  DoEvents
  Load auftrittslisten
  ort = Text1(7).text
  If ort <> "" Then ort = " " + ort
  auftrittslisten.Label2(0).Caption = form1.get_atabkz(transo(Text1(5).text)) + ort + " " + Text1(2).text + " " + Text1(3).text
  auftrittslisten.Label1.Caption = Text1(0).text
  auftrittslisten.typ.Caption = transo(Text1(5).text)
  On Error GoTo 0
  DoEvents
  auftrittslisten.List3(0).AddItem formtranso(Label2(lblnum).Caption)
  For i = 0 To auftrittslisten.List3(1).ListCount - 1
    If auftrittslisten.List3(1).List(i) = formtranso(Label2(lblnum).Caption) Then
      auftrittslisten.List3(1).RemoveItem i
    End If
  Next i
  Exit Sub
End If
If snum = "Name" Then
  Call gd1after(Index, "namsort")
  Call makemedirty
  Exit Sub
End If
If Not isnumber(snum) Then Exit Sub
neuid = InputBox(transe("Neue Spaltenüberschrift"), transe("Neue Spaltenüberschrift"))
If neuid <> "" Then
  Call form1.setsystemsetting(transo(Text1(5).text) + "_" + formtranso(Label2(lblnum).Caption) + "_" + snum, neuid)
  gd1(Index).ColumnHeaders(Val(snum) + 1).text = neuid
End If

End Sub

Private Sub gd1_DblClick(Index As Integer)
Dim id$, sid$, p%, sida$, sidk$

'd2infile = "auftritt": d2insub = "gd1_DblClick"
If gd1(Index).ListItems.Count <= 0 Then Exit Sub
id$ = gd1(Index).SelectedItem.SubItems(1)
sid$ = id$
p% = InStr(sid$, "{")
sida$ = sid$: sidk$ = ""
If p% > 0 Then
  sidk$ = trm(Left(sid$, p% - 1))
  sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
End If
If Len(sida$) > 0 Then
  Load shwAdrDetail
  Call shwAdrDetail.refreshadrdetail(sida$, sidk$)
  On Error Resume Next
  Call shwAdrDetail.SetFocus
  On Error GoTo 0
End If

End Sub

Private Sub gd1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
'd2infile = "auftritt": d2insub = "gd1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
    gd1(Index).SelectedItem.text = "delete"
    Call gd1after(Index, "numsort")
    Call makemedirty
End If
End Sub

Private Sub gd1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'd2infile = "auftritt": d2insub = "gd1_MouseMove"
gdmx = X
End Sub

Private Sub gd1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim gds As Integer, xp As Integer, hdr As String, wert As String

'd2infile = "auftritt": d2insub = "gd1_MouseUp"
If Button = 2 Then
  xp = gdmx
  gds = 1
  Do
    gds = gds + 1
    xp = xp - gd1(Index).ColumnHeaders(gds - 1).Width
  Loop Until xp < 0
  gds = gds - 1
  hdr = gd1(Index).ColumnHeaders(gds)
  If gds > 2 Then
    wert = gd1(Index).SelectedItem.SubItems(gds - 1)
    wert = InputBox(transe("Neuer Wert") + vbCrLf + transe("(Leerzeichen zum Löschen)"), transe("Neuer Wert von") + " " + hdr, wert)
    If trm(wert) = "" And wert <> " " Then Exit Sub
    wert = trm(wert)
    gd1(Index).SelectedItem.SubItems(gds - 1) = wert
    Call makemedirty
  End If
End If

End Sub

Private Sub kalimmer_Click()

'd2infile = "auftritt": d2insub = "kalimmer_Click"
If kalimmer.value = 0 Then
  Call form1.setbenutzerdaten("immer_kalender", "nein")
  kalres.value = 0
  kalres.Enabled = False
Else
  Call form1.setbenutzerdaten("immer_kalender", "ja")
  kalres.Enabled = True
  klrv% = Val(form1.mylastFormVar(Me.name, "kalres", "0"))
  If klrv% <> 0 Then klrv% = 1
  kalres.value = klrv%
End If

End Sub

Private Sub kalres_Click()
'd2infile = "auftritt": d2insub = "kalres_Click"
Call form1.setmylastFormVar(Me.name, "kalres", trm(kalres.value))

End Sub

Private Sub Label1_DblClick(Index As Integer)
Dim c$, aid$, typ$, s$, altwert$

'd2infile = "auftritt": d2insub = "Label1_DblClick"
If Index = 1 Then
  tpid$ = Text1(1).text
  If Len(tpid$) <> 0 Then
    Load tplan
    Call tplan.rlists
    Call tplan.nulldsp
    Call tplan.showrec(tpid$)
    On Error Resume Next
    Call tplan.SetFocus
    On Error GoTo 0
  End If
End If
If Index = 5 Then
  Call savecheck
  Combo1.Visible = True
  Text1(5).Enabled = True
  Command9.Visible = False
  mwst.Visible = False
  Label5.Visible = True
End If
End Sub

Private Sub Label2_DblClick(Index As Integer)
Dim i%, neuwert, tpid$, matchlen%, neukwert, fn$, p%, X, cgft$, lcounter As Integer
Dim r As ADODB.Recordset, rtmp As ADODB.Recordset, cmd$, opn As Boolean, listno As Integer
Dim s1 As ADODB.Recordset, p1%, preis$, preisfeld$, neudraw As Boolean, l2c$, gn$
Dim lvitem, ort As String, abvno As Integer, rrr, currl As String, nid$
Dim neuid As String

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "Label2_DblClick"
tpid$ = trm(Text1(1).text)
i% = Index
Call form1.dbg2f("auftritt:lbl2_dblclick:" & i%)
If Text2(Index).Enabled = False Then
  Text2(Index).Enabled = True
  Call Text2(Index).SetFocus
  Text2(Index).text = Text2(Index).text
Else
  If Text2(Index).Visible Then Call Text2(Index).SetFocus
End If
DoEvents
'do NOT Call savecheck
neudraw = False
s$ = clickgetsfromfield(i%)
opn = False
cgft$ = clickgetsfromtable(i%)
If LCase(cgft$) = "vertragsnummer" Then
  If form1.isfieldmissing("opt_vnr", "id") Then
    MsgBox ("Tabelle opt_vnr fehlt.")
  Else
    If trm(Text2(Index).text) <> "" Then
      MsgBox ("Das Feld enthält bereits Daten.")
      Exit Sub
    End If
    neuwert = trm(form1.neuevertragsnummer())
    cmd$ = "update opt_vnr set aid='" + Text1(0).text + "' where id='" + trm(neuwert) + "'"
    Call form1.sqlqry(cmd$)
    i% = Index
    Call Text2_GotFocus(i%)
    Text2(i%).text = neuwert
    DoEvents
    Call Text2_LostFocus(i%)
    Call Command10_Click
    Exit Sub
  End If
  opn = True
End If
If LCase(cgft$) = "besetzung" Then
  Call savecheck
  Unload besetzung
  DoEvents
  Load besetzung
  On Error Resume Next
  Call besetzung.SetFocus
  On Error GoTo 0
  besetzung.werkid = "T:" + Text1(0).text
  DoEvents
  ort = Text1(7).text
  If ort <> "" Then ort = " " + ort
  besetzung.werknam.Caption = formtranso(Label2(i%).Caption) + " " + form1.get_atabkz(transo(Text1(5).text)) + ort + " " + Text1(2).text + " " + Text1(3).text
  opn = True
End If
If LCase(cgft$) = "tabelle" And Not opn Then
  Call savecheck
  Unload tabkalk
  DoEvents
  Load tabkalk
  On Error Resume Next
  Call tabkalk.SetFocus
  On Error GoTo 0
  tabkalk.Label2.Caption = Text1(0).text
  tabkalk.Label3.Caption = formtranso(Label2(i%).Caption)
  tabkalk.Label1.Caption = transo(Text1(5).text)
  opn = True
End If
If Not opn Then
l2c$ = LCase(transo(formtranso(Label2(i%).Caption)))
If InStr(l2c$, "honorar") = 1 Or _
    InStr(l2c$, "betrag") > 0 Or _
    InStr(l2c$, "preis") > 0 Or _
    InStr(l2c$, "konzertauslastung") > 0 Or _
    LCase(cgft$) = "finanzen" Then
  Call savecheck
  Unload kalku
  DoEvents
  Load kalku
  kalku.afeld = formtranso(Label2(i%).Caption)
  kalku.kerg.Caption = Text2(i%).text
  kalku.atyp = transo(Text1(5).text)
  kalku.aid = Text1(0).text
  DoEvents
  On Error Resume Next
  Call kalku.SetFocus
  On Error GoTo 0
  Exit Sub
End If
End If
If cgft$ = "" Or s$ = "" Then Exit Sub
'MsgBox clickgetsfromtable(i%) & " " & s$

neuwert = ""
If LCase(cgft$) = "programm" Then
  s0$ = trm(Text2(Index).text)
  If LCase(Left$(s0$, 6)) = "datei:" Or LCase(Left$(s0$, 3)) = "fn:" Then
    p% = InStr(s0$, ":")
    fn$ = ""
    If p% > 0 Then
      fn$ = trm(Mid$(s0$, p% + 1))
      If fn$ = "" Or nexist(fn$) Then
        fn$ = form1.myuniquedocname("noask") & ".txt"
'        fn$ = form1.filenamekurz(form1.saveasBox(fn$))
        fn$ = form1.saveasBox(fn$)
      End If
      If fn$ <> "" Then
        If InStr(fn$, "(") = 1 Then fn$ = cut_d2bis(fn$, ")")
        If Not nexist(fn$) Then
          If InStr(LCase(fn$), ".doc") > 0 Or InStr(LCase(fn$), ".rtf") > 0 Then
            Call form1.openthisdoc(fn$, "")
          Else
            X = Shell("notepad.exe " & fn$, 1)
          End If
          neuwert = "fn:" & fn$
        Else
          MsgBox (transe("Datei nicht gefunden"))
        End If
      End If
    End If
  Else
    If CtrlKey() Then
      nid$ = trm(onlyalpha(Text1(6).text) + " " + onlyalpha(Text1(7)) + " " + datum2sql(Text1(2).text) + " " + Text1(0).text)
      If Len(nid$) > 100 Then nid = Left$(nid$, 100)
      neuid = InputBox(transe("Neues Programm"), "", nid$)
      If trm(neuid) = "" Then Exit Sub
      Call form1.sqlqry("INSERT INTO programm (programmID) VALUES('" & neuid & "')")
      Load prog
      Call prog.SetFocus
      Call prog.selectone(neuid)
      Call Text2_GotFocus(Index)
      Text2(Index).text = neuid
      neuwert = neuid
    Else
      Load prog
      Call prog.SetFocus
      Call prog.callbackinit("auftritt", tpid$)
      neuprogid$ = ""
      Call prog.rlist1
      If s0$ <> "" Then
        Call prog.selectone(s0$)
      End If
      While neuprogid$ = "": DoEvents: Wend
      If neuprogid$ <> "" And neuprogid$ <> "_LOGOUT_" Then neuwert = neuprogid$
      Unload prog
    End If
  End If
End If
If LCase(cgft$) = "tabelle" Then
  neuwert = trm(Text2(Index).text)
End If
If LCase(cgft$) = "adrselect" Then
  Load adrselect
  s0$ = Text2(Index).text
  If Len(s0$) = 0 Then s0$ = ""
  matchlen% = Val(form1.getusersetting("AdressgruppenMatch", "0"))
  If matchlen% > 0 Then
    If Len(s$) > 4 Then s$ = Left(s$, 4) & "*"
    s$ = LCase(s$)
  End If
  Call adrselect.sel_init(s0$, transe(s$))
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then

  neukwert = adrselect.get_kontsel()
  neuwert = adrselect.sel_getselected(): neuawert = neuwert
  If Text2(i%).Visible Then
    If InStr(LCase$(transo(formtranso(Label2(i% + 1).Caption))), "kontakt") = 1 Then
      If neukwert <> "" Then
        Call Text2_GotFocus(i% + 1)
        Text2(i% + 1).text = neukwert
        Call Text2_LostFocus(i% + 1)
      End If
    Else
      If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
    End If
  End If

  End If
  Unload adrselect
  If transo(formtranso(Label2(Index).Caption)) = "Halle" Or transo(formtranso(Label2(Index).Caption)) = "Saal" Then
    If Len(Text1(7).text) = 0 Then
      o$ = ohnePLZ(form1.ortausadr("" & neuawert & ""))
      If o$ <> "" Then
        Call Text1_GotFocus(7)
        Text1(7).text = o$
        Call Text1_LostFocus(7)
      End If
    End If
    If transo(Text1(5).text) = "Veranstaltung" And form1.isoftype(neuwert, "Bestuhlung") <> "-1" Then
      MousePointer = 11: DoEvents
      'gleichnamige felder finden
      cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='Bestuhlung'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r.EOF
        Set rtmp = New ADODB.Recordset
        rtmp.CursorLocation = adUseServer
        cmd$ = "SELECT feldname FROM auftrittsfelder where typ='" & transo(Text1(5).text) & "' and (feldname='" & r!feldname & "' or instr(feldname,'." & r!feldname & ".')>0 )"
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not rtmp.EOF Then
          ' eintrag im termin ermitteln
          cmd$ = "select id,felddaten from auftritthigru where auftrittsid='" & Text1(0).text & "' and auftrittstyp='" & transo(Text1(5).text) & "' and feldname='" & r!feldname & "'"
          Set s1 = New ADODB.Recordset
          s1.CursorLocation = adUseServer
rrr = form1.adoopen(s1, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If s1.EOF Then
            cmd$ = ""
          Else
            cmd$ = trm(s1!felddaten)
            If cmd$ = "" Then
              'datensatz da, aber leer -->löschen
              cmd$ = "delete from auftritthigru where id='" & s1!id & "'"
              Call form1.sqlqry(cmd$)
              cmd$ = ""
            End If
          End If
          If cmd$ = "" Then
            ' eintrag im termin ist leer, daten suchen
            cmd$ = "select felddaten from auftritthigru where auftrittsid='" & form1.AdrIDErmittlung(neuwert) & "' and auftrittstyp='Bestuhlung' and feldname='" & r!feldname & "'"
            Set s1 = New ADODB.Recordset
            s1.CursorLocation = adUseServer
rrr = form1.adoopen(s1, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not s1.EOF Then
              If trm(s1!felddaten) <> "" Then
                If InStr(LCase(r!feldname), "name") > 0 Then
                  neudraw = True
                  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
                        form1.newid("auftritthigru", "id", 50) & "','" & _
                        Text1(0).text & "','" & _
                        "Veranstaltung" & "','" & _
                        r!feldname & "','" & _
                        s1!felddaten & "')"
                  Call form1.sqlqry(cmd$)
                  cmd$ = "update usr_veranstaltung set " & r!feldname & "='" & s1!felddaten & "' where id='" & Text1(0).text & "'"
                  Call form1.sqlqry(cmd$)
                End If
                p% = InStr(LCase(r!feldname), "sitze")
                If p% > 0 Then
                  neudraw = True
                  preisfeld$ = Left$(r!feldname, p% - 1) & "Preis"
                  cmd$ = trm(s1!felddaten)
                  p% = InStr(cmd$, "/")
                  If p% > 1 Then
                    o$ = trm(Left$(cmd$, p% - 1))
                    preis$ = trm(Mid$(cmd$, p% + 1))
                  Else
                    o$ = cmd$
                    preis$ = ""
                  End If
                  cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
                        form1.newid("auftritthigru", "id", 50) & "','" & _
                        Text1(0).text & "','" & _
                        "Veranstaltung" & "','" & _
                        r!feldname & "','" & _
                        o$ & "')"
                  Call form1.sqlqry(cmd$)
                  cmd$ = "update usr_veranstaltung set " & r!feldname & "='" & o$ & "' where id='" & Text1(0).text & "'"
                  Call form1.sqlqry(cmd$)
                  If preis$ <> "" Then
                    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" & _
                        form1.newid("auftritthigru", "id", 50) & "','" & _
                        Text1(0).text & "','" & _
                        "Veranstaltung" & "','" & _
                        preisfeld$ & "','" & _
                        preis$ & "')"
                    Call form1.sqlqry(cmd$)
                    cmd$ = "update usr_veranstaltung set " & preisfeld$ & "='" & preis$ & "' where id='" & Text1(0).text & "'"
                    Call form1.sqlqry(cmd$)
                  End If
                End If
              End If
            End If
          End If
        End If
        r.MoveNext
      Wend
      MousePointer = 0
    End If
  End If
End If
'If InStr(Text2(Index).Height, Chr$(13) + Chr$(10)) > 0 Then

If Text2(Index).Visible Then
  If InStr(Text2(Index).text, Chr$(13) + Chr$(10)) > 0 Then
    If Len(Text2(Index).text) > 0 Then
      Text2(Index).text = Text2(Index).text + Chr$(13) + Chr$(10) + neuwert
    Else
      Text2(Index).text = neuwert
    End If
  Else
    Text2(Index).text = neuwert
  End If
  Call Text2_LostFocus(Index)
Else
  'Eintrag in Liste
  If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
  If neuwert <> "" Then
    If InStr(neuwert, "GRUPPEGEWÄHLT:") = 1 Then
      gn$ = cut_d2bis(trm(neuwert), ":")
      neuwert = ""
      Set rtmp = New ADODB.Recordset
      rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppen where grpid='" + gn$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not rtmp.EOF
        If neuwert <> "" Then neuwert = neuwert + vbCrLf
        If Not IsNull(rtmp!kid) And rtmp!kid <> "-1" Then
          neuwert = neuwert + form1.get_kontaktname_by_id(rtmp!kid) + " {" + rtmp!adressid + "}"
        Else
          neuwert = neuwert + rtmp!adressid
        End If
        rtmp.MoveNext
      Wend
      rtmp.Close
    End If
    listno = listnnumbytextpos(Index)
    If listno > 0 Then
      Call makemedirty
      On Error Resume Next
      abvno = Val(gd1(listno).SelectedItem.text) - 1
      rrr = Err
      On Error GoTo 0
      If rrr Or abvno < 0 Then abvno = 0
      For lcounter = 1 To linesof(trm(neuwert))
        currl = lineof(lcounter, trm(neuwert))
        Set lvitem = gd1(listno).ListItems.add(, , trm(abvno) + "-" + trm(lcounter))
        lvitem.SubItems(1) = currl
        p% = 3
        While p% <= gd1(listno).ColumnHeaders.Count
          If Not isnumber(gd1(listno).ColumnHeaders(p%)) Then
            c$ = "select * from auftritthigru where auftrittsid='" + neuwert + "' and auftrittstyp='" + s$ + "' and feldname='" + gd1(listno).ColumnHeaders(p%) + "';"
            Set r = New ADODB.Recordset
            r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not r.EOF Then
              lvitem.SubItems(p% - 1) = trm(r!felddaten)
            End If
          End If
          p% = p% + 1
        Wend
      Next lcounter
    End If
    Call gd1after(listno, "numsort")
  End If
End If
MousePointer = 0
If neudraw Then Call showrec(Text1(0).text, 0)

End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And (transo(Text1(6).text) = "Neuer Auftritt" Or Text1(6).text = "" Or Text1(6).text = Text1(1).text) Then
  Call Text1_GotFocus(6)
  DoEvents
  Text1(6).text = Text2(Index).text
  Call Text1_LostFocus(6)
End If
End Sub

Private Sub Label6_Click()
'd2infile = "auftritt": d2insub = "Label6_Click"
If kalres.value = 0 Then
  kalres.value = 1
Else
  kalres.value = 0
End If
End Sub

Private Sub Label7_Click()
'd2infile = "auftritt": d2insub = "Label7_Click"
If kalimmer.value = 0 Then
  kalimmer.value = 1
Else
  kalimmer.value = 0
End If

End Sub

Public Sub List1_DblClick()

'd2infile = "auftritt": d2insub = "List1_DblClick"
Call Command6_Click

End Sub

Private Sub listMessages_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub listMessages_Click()
Dim i%, viz As Boolean

viz = False
uselct.Visible = False
For i% = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i%).Selected = True Then Exit For
Next i%
If i% <= listMessages.ListItems.Count Then viz = True
chkown.Visible = viz

End Sub

Private Sub listMessages_DblClick()
Dim frm$, p%, rrr, i%, o%, l$, sbf$, sbj$, trg$, msgid$
Dim lvitem, hdm As Boolean, strMessageHeader As String
Dim r As ADODB.Recordset, c$, n%
Dim id$, pos%, cnf$, cf$

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "listMessages_DblClick"
On Error Resume Next
frm$ = listMessages.SelectedItem
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

p% = 0: n% = 0
For i% = 1 To listMessages.ListItems.Count
  If listMessages.ListItems(i%).Selected = True Then Exit For
Next i%
If i% <= listMessages.ListItems.Count Then

id$ = listMessages.ListItems(i%).SubItems(4)
cnf$ = listMessages.ListItems(i%).SubItems(2)
cf$ = listMessages.ListItems(i%).SubItems(1)
wert$ = InputBox("Confirm: " + cf$ + vbCrLf + "(0=k=ok)", "Reminder", cnf$)
If wert$ = "k" Or wert$ = "ok" Or wert$ = "0" Then
  wert$ = "ok, " + trm(Date) + " " + trm(Time) + " " + form1.getuserid()
End If
If wert$ <> cnf$ And wert <> "" Then
    cf$ = "update opt_checks set confirmed='" + wert$ + "' where id='" + id$ + "'"
    Call form1.sqlqry(cf$)
    listMessages.Visible = False
    Call Command21_Click(1)
End If
id$ = Text1(0).text
If id$ <> "" Then Call achktst(id$)

End If
End Sub

Private Sub listMessages_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, wert$, id$, cnf$

If KeyCode = 17 Or KeyCode = 17 Then Exit Sub
'Debug.Print KeyCode, Shift
If KeyCode = 8 Or KeyCode = 46 Then
  For i% = 1 To listMessages.ListItems.Count
    If listMessages.ListItems(i%).Selected And trm(listMessages.ListItems(i%).SubItems(2)) = "" Then
      id$ = listMessages.ListItems(i%).SubItems(4)
      wert$ = "ok, deleted " + trm(Date) + " " + trm(Time) + " " + form1.getuserid()
      listMessages.ListItems(i%).SubItems(2) = wert$
      DoEvents
      cf$ = "update opt_checks set confirmed='" + wert$ + "' where id='" + id$ + "'"
      Call form1.sqlqry(cf$)
    End If
  Next i%
  listMessages.Visible = False
  Call Command21_Click(1)
  id$ = Text1(0).text
  If id$ <> "" Then Call achktst(id$)
End If
End Sub

Private Sub mwst_Change()

'd2infile = "auftritt": d2insub = "mwst_Change"
  kalkdirty = True
  Call makemedirty
End Sub

Private Function auftrittsverzeichnis() As String
Dim pnm$, anm$, rc$

pnm$ = form1.medienname(Text1(1).text)
anm$ = form1.medienname(form1.get_atabkz(trm(transo(Text1(5).text) & "_" & Text1(0).text)))
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\"
MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + pnm$
rc$ = form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + pnm$ & "\" & anm$
MkDir rc$
auftrittsverzeichnis = rc$
End Function

Private Sub opendir_Click()

X = Shell("explorer.exe " + auftrittsverzeichnis(), vbNormalFocus)
On Error GoTo 0

End Sub

Private Sub prio_Change()
Dim c As String, id, p As String, nid As String

'd2infile = "auftritt": d2insub = "prio_Change"
p = UCase(prio.text)
If p < "A" And p <> "" Then p = "A"
If p > "Z" And p <> "" Then p = "Z"
prio.text = p
id = trmx1(Text1(0).text)
If id <> "" Then
  c = "delete from opt_prios where userid='" + form1.getuserid() + "' and evnt='E:" + id + "';"
  Call form1.sqlqry(c)
  If p <> "" Then
    nid = form1.newid("opt_prios", "id", 36)
    c = "insert into opt_prios (id,evnt,userid,prio) values('" + _
        nid + "','E:" + _
        id + "','" + _
        form1.getuserid() + "','" + _
         p + "');"
    Call form1.sqlqry(c)
  End If
  If form1.priosopen Then Call prios.Command20_Click
End If

End Sub

Private Sub pstt_Click()
'd2infile = "auftritt": d2insub = "pstt_Click"
astatcmb.Visible = True
pstt.Visible = False

End Sub

Private Sub Text1_Change(Index As Integer)
Dim o%, l$, d$, rrr

Call makemedirty
If Index = 5 Then
  formttp = -1
  If trm(Text1(Index).text) = "" Then Exit Sub
  l$ = form1.vorlagendir + "\" + Text1(Index).text + ".frm.txt"
  If Not nexist(l$) Then
    o% = FreeFile
    Open l$ For Input As #o%
    While Not EOF(o%)
      formttp = formttp + 1
      Line Input #o%, l$
      formtt(formttp) = l$
    Wend
    Close #o%
  End If
End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
Dim id$

'd2infile = "auftritt": d2insub = "Text1_DblClick"
id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  Exit Sub
End If
If Index = 2 Then
  With frmCalendar
    .init Text1(2), Text1(2).text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1(2).text = Format(.SelectedDate, "dd.mm.yyyy")
      Call Text1_LostFocus(Index)
    End If
  End With
  Unload frmCalendar
End If

End Sub

Public Sub Text1_GotFocus(Index As Integer)
prv$ = Text1(Index).text

End Sub

Public Sub Text1_LostFocus(Index As Integer)
Dim s$, typ$, c$, dtg, an$, rrr, ask%
Dim betreff$, nachricht$, trgdatum$, trgzeit$
Dim r As ADODB.Recordset, medone As Boolean

If Index = 2 Or Index = 3 Or Index > 5 Then

id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  Exit Sub
End If
nwert$ = strrepl(strrepl(trm(Text1(Index).text), "'", "´"), ";", ",")
If nwert$ <> prv$ Then
  fld$ = transo(Label1(Index).Caption)
  If Index = 2 Then
    nwert$ = datum2sql(nwert$)
    Call movereminders(id$, nwert$)
    typ$ = transo(Text1(5).text)
    s$ = form1.getusersetting("terminalarm" + typ$, "")
    If s$ <> "" Then
      c$ = "delete from todolist where Betreff='[Wiedervorlage] AT:" + id$ + "'"
      Call form1.sqlqry(c$)
      On Error Resume Next
      dtg = CDate(cut_d1(trm(Text1(Index).text), " "))
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        If dtg > Date Then
          c$ = "": an$ = form1.getuserid()
          If InStr(s$, "|") > 0 Then
            an$ = cut_d2bis(s$, "|")
            If Left$(an$, 1) = "+" Then
              an$ = Mid$(s$, 2)
              c$ = "CCMe"
            End If
          End If
          betreff$ = "[Wiedervorlage] AT:" + id$
          intg$ = cut_d1(s$, "|")
          nachricht$ = "Termin in " + intg$ + " Tagen"
          trgdatum$ = datum2sql(trm(CDate(trm(Text1(Index).text)) - CInt(intg$)))
          trgzeit$ = "00:01"
          medone = False
          Set r = New ADODB.Recordset
          r.CursorLocation = adUseServer
          rrr = form1.adoopen(r, "SELECT * FROM benutzergruppen where groupid='" + an$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          If Not r.EOF Then
            While Not r.EOF
              If form1.getuserid() = r!userid Then medone = True
              Call form1.new2do(form1.getuserid(), trm(r!userid), betreff$, nachricht$, trgdatum$, trgzeit$, 0, 0, 0)
              r.MoveNext
            Wend
          Else
            Call form1.new2do(form1.getuserid(), an$, betreff$, nachricht$, trgdatum$, trgzeit$, 0, 0, 0)
          End If
          If c$ = "CCMe" And Not medone And an$ <> form1.getuserid() Then Call form1.new2do(form1.getuserid(), form1.getuserid(), betreff$, nachricht$, trgdatum$, trgzeit$, 0, 0, 0)
        End If
      End If
    End If
  End If
  'If Index = 3 And Len(nwert$) < 8 Then
  '  nwert$ = nwert$ + ":00"
  'End If
  If nwert$ = "" Then
    nwert$ = "' '"
  Else
    nwert$ = "'" + nwert$ + "'"
  End If
  Call chgswrite("update auftritt set " & fld$ & "=" & nwert$ & " where id='" & id$ & "'")
  Call makemedirty
End If

End If 'index=...

End Sub
Sub movereminders(sid$, neuwert$)
Dim rrr
Dim r As ADODB.Recordset, s As ADODB.Recordset, rdtg$
Dim c$, cmd$, i%, dtg$, dto As Variant, movedist As Variant
Dim d2infile As String, d2insub As String, confmode$
'd2infile = "Form1": d2insub = "auftrittstyp"

If form1.isfieldmissing("opt_checks", "id") Then Exit Sub
c$ = "select Datum from auftritt where id='" + sid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Sub
movedist = CDate(neuwert$) - CDate(r!datum)
If movedist <> 0 Then
  c$ = "select id,dtg,confirmed from opt_checks where auftrittsid='" + sid$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    If InStr(trm(r!confirmed), "ok, delete") <> 1 And InStr(trm(r!confirmed), "ok, confirme") <> 1 Then
      dto = CDate(datfromsql(r!dtg)) + movedist
      chgs.AddItem "update opt_checks set dtg='" + datum2sql(dto) + "' where id='" + trm(r!id) + "'"
    End If
    r.MoveNext
  Wend
End If
End Sub

Public Sub initfields(typ$, base%, inimode%)
Dim rtmp As ADODB.Recordset, stmp As ADODB.Recordset, c$, r As ADODB.Recordset, rrr
Dim idx%, anzfelder%, awert1$, hdrs%, lbl$, c1$, fri%
Dim p2left As Integer, c3offset, l2c$, lbflg As Boolean, lbcount As Integer
Dim tooltipshow%, notcomp%, botm%, bl%, viz As Boolean
Dim wert1 As Double, wert2 As Double, hchg%, mws As Double

Dim d2infile As String, d2insub As String
d2infile = "auftritt": d2insub = "initfields"
Call clearlabels
adrfldlist = "|"
Unload besetzung
Unload tpsel
form1.fastsave_copy = False
notcomp% = 0
botm% = 0
lbcount = 0
addedbydefaults = False
recalcplease = False
Shape1.Height = 6495
' no: 150304 btnTopic.Enabled = False
delmode = False
Unload kalku
aid$ = Text1(0).text
opendir.Picture = Picture3(0).Picture
If aid$ = "" Then Exit Sub

c8% = 0
krcount% = 0
hchg% = -1
If typ$ = "Neuer Auftritt" Or typ$ = "" Then
  anzfelder = 0
Else
  On Error Resume Next
  anzfelder = form1.sqla.TableDefs("usr_" & utabn(typ$)).Fields.Count - 1
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then anzfelder = 0
End If
angezeigtefelder% = anzfelder
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM finanzen where id='" & aid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
mws = form1.getusersetting("auftrittsmwst", form1.getusersetting("mwst", 1900)) / 100
If Not rtmp.EOF Then
  mws = 0
  If Not IsNull(rtmp!mwst) Then mws = rtmp!mwst / 100
  mwst.text = d2db(mws / 100)
End If
mwst.text = d2db(mws)
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT feldname,zeilen FROM auftrittsfelder where typ='" & typ$ & "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub

cmd$ = "SELECT * FROM usr_" & utabn(typ$) & " where id='" & aid$ & "'"
Set stmp = New ADODB.Recordset
stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If stmp.EOF Then
  'noch keine Daten, anlegen sonst ärger mit update
  cmd1$ = "insert into usr_" & utabn(typ$) + " (id) values('" + aid$ + "')"
  chgs.AddItem cmd1$
  Call makemedirty
  'neuer Versuch
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
End If
idx% = 0
ptop = 1560
pleft = 120
'p2left = 3960
p2left = 4560
'c3offset = 3480
c3offset = 4080
tooltipshow% = form1.iwanttooltips()
d0 = Time
For i% = 1 To base%: rtmp.MoveNext: Next i%
tbidx% = 6
While Not rtmp.EOF And idx% < fpp% + 1 And pleft > 0
  fn$ = trm(rtmp!feldname)
  p% = InStr(fn$, ".")
  If p% > 0 Then
    clickgetsfromtable(idx%) = Left$(fn$, p% - 1)
    fn$ = Mid$(fn$, p% + 1)
    p% = InStr(fn$, ".")
    If p% > 0 Then
      clickgetsfromfield(idx%) = Mid$(fn$, p% + 1)
      fn$ = Left$(fn$, p% - 1)
    Else
      If clickgetsfromtable(idx%) = "finanzen" Then
        clickgetsfromfield(idx%) = fn$
      Else
        If clickgetsfromtable(idx%) <> "tabelle" And clickgetsfromtable(idx%) <> "besetzung" And clickgetsfromtable(idx%) <> "Vertragsnummer" Then
          clickgetsfromtable(idx%) = ""
        End If
      End If
    End If
    Label2(idx%).ForeColor = form1.lnkcolor
    Command3(idx%).Visible = True
  End If
  viz = True: lbflg = False
  z% = rtmp!zeilen
  If z% < 0 Then
    z% = Abs(z%)
    viz = False
  End If
  If z% > 10 Then
    hdrs% = Int(z% / 10)
    z% = z% Mod 10
    lbflg = True
  End If
  pheight = 285 + (z% - 1) * 240
  Text2(idx%).TabIndex = tbidx%
  tbidx% = tbidx% + 1
  If viz And (ptop + pheight > Shape1.Top + Shape1.Height) Then
    ptop = 1560
    If pleft = 120 Then
      pleft = p2left
    Else
      pleft = 0
    End If
  End If
  If pleft > 0 Then
    Label2(idx%).Top = ptop
    Text2(idx%).Top = ptop
    Command3(idx%).Top = ptop
    Label2(idx%).Left = pleft
    Text2(idx%).Left = pleft + 1440
    Text2(idx%).Height = pheight
    Command3(idx%).Left = pleft + c3offset
    If pheight > 300 Then
      Command3(idx%).Visible = False
    Else
      Command3(idx%).Height = pheight
    End If
    If viz Then
      ptop = ptop + pheight
      If ptop > Shape1.Top + Shape1.Height Then
        ptop = 1560
        If pleft = 120 Then
          pleft = p2left
        Else
          pleft = 0
        End If
      End If
    End If
    If ptop > botm% Then botm% = ptop
    Label2(idx%).Caption = formtranse(transe(fn$))
    Label2(idx%).Visible = True
    Text2(idx%).Visible = True
    If lbflg Then
      Text2(idx%).text = ""
      Text2(idx%).Visible = False
      lbcount = lbcount + 1
      On Error Resume Next
      Load gd1(lbcount)
      rrr = Err
      On Error GoTo 0
      gd1(lbcount).Left = Text2(idx%).Left
      gd1(lbcount).Top = Text2(idx%).Top
      gd1(lbcount).Width = Text2(idx%).Width + Command3(idx%).Width + (Command3(idx%).Left - (Text2(idx%).Left + Text2(idx%).Width))
      gd1(lbcount).Height = Text2(idx%).Height
      gd1(lbcount).Visible = True
      If rrr = 0 Then
        Set colHeader = gd1(lbcount).ColumnHeaders.add(, , "?", 300)
        For i% = 1 To hdrs%
          If i% = 1 Then
            lbl$ = transe("Name")
          Else
            lbl$ = form1.getsystemsetting(transo(Text1(5).text) + "_" + formtranso(Label2(idx%).Caption) + "_" + trm(i%), trm(i%))
          End If
          Set colHeader = gd1(lbcount).ColumnHeaders.add(, , lbl$, (gd1(lbcount).Width - 600) / hdrs%)
        Next i%
      End If
      c$ = "SELECT * FROM auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='" & transo(formtranso(Label2(idx%).Caption)) & "' and auftrittstyp='" & transo(Text1(5).text) & "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        c$ = trm(r!felddaten)
        If c$ <> "" Then
          For i% = 1 To linesof(c$)
            c1$ = lineof(i%, c$)
            Set lvitem = gd1(lbcount).ListItems.add(, , trm(i%))
            p% = 1
            Do
              wert$ = cut_d1(c1$, "|")
              c1$ = cut_d2bis(c1$, "|")
              lvitem.SubItems(p%) = wert$
              p% = p% + 1
            Loop Until c1$ = ""
          Next i%
        End If
      End If
      GoTo skptxtfuncs
    End If
    If Not viz Then
      Label2(idx%).Visible = False
      Text2(idx%).Visible = False
      Command3(idx%).Visible = False
    End If
    wert$ = ""
    If Not stmp.EOF Then
      wert$ = form1.fieldnameonly(rtmp!feldname)
      On Error Resume Next
      wert$ = trm(stmp.Fields(wert$).value)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then wert$ = ""
    End If
    wert$ = unx2dos(wert$)
    Text2(idx%).text = wert$
    If InStr(LCase(transo(formtranso(Label2(idx%).Caption))), "honorar") = 1 Or _
        InStr(transo(LCase(formtranso(Label2(idx%).Caption))), "auslastung") > 0 Or _
        InStr(transo(LCase(formtranso(Label2(idx%).Caption))), "betrag") > 0 Then
      c$ = "SELECT * FROM auftritthigru where auftrittsid='" + Text1(0).text & "' and feldname='" & transo(formtranso(Label2(idx%).Caption)) & "' and auftrittstyp='kalku_" & transo(Text1(5).text) & "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        Label2(idx%).ForeColor = RGB(0, 0, 255)
      End If
      r.Close
    End If
    If transo(formtranso(Label2(idx%).Caption)) = "Honorar" Then
      On Error Resume Next
      w1$ = stmp.Fields("betrag_pro_stunde").value
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        waehr$ = form1.nurdiewaehrung(w1$)
        w1$ = word1(w1$): wert1 = 0
        On Error Resume Next: wert1 = var2dbl(w1$): On Error GoTo 0
        On Error Resume Next
        w2$ = stmp.Fields("dauer").value
        rrr = Err
        On Error GoTo 0
        If rrr = 0 Then
          w2$ = word1(w2$): wert2 = 0
          On Error Resume Next: wert2 = var2dbl(w2$): On Error GoTo 0
          Text2(idx%).text = Format$(wert1 * wert2, "0.00") + " " + waehr$
          Text2(idx%).Enabled = False
        End If
      End If
      If Text2(idx%).text <> wert$ Then hchg% = idx%
    End If
    If clickgetsfromtable(idx%) = "adrselect" And Len(wert$) > 0 Then
      kalrestrict$(krcount%) = wert$
      krcount% = krcount% + 1
      adrfldlist = adrfldlist + transo(formtranso(Label2(idx%).Caption)) + "|"
    End If
    If Len(wert$) = 0 And inimode% <> 0 Then
      tpid$ = Text1(1).text
      If Len(tpid$) = 0 Then
        Text2(idx%).text = ""
      Else
        l2c$ = transo(formtranso(Label2(idx%).Caption))
        If l2c$ = "Veranstalter" _
           Or l2c$ = "Künstler" _
           Or l2c$ = "Partner" _
           Or l2c$ = "Dirigent" _
           Or l2c$ = "Tourneeleitung" _
           Or l2c$ = "Projektbetreuer" _
           Or l2c$ = "Enddatum" _
           Or l2c$ = "Solist" _
           Or l2c$ = "Orchester" _
           Then
        wert$ = transo(formtranso(Label2(idx%).Caption))
        If wert$ = "Partner" Then wert$ = "mehr_Solisten"
        If wert$ = "Enddatum" Then wert$ = "bis"
        awert1$ = wert$
        If wert$ = "Künstler" Then awert1$ = "Solist"
        wert$ = form1.getfromtplan(tpid$, awert1$)
        If Len(wert$) > 0 Or awert1$ = "bis" Then
          If awert1$ = "bis" Then
          If wert$ = "" Then
            wert$ = Text1(2).text
          Else
            wert$ = datfromsql(wert$)
          End If
          End If
          Text2(idx%).text = wert$
          c8% = 1
          Text2(idx%).Enabled = False
          notcomp% = 1
        End If
        End If
      End If
    End If
skptxtfuncs:
    If trm(Text2(idx%)) <> "" And (transo(Text1(6).text) = "Neuer Auftritt" Or trm(Text1(6).text) = "") Then
      Label2(idx%).ToolTipText = transe("Rechtsklick") + ": " + trm(Text2(idx%).text) + " -> " + transe("Bezeichnung")
    Else
      Label2(idx%).ToolTipText = ""
    End If
    idx% = idx% + 1
  Else
    Command3(idx%).Visible = False
  End If
  rtmp.MoveNext
Wend
angezeigtefelder% = idx%
d1 = Time

Command4.Visible = False
Command5.Visible = False
pbase% = 0
Label3.Caption = base% + idx%
basemerk.Caption = base%
v0base% = base%
b% = base% - (fpp% + 1): If b% < 0 Then b% = 0
Label4.Caption = b%
If base% > 0 Then Command4.Visible = True
Call form1.dbg2f("base(" + trm(base%) + ") + idx(" + trm(idx%) + ") < anzfelder(" + trm(anzfelder%) + ") ?")
If base% + idx% < anzfelder% Then
  Call form1.dbg2f("Command5 set to visible @" + trm(Command5.Top) + "/" + trm(Command5.Left))
  Command5.Visible = True
End If
If c8% = 1 Then
  Command8.Enabled = True
Else
  Command8.Enabled = False
End If
BackColor = form1.cleancolor()
If hchg% >= 0 Then
  wert$ = Text2(hchg%).text
  Text2(hchg%).text = "0"
  Call Text2_GotFocus(hchg%)
  Text2(hchg%).text = wert$
  Call Text2_LostFocus(hchg%)
  Call makemedirty
End If
If notcomp% = 1 Then Call Command8_Click
bl% = botm% - Shape1.Top + 120
Shape1.Height = bl%
bl% = bl% + Shape1.Top + 120
Command1.Top = bl%
Command18.Top = bl%
opendir.Top = bl%
Command20.Top = bl%
Command11.Top = bl%
Command4.Top = bl%
Command10.Top = bl%
'dtst.Top = bl%
wert$ = LCase(form1.getuserid())
'Command10.Width = Command10.Left - 40
Command13.Top = bl%
btnTopic.Top = bl%
'Command15.Top = bl%
delme.Top = bl%
Command5.Top = bl%
Command12.Top = bl%
Command42.Top = bl%
kalimmer.Top = bl%
Label7.Top = bl%
kalres.Top = bl% + delme.Height / 2
Label6.Top = kalres.Top
wvl.Top = bl%
Height = delme.Top + delme.Height + 620
'calcol.BackColor = form1.get_eventcolor(typ$)

End Sub

Private Sub Text2_Change(Index As Integer)
  Call makemedirty

End Sub

Private Sub Text2_DblClick(Index As Integer)
Dim tx$, p%, l$, fn$, X

'd2infile = "auftritt": d2insub = "Text2_DblClick"
tx$ = Text2(Index).text
Clipboard.Clear
If LCase(Left$(tx$, 6)) = "datei:" Or LCase(Left$(tx$, 3)) = "fn:" Then
  p% = InStr(tx$, ":") + 1
  tx$ = Mid$(tx$, p%)
  p% = FreeFile
  If tx$ = "" Then Exit Sub
  Open tx$ For Input As #p%
  tx$ = ""
  While Not EOF(p%)
    Line Input #p%, l$
    If tx$ <> "" Then tx$ = tx$ & vbCrLf
    tx$ = tx$ & l$
  Wend
  Close #p%
  GoTo happtx
End If
If InStr(LCase(transo(formtranso(Label2(Index).Caption))), "programm") = 1 Then
  tx$ = form1.rdprog(tx$)
  fn$ = auftrittsverzeichnis() + "\programm.txt"
  If tx$ = "" And nexist(fn$) Then
    p% = FreeFile
    Open fn$ For Output As #p%
    Close #p%
  End If
  If exist(fn$) Then X = Shell("notepad.exe " + fn$, 1)
  GoTo happtx
End If
happtx:
Clipboard.settext tx$
End Sub

Public Sub Text2_GotFocus(Index As Integer)

'd2infile = "auftritt": d2insub = "Text2_GotFocus"
prvd$ = Text2(Index).text
End Sub

Public Sub Text2_LostFocus(Index As Integer)
Dim s$, nwert$
  
id$ = Text1(0).text
nwert$ = strrepl(trm(Text2(Index).text), "'", "´")
If Len(nwert) > 30000 Then
  nwert = Left(nwert, 30000)
  Text2(Index).text = nwert
  MsgBox ("Maximal 30000 Zeichen erlaubt. Die Daten wurden abgeschnitten.")
End If
typ$ = transo(Text1(5).text)
If nwert$ <> prvd$ Then
  
  fld$ = transo(formtranso(Label2(Index).Caption))
  If fld$ = "Vertragsnummer" Then
    ask% = MsgBox("Wollen Sie wirklich die Vertragsnummer ändern?", vbYesNo + vbCritical + vbDefaultButton2, "Vertragsnummer ändern?")
    If ask% <> vbYes Then
      Text2(Index).text = prvd$
      Exit Sub
    End If
  End If
  If nwert$ <> "" Then
    nwert$ = "'" + nwert$ + "'"
  End If
  Call makemedirty
  chgs.AddItem "delete from auftritthigru where feldname='" + fld$ + "' and auftrittstyp='" + typ$ + "' and auftrittsid='" + id$ + "'"
  If nwert$ <> "" Then
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,feldname,felddaten) values('" + form1.newid("auftritthigru", "id", 20) + "','" + id$ + "','" + typ$ + "','" + fld$ + "'," + nwert$ + ")"
    Call chgswrite(cmd$)
  End If
  If nwert$ = "" Then nwert$ = "' '"
  cmd$ = "update usr_" & utabn(typ$) + " set " + fld$ + "=" + nwert$ + " where id='" + id$ + "'"
  Call chgswrite(cmd$)
  If InStr(recalclist, "|" + LCase(fld$) + "|") > 0 Then recalcplease = True
End If

End Sub

Public Function isvis(j%)

'd2infile = "auftritt": d2insub = "isvis"
isvis = Label2(j%).Visible

End Function
Private Sub savecheck()
'd2infile = "auftritt": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
  If form1.fastsave_copy Or form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command10_Click
  End If
End If
BackColor = form1.cleancolor()
End Sub

Private Sub Text2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim idx%, wert$, mwhr$
Dim xnet As Double, xwae$, xdat$

idx% = Index
  wert$ = Text2(idx%).text
  If prvtt$ <> wert$ Then
    prvtt$ = wert$
    If clickgetsfromtable(idx%) = "adrselect" And Len(wert$) > 0 Then
      Call ttset(idx%, form1.gettelfaxmail(wert$), X, Y)
      Exit Sub
    End If
    If clickgetsfromtable(idx%) = "programm" And Len(wert$) > 0 Then
      Call ttset(idx%, form1.getwerke(wert$), X, Y)
      Exit Sub
    End If
    lbl$ = LCase(transo(Label2(idx%).Caption))
    If InStr(lbl$, "honorar") > 0 And InStr(lbl$, "honoraranmerk") = 0 And InStr(lbl$, "anmerk_honorar") = 0 And InStr(lbl$, "honorar_anmerk") = 0 Then
      On Error Resume Next
      xwae$ = form1.nurdiewaehrung(trm0(Text2(idx%).text))
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then Exit Sub
      mwhr$ = form1.getusersetting("MeineWaehrung", transe("€"))
      If xwae$ <> "" And mwhr$ <> xwae$ Then
        On Error Resume Next
        xnet = CDbl(form1.ohnewaehrung(trm(Text2(idx%).text)))
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then Exit Sub
        xdat$ = trm(datum2sql(Text1(2).text))
        xnet = form1.xkurs(xdat, xwae$, xnet)
        xdat$ = form1.kursdatum(xwae$, xdat$)
        xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
        Call ttset(idx%, xdat$ + " / " + fixeur(xnet) + " " + trm(xwae$), X, Y)
        Exit Sub
      End If
    End If
    ttform.Hide
    tlform.Hide
  End If

End Sub

Private Sub Text3_Change()
Call makemedirty
End Sub

Private Sub Text3_GotFocus()
prvt3$ = Text3.text
End Sub

Private Sub Timer_dtst_Timer()
Dim Index As Integer, col As Long

'd2infile = "auftritt": d2insub = "Timer_dtst_Timer"
Call form1.dbg2f("auftritt Timer dtst start")
Timer_dtst.Enabled = False
col = RGB(0, 0, 0)
For Index = 0 To angezeigtefelder% - 1
  Text2(Index).ForeColor = col
Next Index
Call form1.dbg2f("auftritt Timer dtst exit")
End Sub

Private Sub Timer2_Timer()
Dim c As Long, cmd$, id$
Dim w As Long, r As Long, g As Long, b As Long

'd2infile = "auftritt": d2insub = "Timer2_Timer"
c = form1.getcolorselected()

If c < -10 Then Exit Sub
Timer2.Enabled = False
If c < 0 Then Exit Sub
Call form1.dbg2f("auftritt Timer2 start")
id$ = Text1(0).text
b = c / 65536
w = c Mod 65536
g = w / 256
r = w Mod 256
calcol.BackColor = c
cmd$ = "update auftritt set optkalcolor='" + trm(c) + "' where id='" + id$ + "';"
Call form1.sqlqry(cmd$)
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click
Call form1.dbg2f("auftritt Timer2 exit")

End Sub

Private Sub vwopts_Click()

'd2infile = "auftritt": d2insub = "vwopts_Click"
If Text1(0).text = "" Then Exit Sub
Load termviewlist
On Error Resume Next
Call termviewlist.SetFocus
On Error GoTo 0
termviewlist.Caption = transe("Sichtbarkeitseinstellungen")
termviewlist.aid.Caption = Text1(0).text
End Sub

Private Sub wvl_Click()

'd2infile = "auftritt": d2insub = "wvl_Click"
Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), Text1(6) & " [Wiedervorlage] Auftritt:" + _
               Text1(0).text, "", Date, Left(Time, 5))
Call create2do.SetFocus
create2do.Text1(1).Enabled = False
create2do.Text1(3).Enabled = False

End Sub

Public Sub recalc()
Dim mbase%, repfl As Boolean

'd2infile = "auftritt": d2insub = "recalc"
mbase% = v0base%
recalcplease = False
Do
  repfl = False
  Call Command11_Click
  DoEvents
  If recalcpg() Then repfl = True
  While Command5.Visible
    Call Command5_Click
    If recalcpg() Then repfl = True
    DoEvents
  Wend
Loop Until Not repfl
If mbase% <> v0base% Then Call initfields(transo(Text1(5).text), mbase%, 0)

End Sub

Function recalcpg() As Boolean
Dim i%, cgft$, ein$, l2c As String

'd2infile = "auftritt": d2insub = "recalcpg"
recalcpg = False
For i% = 0 To 33
If Label2(i%).Visible Then
cgft$ = clickgetsfromtable(i%)
l2c = transo(LCase(formtranso(Label2(i%).Caption)))
If InStr(l2c, "honorar") = 1 Or _
   InStr(l2c, "betrag") > 0 Or _
   InStr(l2c, "preis") > 0 Or _
   LCase(l2c) = "konzertauslastung" Or _
   LCase(cgft$) = "finanzen" Or _
   LCase(cgft$) = "tabelle" Then
  ein$ = Text2(i%).text
  If LCase(cgft$) = "tabelle" Then
    Unload tabkalk
    DoEvents
    Load tabkalk
    tabkalk.Hide
    On Error Resume Next
    Call tabkalk.SetFocus
    On Error GoTo 0
    tabkalk.Label2.Caption = Text1(0).text
    tabkalk.Label3.Caption = formtranso(Label2(i%).Caption)
    tabkalk.Label1.Caption = transo(Text1(5).text)
    DoEvents
    On Error Resume Next
    Call tabkalk.SetFocus
    On Error GoTo 0
    Call tabkalk.Command2_Click
    DoEvents
    Unload tabkalk
  Else
    Unload kalku
    DoEvents
    Load kalku
    kalku.Hide
    kalku.afeld = formtranso(Label2(i%).Caption)
    kalku.kerg.Caption = Text2(i%).text
    kalku.atyp = transo(Text1(5).text)
    kalku.aid = Text1(0).text
    DoEvents
    On Error Resume Next
    Call kalku.SetFocus
    On Error GoTo 0
    Call kalku.Command5_Click
    DoEvents
    Unload kalku
  End If
  If ein$ <> Text2(i%).text Then
    recalcpg = True
  End If
End If
End If
Next i%

End Function

Public Sub clearlabels()
Dim i%

'd2infile = "auftritt": d2insub = "clearlabels"
i% = 1
Do
  On Error Resume Next
  gd1(i%).ListItems.Clear
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    gd1(i%).Visible = False
    i% = i% + 1
  End If
Loop Until rrr <> 0
For i% = 0 To fpp%
  clickgetsfromtable(i%) = ""
  Label2(i%).ForeColor = &H80000012
  Label2(i%).Visible = False
  Text2(i%).Visible = False
  Command3(i%).Visible = False
Next i%

End Sub

Function labelnumbylistnum(j%) As Integer
Dim i%, rrr, pt As Integer, pl As Integer

'd2infile = "auftritt": d2insub = "labelnumbylistnum"
labelnumbylistnum = -1
i% = 0
Do
  On Error Resume Next
  pl = Text2(i%).Left
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    pt = Text2(i%).Top
    If gd1(j%).Top = pt And gd1(j%).Left = pl Then
      labelnumbylistnum = i%
      Exit Function
    End If
    i% = i% + 1
  End If
Loop Until i% > 33 Or rrr <> 0

End Function
Function listnnumbytextpos(j%) As Integer
Dim i%, rrr, pt As Integer, pl As Integer

'd2infile = "auftritt": d2insub = "listnnumbytextpos"
listnnumbytextpos = -1

i% = 1
Do
  On Error Resume Next
  pl = gd1(i%).Left
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    pt = gd1(i%).Top
    If Text2(j%).Top = pt And Text2(j%).Left = pl Then
      listnnumbytextpos = i%
      Exit Function
    End If
    i% = i% + 1
  End If
Loop Until rrr <> 0

End Function

Sub gd1after(Index As Integer, sortmode As String)
Dim j As Integer, i As Integer, gdcmd$, allgdcmd$, ad$, c$, c1$, p%

'd2infile = "auftritt": d2insub = "gd1after"
tmpsort.Clear
i = Index

    For j = 1 To gd1(i).ListItems.Count
      gd1(i).SelectedItem = gd1(i).ListItems(j)
      If gd1(i).SelectedItem.text <> "d" And gd1(i).SelectedItem.text <> "delete" Then
        gdcmd = gd1(i).SelectedItem.text
        allgdcmd = gdcmd
        ad$ = ""
        If InStr(gdcmd, "-") > 1 Then
          allgdcmd = Left(gdcmd, InStr(gdcmd, "-") - 1)
          ad$ = Mid(gdcmd$, InStr(gdcmd, "-"))
        End If
        Select Case sortmode
          Case "namsort": gdcmd = ""
          Case Else: gdcmd = Left("000000000", 8 - Len(allgdcmd$)) + allgdcmd$ + ad$
        End Select
        For p = 1 To gd1(i).ColumnHeaders.Count - 1
          If gdcmd <> "" Then gdcmd = gdcmd + "|"
          gdcmd = gdcmd + gd1(i).SelectedItem.SubItems(p)
        Next p
        tmpsort.AddItem gdcmd
      End If
    Next j
    gdcmd$ = ""
    For j = 0 To tmpsort.ListCount - 1
      If gdcmd <> "" Then gdcmd = gdcmd + vbCrLf
      Select Case sortmode
        Case "namsort": gdcmd = gdcmd + tmpsort.List(j)
        Case Else: gdcmd = gdcmd + cut_d2bis(tmpsort.List(j), "|")
      End Select
    Next j
    c$ = gdcmd$
    gd1(i).ListItems.Clear
    If c$ <> "" Then
          For j = 1 To linesof(c$)
            c1$ = lineof(j, c$)
            Set lvitem = gd1(i).ListItems.add(, , trm(j))
            p% = 1
            Do
              wert$ = cut_d1(c1$, "|")
              c1$ = cut_d2bis(c1$, "|")
              lvitem.SubItems(p%) = wert$
              p% = p% + 1
            Loop Until c1$ = ""
          Next j
    End If
End Sub

Sub makemedirty()
'd2infile = "auftritt": d2insub = "makemedirty"
  BackColor = form1.dirtycolor()
  Command10.Enabled = True
  Command13.Enabled = False
End Sub

Function formtranse(t$)
Dim i%, rc$

rc$ = t$
For i% = 0 To formttp
  If InStr(formtt(i%), t$ + "|") = 1 Then
    p% = InStr(formtt(i%), "|") + 1
    rc$ = Mid(formtt(i%), p%)
    Exit For
  End If
Next i%
formtranse = rc$

End Function
Function formtranso(t$)
Dim i%, rc$

rc$ = t$
For i% = 0 To formttp
  If InStr(formtt(i%), "|" + t$) > 0 Then
    p% = InStr(formtt(i%), "|") - 1
    rc$ = Left(formtt(i%), p%)
    Exit For
  End If
Next i%
formtranso = rc$

End Function

Sub ttset(idx%, tt$, X, Y)
Dim fx As Double, fy As Double, pfx As Double, pfy As Double
Dim ttfT, ttfL, wd As Integer, ttedlm As String, ttlines As Integer
Dim j%, wdw As Long


  If tt$ = "" Then
    ttform.Hide
    tlform.Hide
  Else
      wdw = 2
      For j% = 0 To linesof(tt$)
        If Len(lineof(j%, tt$)) > wdw Then
          wdw = Len(lineof(j%, tt$))
          Debug.Print wdw
        End If
      Next j%
      ttfL = auftritt.Left + Text2(idx%).Left + Text2(idx%).Width
      ttfT = auftritt.Top + Text2(idx%).Top + Text2(idx%).Height / 2 + 200
      wdw = 20
      ttform.Width = wdw * 100
      ttform.Text1.Width = ttform.Width
      ttform.Top = ttfT
      ttform.Left = ttfL
      ttlines = linesof(tt$) + 1
      ttform.Hide
      DoEvents
      ttform.Text1.text = tt$
      ttform.Height = 200 * ttlines
      ttform.Text1.Height = ttform.Height
      ttform.Show
      'ttform.SetFocus
  End If

End Sub

Sub tlset(tti, X, Y)
Dim fx As Double, fy As Double, pfx As Double, pfy As Double
Dim ttfT, ttfL, wd As Integer, ttedlm As String, ttlines As Integer
Dim j%, wdw As Long, ll$, tt$

tt$ = tti

  If tt$ = "" Then
    tlform.Hide
  Else
      wdw = 2
      For j% = 0 To linesof(tt$)
        ll$ = lineof(j%, tt$)
        If Len(ll$) > wdw Then
          wdw = InStr(lineof(j%, tt$), "  ")
          Debug.Print wdw
          tlform.List1.AddItem ll$
        End If
      Next j%
      tlform.Width = wdw * 80 + 100
      ttlines = linesof(tt$) + 1
      tlform.Height = 200 * ttlines
      tlform.List1.Height = tlform.Height
      ttfL = Command42.Left + Me.Left
      ttfT = Command42.Top + Me.Top + 200
      tlform.List1.Width = tlform.Width
      tlform.Top = ttfT
      tlform.Left = ttfL
      tlform.Hide
      DoEvents
'      tlform.Show
      'ttform.SetFocus
  End If

End Sub

Private Sub umkrtest(srcid$, dst$)
Dim r As ADODB.Recordset, c$, rrr, dst1$, dst2$, difft$, srczip As String, d
Dim s As ADODB.Recordset, plzl, diffdays, blocklist As String
Dim rl As ADODB.Recordset, von As String, bis As String, d2 As Integer, tsthalle As String, tstplz As String

If Not form1.geodbok Then Exit Sub
Command42.BackColor = &HFFFF&
Command42.ToolTipText = "unknown, possile error"
dst1$ = word1(dst$)
dst2$ = word2bis(dst$)
If dst2$ = "" Then dst2$ = "90"

blacklist = form1.getusersetting("umkreisblacklist", "auftrittstyp not like '%rob%' ")
blocklist = ""
If transo(Text1(5).text) = "Künstlerauftritt" Then
  c$ = "SELECT * FROM auftritthigru where auftrittsid='" + srcid$ + "' and (feldname like 'Künstle%')"
End If
If transo(Text1(5).text) = "Orchesterauftritt" Then
  c$ = "SELECT * FROM auftritthigru where auftrittsid='" + srcid$ + "' and (feldname like 'Orcheste%')"
End If
If transo(Text1(5).text) = "Dirigentenauftritt" Then
  c$ = "SELECT * FROM auftritthigru where auftrittsid='" + srcid$ + "' and (feldname like 'Dirigen%')"
End If
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not r.EOF
    c$ = trm(r!felddaten)
    If c$ <> "" Then
      If blocklist <> "" Then blocklist = blocklist + " or "
      blocklist = blocklist + "Felddaten='" + c$ + "'"
    End If
    r.MoveNext
  Wend
End If
If blocklist = "" Then
  Command42.BackColor = &HFF&
  Command42.ToolTipText = "nobody to test"
  Exit Sub
End If

c$ = "SELECT * FROM auftritthigru where auftrittsid='" + srcid$ + "' and (feldname='Halle' or feldname='Saal')"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not r.EOF Then
    srczip = trm(r!felddaten)
    srczip = form1.plzausadr("" & srczip & "")
    If srczip <> "" Then
      If dst2$ <> "" Then
        d = Val(dst2$)
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Or d = 0 Then
          d2 = 90
        Else
          d2 = Abs(d)
        End If
        On Error Resume Next
      Else
        d2 = 90
      End If
      d = Val(dst1$)
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Or d = 0 Then Exit Sub
      d = Abs(d)

c$ = "SELECT zc_id, zc_location_name, zc_lat, zc_lon from zip_coordinates WHERE zc_zip = '" + srczip + "'"
Call form1.dbg2f("umkreissuche:" + c$)
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
On Error Resume Next
s.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
If s.EOF Then Exit Sub
c$ = "SELECT "
c$ = c$ + "dest.zc_zip,dest.zc_location_name,ACOS(Sin (RADIANS(src.zc_lat)) * Sin(RADIANS(dest.zc_lat))+ COS(RADIANS(src.zc_lat)) * COS(RADIANS(dest.zc_lat))* COS(RADIANS(src.zc_lon) - RADIANS(dest.zc_lon))) * 6380 AS distance "
c$ = c$ + "FROM zip_coordinates dest CROSS JOIN zip_coordinates src "
c$ = c$ + "Where src.zc_id = " + trm(s!zc_id) + " And dest.zc_id <> src.zc_id "
c$ = c$ + "Having distance < " + trm(d) + " "
c$ = c$ + "ORDER BY distance;"
s.Close
Call form1.dbg2f("umkreissuche:" + c$)
s.CursorLocation = adUseServer
On Error Resume Next
s.Open c$, form1.geodb, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
If s.EOF Then Exit Sub
While Not s.EOF
  If plzl = "" Then plzl = "|" + srczip + "|"
  plzl = plzl + trm(s!zc_zip) + "|"
  s.MoveNext
  DoEvents
Wend

      von = datum2sql(CDate(trm(Text1(2).text)) - d2)
      bis = datum2sql(CDate(trm(Text1(2).text)) + d2)
      plzlt = ""
      c$ = "select auftrittsid,auftritt.Datum as dtg,auftritt.Bezeichnung,auftritt.auftrittstyp,auftritt.Ort,auftritthigru.Felddaten "
      c$ = c$ + "from auftritthigru INNER JOIN auftritt ON auftritthigru.auftrittsid = auftritt.id "
      c$ = c$ + "where (feldname='Halle' or feldname='Saal') and auftritt.Datum>='" + von + "' and auftritt.Datum<='" + bis + "' "
      c$ = c$ + "order by auftrittsid"
      Set rl = New ADODB.Recordset
      rl.CursorLocation = adUseServer
      rrr = form1.adoopen(rl, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
      If rrr = 0 Then
        While Not rl.EOF
          DoEvents
          If rl!auftrittsid <> srcid$ Then
            tsthalle = trm(rl!felddaten): tstplz = form1.plzausadr("" & tsthalle & "")
            If InStr(plzl, "|" + tstplz + "|") > 0 Then
              c$ = "select count(*) as wert from auftritthigru where auftrittsid='" + rl!auftrittsid + "' and (" + blocklist + ") and (" + blacklist + ")"
              c$ = form1.get1erg(c$)
              If c$ <> "0" Then
                diffdays = Int(CDate(trm(rl!dtg)) - CDate(trm(Text1(2).text)))
                Debug.Print srcid$ + " vs. " + rl!auftrittsid + " " + rl!dtg + " " + trm(diffdays) + " Tg. " + trm(rl!auftrittstyp) + " " + trm(rl!bezeichnung)
                Command42.BackColor = &HFF&
                plzlt = plzlt + trm(rl!dtg) + " " + trm(rl!auftrittstyp) + " " + trm(rl!bezeichnung) + Space$(129) + "(ID:" + (rl!auftrittsid) + vbCrLf
              End If
            End If
          End If
          rl.MoveNext
        Wend
      End If
    End If
    Command42.ToolTipText = ""
    If plzlt <> "" Then
      Call tlset(plzlt, 100, 100)
    Else
      Command42.BackColor = &HC000&
      Command42.ToolTipText = "ok (" + dst1$ + " km, days: " + dst2$ + ")"
    End If
  End If
End If

End Sub

