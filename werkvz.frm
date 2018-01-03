VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Begin VB.Form werkvz 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Werkeverzeichnis"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14415
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8280
      TabIndex        =   133
      Top             =   7800
      Width           =   2655
   End
   Begin VB.ListBox List7 
      Enabled         =   0   'False
      Height          =   495
      IntegralHeight  =   0   'False
      Left            =   12000
      Sorted          =   -1  'True
      TabIndex        =   132
      Top             =   7200
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   12000
      TabIndex        =   131
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   12000
      TabIndex        =   129
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   10
      ToolTipText     =   "Geburtsdatum"
      Top             =   7320
      Width           =   975
   End
   Begin VB.ListBox List6 
      Enabled         =   0   'False
      Height          =   495
      IntegralHeight  =   0   'False
      Left            =   12000
      Sorted          =   -1  'True
      TabIndex        =   127
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   12000
      TabIndex        =   126
      Top             =   5760
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   12000
      TabIndex        =   124
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   615
      IntegralHeight  =   0   'False
      Left            =   12000
      Sorted          =   -1  'True
      TabIndex        =   123
      Top             =   4920
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   12000
      TabIndex        =   121
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   555
      IntegralHeight  =   0   'False
      Left            =   12000
      Sorted          =   -1  'True
      TabIndex        =   119
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command26 
      Caption         =   "<-"
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
      Left            =   3360
      TabIndex        =   118
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Repertoire"
      Height          =   255
      Left            =   8520
      TabIndex        =   117
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtStimmton 
      Height          =   1035
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   116
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List2a 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   113
      ToolTipText     =   "Liste der Werke des markierten Komponisten"
      Top             =   8520
      Width           =   5535
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dokumenenliste"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   112
      ToolTipText     =   "Liste aller gespeicherten Dokumente"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1800
      Picture         =   "werkvz.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   111
      ToolTipText     =   "Komponisten speichern"
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      Height          =   375
      Index           =   1
      Left            =   7320
      Picture         =   "werkvz.frx":018A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   110
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      Height          =   375
      Index           =   0
      Left            =   6960
      Picture         =   "werkvz.frx":0314
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   109
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   7200
      Picture         =   "werkvz.frx":049E
      Style           =   1  'Grafisch
      TabIndex        =   108
      ToolTipText     =   "Komponisten speichern"
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   2760
      Picture         =   "werkvz.frx":0628
      TabIndex        =   43
      Top             =   840
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   11640
      TabIndex        =   63
      Top             =   720
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   11280
      Picture         =   "werkvz.frx":0B4C
      Style           =   1  'Grafisch
      TabIndex        =   62
      ToolTipText     =   "löschen"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2400
      Picture         =   "werkvz.frx":103C
      Style           =   1  'Grafisch
      TabIndex        =   42
      ToolTipText     =   "löschen"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command30 
      Caption         =   "&Export"
      Height          =   255
      Left            =   2280
      TabIndex        =   107
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      Picture         =   "werkvz.frx":152C
      Style           =   1  'Grafisch
      TabIndex        =   105
      ToolTipText     =   "Markiertes Werk in die Zwischenablage kopieren"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Werk kopieren"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8520
      TabIndex        =   104
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "&HTML"
      Height          =   255
      Left            =   240
      TabIndex        =   103
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Abwahl"
      Height          =   255
      Left            =   240
      TabIndex        =   102
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9840
      Top             =   120
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Hinweise in Datei speichern"
      Height          =   255
      Left            =   7680
      TabIndex        =   101
      Top             =   7440
      Width           =   3255
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Ende"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   100
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Pos1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   99
      Top             =   6600
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   6600
      Picture         =   "werkvz.frx":1A5E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   98
      Top             =   7680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   6240
      Picture         =   "werkvz.frx":1F50
      ScaleHeight     =   300
      ScaleWidth      =   285
      TabIndex        =   97
      Top             =   7680
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6840
      Style           =   1  'Grafisch
      TabIndex        =   95
      ToolTipText     =   "Datenverzeichnis für dieses Werk im Explorer öffnen"
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1320
      Picture         =   "werkvz.frx":2442
      Style           =   1  'Grafisch
      TabIndex        =   94
      ToolTipText     =   "Komponistenverzeichnis im Explorer öffnen"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Wrk.übertrgn"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   600
      TabIndex        =   92
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "werkvz.frx":2A6C
      Style           =   1  'Grafisch
      TabIndex        =   91
      ToolTipText     =   "Suche das Werk bei bei Google"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "werkvz.frx":2F5E
      Style           =   1  'Grafisch
      TabIndex        =   90
      ToolTipText     =   "Suche den Künstler bei Google"
      Top             =   3000
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Left            =   2640
      Top             =   1080
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Notenverzeichnis"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   87
      ToolTipText     =   "Liste der Bezugsquellen (alle Werke/Komponisten)"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ab"
      Height          =   255
      Left            =   5880
      TabIndex        =   76
      ToolTipText     =   "Den markierten Satz um eine Position nach unten"
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "auf"
      Height          =   255
      Left            =   5880
      TabIndex        =   75
      ToolTipText     =   "Den markierten Satz um eine Position nach oben"
      Top             =   6600
      Width           =   375
   End
   Begin VB.ListBox List4 
      Height          =   1560
      IntegralHeight  =   0   'False
      Left            =   3480
      OLEDropMode     =   1  'Manuell
      TabIndex        =   74
      ToolTipText     =   "Alle Sätze des Werkes"
      Top             =   4920
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6240
      Picture         =   "werkvz.frx":3450
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Den markierten Satz löschen"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      Picture         =   "werkvz.frx":3940
      Style           =   1  'Grafisch
      TabIndex        =   72
      ToolTipText     =   "Neuen Satz anlegen"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   7320
      TabIndex        =   60
      ToolTipText     =   "Satz suchen"
      Top             =   5400
      Width           =   975
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   360
      Top             =   9120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CheckBox Check4 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check5 
      Height          =   255
      Left            =   8520
      TabIndex        =   79
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "streng"
      Height          =   255
      Left            =   4080
      TabIndex        =   78
      Top             =   3240
      Value           =   1  'Aktiviert
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "wo gespielt?"
      Height          =   255
      Left            =   8520
      TabIndex        =   77
      ToolTipText     =   "Liste der Orte, an denen dieses Werk schon gespielt wurde"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6480
      Picture         =   "werkvz.frx":3CD2
      Style           =   1  'Grafisch
      TabIndex        =   71
      ToolTipText     =   "Neues Werk anlegen"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   1155
      Index           =   17
      Left            =   7680
      MultiLine       =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Interne Notizen"
      Top             =   6240
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   16
      Left            =   6840
      TabIndex        =   19
      ToolTipText     =   "Nummer im Werkverzeichnis"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   15
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Werkverzeichnis"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   14
      Left            =   8520
      TabIndex        =   20
      ToolTipText     =   "Namensergänzung, Bemerkung"
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   13
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Offizieller Name des Werkes"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   12
      Left            =   9360
      TabIndex        =   16
      ToolTipText     =   "Nummer des Werkes"
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6840
      Picture         =   "werkvz.frx":4064
      Style           =   1  'Grafisch
      TabIndex        =   64
      ToolTipText     =   "Werk speichern"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   11
      Left            =   10800
      TabIndex        =   17
      ToolTipText     =   "Tonart"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   12000
      TabIndex        =   58
      Text            =   "Text4"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   8
      Left            =   11280
      TabIndex        =   23
      ToolTipText     =   "Jahr der Fertigstellung der Komposition"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   7
      Left            =   9600
      TabIndex        =   22
      ToolTipText     =   "Jahr des Beginns der Komposition"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   6
      Left            =   7560
      TabIndex        =   21
      ToolTipText     =   "Komplett in welchem Jahr geschrieben"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   24
      ToolTipText     =   "Dauer des Gesamtwerks"
      Top             =   5025
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   1035
      Index           =   4
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   25
      ToolTipText     =   "Musikalische Besetzung"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   50
      Text            =   "Text4"
      Top             =   9255
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   48
      Text            =   "Text4"
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   14
      ToolTipText     =   "wird automatisch zusammengesetzt"
      Top             =   3240
      Width           =   6735
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   45
      Text            =   "Text4"
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      Picture         =   "werkvz.frx":4FA6
      Style           =   1  'Grafisch
      TabIndex        =   41
      ToolTipText     =   "Komponisten speichern"
      Top             =   7080
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   11
      ToolTipText     =   "Todesdatum"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   8
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   1725
      Index           =   7
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Andere Schreibweisen des Komponistennamens"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   35
      Text            =   "Text3"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Todesjahr"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Geburtsjahr"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Biografische Daten"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Vorname(n)"
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Name"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   29
      ToolTipText     =   "ID"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5880
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Werk suchen"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   600
      Picture         =   "werkvz.frx":534D
      Style           =   1  'Grafisch
      TabIndex        =   27
      ToolTipText     =   "Neuer Komponist"
      Top             =   7080
      Width           =   375
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3360
      TabIndex        =   13
      ToolTipText     =   "Liste der Werke des markierten Komponisten"
      Top             =   1320
      Width           =   10695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      Picture         =   "werkvz.frx":56DF
      Style           =   1  'Grafisch
      TabIndex        =   96
      ToolTipText     =   "schließen"
      Top             =   7560
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Komponisten suchen"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Ausgewählten anklicken"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2400
      Picture         =   "werkvz.frx":592F
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   82
      ToolTipText     =   "löschen verboten"
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   11280
      Picture         =   "werkvz.frx":5E53
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   83
      ToolTipText     =   "löschen verboten"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Librettist"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   138
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   137
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrangement"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   136
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Co-Composer"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   135
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Album:"
      Height          =   255
      Left            =   7680
      TabIndex        =   134
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Librettist / Textdichter"
      Height          =   255
      Left            =   12000
      TabIndex        =   130
      Top             =   6660
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "GEMA #:"
      Height          =   255
      Left            =   11040
      TabIndex        =   128
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher / Verlag"
      Height          =   255
      Left            =   12000
      TabIndex        =   125
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrangement"
      Height          =   255
      Left            =   12000
      TabIndex        =   122
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "add. Composers / CoKomp."
      Height          =   255
      Left            =   12000
      TabIndex        =   120
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lbl_ston 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stimmton"
      Height          =   255
      Left            =   8400
      TabIndex        =   115
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "---------------------"
      Height          =   255
      Left            =   10680
      TabIndex        =   114
      ToolTipText     =   "vollständig und korrekt eingegeben"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Stand:"
      Height          =   255
      Left            =   6840
      TabIndex        =   106
      ToolTipText     =   "vollständig und korrekt eingegeben"
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "autorisiert"
      Height          =   255
      Left            =   8760
      TabIndex        =   89
      ToolTipText     =   "vollständig und korrekt eingegeben"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "autorisiert"
      Height          =   255
      Left            =   0
      TabIndex        =   88
      ToolTipText     =   "vollständig und korrekt eingegeben"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   15
      Left            =   3480
      TabIndex        =   68
      Top             =   3975
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   14
      Left            =   7560
      TabIndex        =   67
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   66
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   12
      Left            =   8760
      TabIndex        =   65
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   11
      Left            =   10080
      TabIndex        =   61
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   16
      Left            =   5760
      TabIndex        =   69
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sätze"
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
      Left            =   3480
      TabIndex        =   86
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Suchen"
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
      Left            =   3360
      TabIndex        =   85
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Werke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   84
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   46
      Top             =   3255
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2895
      Left            =   3360
      Shape           =   4  'Gerundetes Rechteck
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   3360
      Shape           =   4  'Gerundetes Rechteck
      Top             =   3120
      Width           =   8535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Komponist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   81
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Suchen"
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
      Left            =   240
      TabIndex        =   80
      ToolTipText     =   "Anfrage bei Werke@AgencyProf.de"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   70
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   59
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   11520
      TabIndex        =   57
      Top             =   9375
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   8
      Left            =   10320
      TabIndex        =   56
      Top             =   4575
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   7
      Left            =   8400
      TabIndex        =   55
      Top             =   4575
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   54
      Top             =   4575
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Min."
      Height          =   255
      Left            =   7920
      TabIndex        =   53
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   52
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   51
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   49
      Top             =   9270
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   47
      Top             =   9255
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   44
      Top             =   9255
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   40
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   39
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   38
      Top             =   8640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   37
      Top             =   4920
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   36
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   34
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   32
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   28
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   8055
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   8055
      Left            =   3240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "werkvz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tm_brk%, list2updno As Boolean
Dim callback$, cfn$(19), cfw$(19), swrd$(5), searching As Boolean
Public isviz As Boolean

Sub rlist1()
Dim rtmp As ADODB.Recordset, rrr

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist1"
List1.Clear
List2.Clear
List3.Clear
List5.Clear
List6.Clear
List7.Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT ID,name,vornamen FROM k_loc", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  List1.AddItem "" & rtmp!name & ", " & rtmp!vornamen & Space$(80) & "  (ID:" + rtmp!id + ")"
  rtmp.MoveNext
Wend
rtmp.Close

End Sub
Sub rlist2(kid$)
Dim rtmp As ADODB.Recordset, sorder$, rrr, c$

List3.Clear
List5.Clear
Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist2"
sorder$ = form1.getusersetting("sortierewerke", "name")
List2.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM w_loc where Komponistennummer='" + kid$ + "' order by " & sorder$
Call form1.dbg2f("werkvz.rlist2:" & c$)
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!name) Then
    List2.AddItem rtmp!name & ", " & rtmp!Dauer & " " + transe("Minuten") & Space$(160) & "(WID:" + rtmp!id
  End If
  rtmp.MoveNext
Wend
rtmp.Close
Command7.Enabled = True
End Sub
Sub rlist3()
Dim rtmp As ADODB.Recordset, rrr, na$, wid$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist3"
List3.Clear
If form1.isfieldmissing("opt_cocomposers", "id") Then Exit Sub
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT kid,wid from opt_cocomposers where wid='" + wid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  na$ = form1.getkompnamebyid(trm(rtmp!kid))
  List3.AddItem na$ & Space$(160) & "  (KID:" + rtmp!kid
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Private Sub Check1_Click()

'd2infile = "werkvz": d2insub = "Check1_Click"
If List1.ListIndex < 0 Then
  Check1.value = 1
  Exit Sub
End If
If Check1.value = 1 Then
  Command4.Visible = False
Else
  If List2.ListCount > 0 Then
    MsgBox transe("Komponist hat Werke. Löschen nicht möglich.")
    Check1.value = 1
    Command4.Visible = False
    Exit Sub
  End If
  Command4.Visible = True
End If

End Sub

Private Sub Check2_Click()
'd2infile = "werkvz": d2insub = "Check2_Click"
If List2.ListIndex < 0 Then
  Check1.value = 1
  Exit Sub
End If
If Check2.value = 1 Then
  Command5.Visible = False
Else
Command5.Visible = True
End If

End Sub

Private Sub Check3_Click()
Dim streng%

'd2infile = "werkvz": d2insub = "Check3_Click"
streng% = Check3.value
If streng% = 1 Then
  Text4(1).Enabled = False
Else
  Text4(1).Enabled = True
End If

End Sub

Private Sub Check4_Click()
Dim kid$

'd2infile = "werkvz": d2insub = "Check4_Click"
If List1.ListIndex < 0 Then Exit Sub

kid$ = List1.List(List1.ListIndex)
kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
kid$ = Left$(kid$, InStr(kid$, ")") - 1)
form1.sqlqry ("delete from aut_werke where tabid ='" + kid$ + "' and tabelle='k_loc'")
If Check4.value = 1 Then
  form1.sqlqry ("insert into aut_werke (id,tabid,tabelle) values('" + form1.newid("aut_werke", "id", 20) + "','" + kid$ + "','k_loc')")
End If

End Sub

Private Sub Check5_Click()
Dim kid$

'd2infile = "werkvz": d2insub = "Check5_Click"
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then
  Check5.value = 0
  Exit Sub
End If
kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
form1.sqlqry ("delete from aut_werke where tabid ='" + kid$ + "' and tabelle='w_loc'")
If Check5.value = 1 Then
  form1.sqlqry ("insert into aut_werke (id,tabid,tabelle) values('" + form1.newid("aut_werke", "id", 20) + "','" + kid$ + "','w_loc')")
End If

End Sub

Private Sub Combo1_Click()
Dim id$, p%, wid$, c$, na$, i%

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

id$ = Combo1.text
p% = InStr(id$, "(KID:")
If p% = 0 Then Exit Sub
na$ = trm(Left$(id$, p% - 1))
id$ = Mid$(id$, p% + 5)
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
For i% = 0 To List3.ListCount - 1
  If na$ = List3.List(i%) Then Exit Sub
Next i%
c$ = "insert into opt_cocomposers (kid,wid) values('" + id$ + "','" + wid$ + "')"
Call form1.sqlqry(c$)
Call rlist3
Combo1.text = ""
End Sub

Private Sub Combo1_DropDown()
Dim c$, rrr, wid$
Dim rtmp As ADODB.Recordset

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,name,vornamen FROM k_loc where name like '%" + trm(Combo1.text) + "%' or vornamen like '%" + trm(Combo1.text) + "%' or name like '" + trm(Combo1.text) + "%' or vornamen like '" + trm(Combo1.text) + "%' order by name,vornamen"
Combo1.Clear
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!name) Then
    Combo1.AddItem rtmp!name & ", " & rtmp!vornamen & Space$(160) & "(KID:" + rtmp!id
  End If
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Private Sub Combo2_Click()
Dim id$, p%, wid$, c$, na$, i%

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

na$ = Combo2.text
If na$ = "" Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
For i% = 0 To List5.ListCount - 1
  If na$ = List5.List(i%) Then Exit Sub
Next i%
c$ = "insert into opt_arranged (aid,wid) values('" + na$ + "','" + wid$ + "')"
Call form1.sqlqry(c$)
Call rlist5
Combo2.text = ""

End Sub

Private Sub Combo2_DropDown()
Dim c$, rrr, wid$
Dim rtmp As ADODB.Recordset

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id FROM adresse where id like '%" + trm(Combo2.text) + "%' or id like '" + trm(Combo2.text) + "%' order by id limit 0,30"
Combo2.Clear
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!id) Then
    Combo2.AddItem rtmp!id
  End If
  rtmp.MoveNext
Wend
rtmp.Close

End Sub


Private Sub Combo3_Click()
Dim id$, p%, wid$, c$, na$, i%

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

na$ = Combo3.text
If na$ = "" Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
For i% = 0 To List6.ListCount - 1
  If na$ = List6.List(i%) Then Exit Sub
Next i%
c$ = "insert into opt_published (aid,wid) values('" + na$ + "','" + wid$ + "')"
Call form1.sqlqry(c$)
Call rlist6
Combo3.text = ""

End Sub

Private Sub Combo4_Click()
Dim id$, p%, wid$, c$, na$, i%

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

na$ = Combo4.text
If na$ = "" Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
For i% = 0 To List7.ListCount - 1
  If na$ = List7.List(i%) Then Exit Sub
Next i%
c$ = "insert into opt_textdichter (aid,wid) values('" + na$ + "','" + wid$ + "')"
Call form1.sqlqry(c$)
Call rlist7
Combo4.text = ""

End Sub

Private Sub Combo3_DropDown()
Dim c$, rrr, wid$
Dim rtmp As ADODB.Recordset

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id FROM adresse where id like '%" + trm(Combo3.text) + "%' or id like '" + trm(Combo3.text) + "%' order by id limit 0,30"
Combo3.Clear
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!id) Then
    Combo3.AddItem rtmp!id
  End If
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Private Sub Combo4_DropDown()
Dim c$, rrr, wid$
Dim rtmp As ADODB.Recordset

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id FROM adresse where id like '%" + trm(Combo3.text) + "%' or id like '" + trm(Combo4.text) + "%' order by id limit 0,30"
Combo4.Clear
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then Exit Sub
If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!id) Then
    Combo4.AddItem rtmp!id
  End If
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Private Sub Command1_Click()

'd2infile = "werkvz": d2insub = "Command1_Click"
isviz = False
Hide

End Sub

Public Sub Command10_Click()
Dim kid$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command10_Click"
If List2.ListIndex < 0 Then Exit Sub

kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub

kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
Load dochist2
DoEvents
Call dochist2.setkrit("((Werke: " + kid$, "")
On Error Resume Next
Call dochist2.SetFocus
On Error GoTo 0

End Sub

Private Sub Command11_Click()
Dim i%, id$

'd2infile = "werkvz": d2insub = "Command11_Click"
i% = List4.ListIndex
If i% < 1 Then Exit Sub

id$ = List4.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & i% & " where id='" & id$ & "'")
id$ = List4.List(i% - 1)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & i% + 1 & " where id='" & id$ & "'")
Call List2_Click
List4.ListIndex = i% - 1

End Sub

Private Sub Command12_Click()
Dim i%, id$
'd2infile = "werkvz": d2insub = "Command12_Click"
i% = List4.ListIndex
If i% >= List4.ListCount - 1 Or i% < 0 Then Exit Sub

id$ = List4.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & i% + 2 & " where id='" & id$ & "'")
id$ = List4.List(i% + 1)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & i% + 1 & " where id='" & id$ & "'")
Call List2_Click
List4.ListIndex = i% + 1

End Sub


Private Sub Command13_Click()
Dim V$, X

V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(form1.getkompidbywerkid(trm(Text4(0).text)))
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(0).text))
On Error Resume Next
MkDir V$
On Error GoTo 0
X = Shell("explorer.exe " & V$, vbNormalFocus)

End Sub

Private Sub Command14_Click()
Dim V$, X

'd2infile = "werkvz": d2insub = "Command14_Click"
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(trm(Text3(0)))
On Error Resume Next
MkDir V$
On Error GoTo 0
X = Shell("explorer.exe " & V$, vbNormalFocus)

End Sub

Private Sub Command15_Click()
Dim V$, X, o%, fn$, l$, l2apx As Integer, l2apy As Integer, bsl As Integer
Dim wrtn As Boolean

'd2infile = "werkvz": d2insub = "Command22_Click"
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
fn$ = strrepl(form1.myuniquedocname(""), " ", "-") + ".htm"
o% = FreeFile
l2apx = List2a.Left
l2apy = List2a.Top
List2a.Left = List2.Left
List2a.Top = List2.Top
List2a.Width = List2.Width
List2a.Height = List2.Height
List2.Visible = False
If Len(trm(fn$)) > 0 Then
  Open fn$ For Output As #o%
  List2a.Clear
  List2a.AddItem V$
  bsl = Len(V$)
  Do
    V$ = List2a.List(0): List2a.RemoveItem 0: Label12.Caption = trm(List2a.ListCount): DoEvents
    l$ = dirlist(V$ + "\")
    While Len(l$) > 0
      List2a.AddItem V$ + "\" + cut_d1(l$, "|")
'      List2a.ListIndex = List2a.ListCount - 1
      l$ = cut_d2bis(l$, "|")
    Wend
    l = Dir(V$ + "\*.*")
    wrtn = False
    While l$ <> ""
      Print #o%, "<a href=""" + V$ + "\" + l$ + """ target=_blank><small>" + Mid(V$, bsl + 2) + "\" + l$ + "</small></a><br>"
      l$ = Dir()
      wrtn = True
    Wend
    If wrtn Then Print #o%, "<br>"
  Loop Until List2a.ListCount = 0
  Close #o%
  'X = Shell("notepad.exe " & fn$, 1)
  X = Shell("explorer.exe " & fn$, 1)
End If
Label12.Caption = ""
List2a.Left = l2apx
List2a.Top = l2apy
List2.Visible = True
Call List1_Click
End Sub

Private Sub Command16_Click()
Dim r As ADODB.Recordset, rrr
Dim i%, fn$, o%, c$, l$, X

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command16_Click"
fn$ = strrepl(form1.myuniquedocname(""), " ", "-")
o% = FreeFile
If Len(trm(fn$)) > 0 Then
  Open fn$ For Output As #o%
  c$ = "SELECT sbz_loc.Satzbezeichnung as sbz, w_loc.Name as wbz, k_loc.Name as kname, k_loc.Vornamen as kvornamen, k_loc.Daten as kdaten " & _
       "FROM (sbz_loc INNER JOIN w_loc ON sbz_loc.wid = w_loc.id) INNER JOIN k_loc ON w_loc.KomponistenNummer = k_loc.id " & _
       "WHERE ((instr(sbz_loc.Satzbezeichnung,'Noten: ')=1)) order by k_loc.Name,w_loc.Name,sbz_loc.Satzbezeichnung;"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    l$ = trm(Mid(r!sbz, 8)) & ": " & r!kname & ", " & r!kvornamen & " (" & r!kdaten & ") - " & r!wbz
    Print #o%, l$
    r.MoveNext
  Wend
  Close #o%
  X = Shell("notepad.exe " & fn$, 1)
End If

End Sub

Private Sub Command17_Click()
Dim brw$, X, sa$
'd2infile = "werkvz": d2insub = "Command17_Click"
Unload frmBrowser
DoEvents
      
sa$ = "http://www.google.de/search?q=" & _
                googlrepl(Text3(2).text) & "+" & _
                googlrepl(Text3(1).text) & _
                "&ie=UTF-8&oe=UTF-8&hl=de&btnG=Google-Suche&meta="
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & sa$, 1)
Else
  frmBrowser.StartingAddress = sa$
  Load frmBrowser
End If
End Sub

Function googlrepl(l$) As String
Dim t$

'd2infile = "werkvz": d2insub = "googlrepl"
googlrepl = l$
t$ = strrepl(trm(l$), " ", "+")
t$ = strrepl(t$, "ä", "%C3%A4")
t$ = strrepl(t$, "ö", "%C3%B6")
t$ = strrepl(t$, "ü", "%C3%BC")
t$ = strrepl(t$, "Ä", "%C3%84")
t$ = strrepl(t$, "Ö", "%C3%96")
t$ = strrepl(t$, "Ü", "%C3%9C")
t$ = strrepl(t$, "ß", "%C3%9F")
t$ = strrepl(t$, "ò", "%C3%B2")

googlrepl = t$

End Function

Private Sub Command18_Click()
Dim brw$, X, sa$
'd2infile = "werkvz": d2insub = "Command17_Click"
Unload frmBrowser
DoEvents
      
sa$ = "http://www.google.de/search?q=" & _
                googlrepl(Text3(2).text) & "+" & _
                googlrepl(Text3(1).text) & "+" & _
                googlrepl(Text4(1).text) & _
                "&ie=UTF-8&oe=UTF-8&hl=de&btnG=Google-Suche&meta="
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & sa$, 1)
Else
  DoEvents
  frmBrowser.StartingAddress = sa$
  Load frmBrowser
End If
End Sub

Private Sub Command19_Click()
'd2infile = "werkvz": d2insub = "Command19_Click"
Call form1.handbuchcall("07-Werkeverzeichnis.htm")

End Sub

Public Sub Command2_Click()
Dim i%, up$, cmd$, rtmp As QueryDef, nflds As Integer, rrr
Dim stmp As ADODB.Recordset, id$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command2_Click"
nflds = 10

up$ = "insert into k_loc ("
For i% = 0 To nflds
  up$ = up$ + form1.sqla.TableDefs("k_loc").Fields(i%).name + ","
Next i%
up$ = Left$(up$, Len(up$) - 1) + ") values("
For i% = 0 To nflds
  Select Case i%
    Case 1: Text3(i%).text = "Komponist"
    Case 2: Text3(i%).text = "Neuer"
    Case 0: Do
              id$ = Left(trm(str$(Rnd)), 10)
              Set stmp = New ADODB.Recordset
              stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT id FROM k_loc where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            Loop Until stmp.EOF
            Text3(i%).text = id$
    Case Else: Text3(i%).text = ""
  End Select

  If Len(Text3(i%).text) = 0 Then
    up$ = up$ + "NULL,"
  Else
    up$ = up$ + "'" + strrepl(Text3(i%).text, "'", "´") + "',"
  End If
Next i%
up$ = Left$(up$, Len(up$) - 1)
cmd$ = up$ + ")"
Call form1.sqlqry(cmd$)

Call rlist1
Text1.text = "Komponist, Neuer"
DoEvents
'Text1.Text = ""
End Sub

Private Sub Command20_Click()
Dim i%, p%, a$, k2id$, k2n$, kid$, kn$, cmd$

'd2infile = "werkvz": d2insub = "Command20_Click"
For i% = 0 To List1.ListCount - 2
  p% = InStr(List1.List(i%), ".") - 1
  If p% > 5 Then
    a$ = Left(List1.List(i%), p%)
    If Left(List1.List(i% + 1), p%) = a$ Then
       List1.ListIndex = i% + 1
       DoEvents
       k2id$ = List1.List(List1.ListIndex)
       k2n$ = trm(Left$(k2id$, InStr(k2id$, "(ID:") - 1))
       k2id$ = Mid$(k2id$, InStr(k2id$, "(ID:") + 4)
       k2id$ = Left$(k2id$, InStr(k2id$, ")") - 1)
       List1.ListIndex = i%
       DoEvents
       kid$ = List1.List(List1.ListIndex)
       kn$ = trm(Left$(kid$, InStr(kid$, "(ID:") - 1))
       kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
       kid$ = Left$(kid$, InStr(kid$, ")") - 1)
       cmd$ = "update w_loc set Komponistennummer='" + k2id$ + "' where Komponistennummer='" + kid$ + "'"
       Call form1.sqlqry(cmd$)
       cmd$ = "delete from k_loc where id='" + kid$ + "'"
       Call form1.sqlqry(cmd$)
    End If
  End If
Next i%
End Sub

Private Sub Command21_Click()
Dim V$, X

'd2infile = "werkvz": d2insub = "Command21_Click"
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(trm(Text3(1) & " " & Text3(2)))
On Error Resume Next
MkDir V$
On Error GoTo 0
X = Shell("explorer.exe " & V$, vbNormalFocus)
End Sub

Private Sub Command22_Click()
Dim V$, X, kn$

'd2infile = "werkvz": d2insub = "Command22_Click"
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
On Error Resume Next
MkDir V$
On Error GoTo 0
kn$ = form1.getkompnamebyid(form1.getkompidbywerkid(trm(Text4(0).text)))
V$ = V$ & "\" & form1.mkkompdn(strrepl(kn$, ",", ""))
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(1)))
On Error Resume Next
MkDir V$
On Error GoTo 0
X = Shell("explorer.exe " & V$, vbNormalFocus)

End Sub

Private Sub Command23_Click()
Dim i%, id$
'd2infile = "werkvz": d2insub = "Command23_Click"
i% = List4.ListIndex
If i% < 1 Then Exit Sub

id$ = List4.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & 0 & " where id='" & id$ & "'")
Call List2_Click
List4.ListIndex = 0

End Sub

Private Sub Command24_Click()
Dim i%, id$
'd2infile = "werkvz": d2insub = "Command24_Click"
i% = List4.ListIndex
If i% < 0 Then Exit Sub

id$ = List4.List(i%)
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
id$ = Left$(id$, Len(id$) - 1)
form1.sqlqry ("update sbz_loc set satznummer=" & List4.ListCount * 10 & " where id='" & id$ & "'")
Call List2_Click
List4.ListIndex = List4.ListCount - 1

End Sub

Private Sub Command25_Click()
Dim V$, o%, fn$, X, rrr

'd2infile = "werkvz": d2insub = "Command25_Click"
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(form1.getkompnamebyid(form1.getkompidbywerkid(trm(Text4(0).text))))
On Error Resume Next
MkDir V$
On Error GoTo 0
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(1)))
On Error Resume Next
MkDir V$
On Error GoTo 0
o% = FreeFile
fn$ = V$ & "\pd-bib.txt"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, Text4(17).text
  Close #o%
  X = Shell("notepad.exe " & fn$, vbNormalFocus)
Else
  MsgBox transe("Die Datei kann nicht geschrieben werden:") & vbCrLf & fn$
End If
End Sub

Private Sub Command26_Click()
Dim wid$, kid$, wert$, r As ADODB.Recordset, i As Integer, c$, rrr, wn$

If List2.ListIndex < 0 Then Exit Sub

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub

wn$ = trm(Left$(wid$, InStr(wid$, "(WID:") - 1))
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
wert$ = trm(InputBox(transe("Werk verschieben:") + vbCrLf + wn$, transe("Werk zu einem anderen Komponisten übertragen"), wert$))
If wert$ = "" Then Exit Sub
c$ = "SELECT id FROM k_loc where id='" & wert$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If r.EOF Then
  MsgBox transe("Ein Komponist mit dieser ID existiert nicht") + "."
  Exit Sub
Else
  c$ = "update w_loc set KomponistenNummer='" + wert$ + "' where id='" + wid$ + "'"
  Call form1.sqlqry(c$)
End If
kid$ = form1.getkompnamebywerkid(wid$)
Call showkompdetailbyname(kid$)
Call Timer2_Timer
'Call werkvz.showwerkdetail(wid$) war hier nich sooo gut. so bessa
c$ = form1.getwerknamebyid(wid$)
For i% = 0 To werkvz.List2.ListCount - 1
  If InStr(List2.List(i%), c$) = 1 Then
    List2.ListIndex = i%
    Exit Sub
  End If
Next i%

End Sub

Private Sub Command27_Click()
'd2infile = "werkvz": d2insub = "Command27_Click"
List1.ListIndex = -1
End Sub

Private Sub Command28_Click()
Dim i%, xd$, kid$, kdir$, o%, fn$, xdw$, f%, V$, l$, kl$, g%, f1%, v1$, j%
Dim uckl$, suchw$, xdk$, j3%, sz$, szp%, wid$, kompchg%, werkchg%, nkomp%, nkompv%, ci%
Dim wk$, wkf%, v2$, f2%, j1%, kompname$, wrkf%, f3%, wkchg$, wkfchg$, c$, zgvon$, zgbis$
Dim rtmp As ADODB.Recordset, rrtmp As ADODB.Recordset, kompcount%, werkcount%, X
Dim immersuch$, d7z As String


Dim d2infile As String, d2insub As String
MousePointer = 11: DoEvents
d2infile = "werkvz": d2insub = "Command28_Click"
kompcount% = 0: werkcount% = 0
xd$ = form1.getusersetting("htmlwerke", "")
If xd$ = "" Then Exit Sub
xdw$ = xd$ & "\werke"
On Error Resume Next
MkDir xdw$
On Error GoTo 0

If 0 = 1 Then
immersuch$ = "Künstleragentur, Software, Veranstalter, Konzertveranstalter, Branchenlösung, Datenbank, Datenverarbeitung, Agentur, Veranstaltungsagentur, Datenbank, "
V$ = form1.vorlagenverzeichnis() + "\kompabc.htm"
f% = FreeFile
Open V$ For Input As #f%
o% = FreeFile
Open xdw$ & "\kompnews.htm" For Output As #o%
While Not EOF(f%)
  Line Input #f%, l$
  If l$ = "lobgesang" Then l$ = lobgesang()
  If l$ = "<bkmk>index</bkmk>" Then
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open "SELECT * FROM k_loc order by stand desc;", form1.adoc, adOpenDynamic, adLockReadOnly
    i% = 0
    Print #o%, "<table border=0>"
    Print #o%, "<tr><td><b>Komponist:</b></td><td>geändert am:</td></tr>"
    While i% < 30 And Not rtmp.EOF
      Print #o%, "<tr><td><a href=komp" & trm(rtmp!id) & "/index.htm>" & trm(rtmp!name) & ", " & trm(rtmp!vornamen) & "</a> </td><td>" & datfromsql(rtmp!stand) & "</td></tr>"
      i% = i% + 1
      rtmp.MoveNext
    Wend
    rtmp.Close
    Print #o%, "</table>"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rtmp.Open "SELECT * FROM w_loc order by stand desc;", form1.adoc, adOpenDynamic, adLockReadOnly
    i% = 0
    Print #o%, "<table border=0>"
    Print #o%, "<tr><td><b>Werk:</b></td><td>geändert am:</td></tr>"
    While i% < 30 And Not rtmp.EOF
      Print #o%, "<tr><td><a href=komp" & form1.getkompidbywerkid(rtmp!id) & "/werk" & rtmp!id & ".htm><small>" & form1.getkompnamebywerkid(rtmp!id) & ": " & trm(rtmp!name) & "</small></a> </td><td><small>" & datfromsql(rtmp!stand) & "</small></td></tr>"
      i% = i% + 1
      rtmp.MoveNext
    Wend
    rtmp.Close
    Print #o%, "</table>"
  Else
    If InStr(l$, "<meta name=""date"" content=") = 1 Then
      Print #o%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
    Else
      Print #o%, l$
    End If
  End If
Wend
Close #f%
Close #o%

V$ = form1.vorlagenverzeichnis() + "\kompabc.htm"
f% = FreeFile
Open V$ For Input As #f%
o% = FreeFile
Open xdw$ & "\kompabc.htm" For Output As #o%
While Not EOF(f%)
  Line Input #f%, l$
  If l$ = "lobgesang" Then
    l$ = lobgesang()
    l$ = "<p><span class=""boldblue"">Die Datenbank enthält " & trm(kompcount%) & " Komponisten und " & trm(werkcount%) & " Werke.</span></p>" & l$
  End If
  If l$ = "<bkmk>index</bkmk>" Then
    'neukomp
    nkompv% = FreeFile
    Open form1.vorlagenverzeichnis() + "\kompa.htm" For Input As #nkompv%
    nkomp% = FreeFile
    Open xdw$ & "\neukomp.htm" For Output As #nkomp%
    While Not EOF(nkompv%)
      Line Input #nkompv%, l$
      If l$ = "lobgesang" Then l$ = lobgesang()
      If l$ = "<bkmk>index</bkmk>" Then
        Print #nkomp%, "<form name=kchg action=/cgi-bin/chgdata><input type=hidden name=id value=""NEU""><input type=hidden name=chgtyp value=komponist>"
        Print #nkomp%, "<table><tr><td>Name:</td><td><input type=text name=name></td></tr>"
        Print #nkomp%, "<tr><td>Vorname:</td><td><input type=text name=vorname></td></tr>"
        Print #nkomp%, "<tr><td>Von:</td><td><input type=text name=von></td></tr>"
        Print #nkomp%, "<tr><td>Bis:</td><td><input type=text name=bis></td></tr>"
        Print #nkomp%, "<tr><td><hr></td><td><hr></td></tr>"
        Print #nkomp%, "<tr><td>Herkunft der Daten:<br><small>(Verlag, Noten, CD, ...)<small></td><td valign=top><input type=text name=quelle></td></tr>"
        Print #nkomp%, "<tr><td>Ihre Emailadresse:</td><td><input type=text name=srcmail></td></tr>"
        Print #nkomp%, "<tr><td><input type=reset></td><td><input type=submit name=submit value=""Daten absenden""></td></tr>"
        Print #nkomp%, "</table>"
        Print #nkomp%, "</form>"
      Else
        If InStr(l$, "<meta name=""date"" content=") = 1 Then
          Print #nkomp%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
        Else
          Print #nkomp%, l$
        End If
      End If
    Wend
    Close #nkompv%
    Close #nkomp%
    'EOneukomp
    'neuwerk
    nkompv% = FreeFile
    Open form1.vorlagenverzeichnis() + "\kompa.htm" For Input As #nkompv%
    nkomp% = FreeFile
    Open xdw$ & "\neuwerk.htm" For Output As #nkomp%
    While Not EOF(nkompv%)
      Line Input #nkompv%, l$
      If l$ = "lobgesang" Then l$ = lobgesang()
      If l$ = "<bkmk>index</bkmk>" Then
                  Print #nkomp%, "<p><span class=""headblue"">Neues Werk</span>&nbsp;</p>"
                  Print #nkomp%, "<form name=wchg action=/cgi-bin/chgdata>"
                  Print #nkomp%, "<input type=hidden name=id value=""NEU"">"
                  Print #nkomp%, "<input type=hidden name=chgtyp value=werk>"
                  Print #nkomp%, "<input type=hidden name=kompid value=""" & kid$ & """>"
                  Print #nkomp%, "<input type=hidden name=wid value=""" & Text4(0).text & """>"
                  Print #nkomp%, "<table><tr><td>Komponist:</td><td><input type=text name=kompname size=60></td></tr>"
                  Print #nkomp%, "<tr><td>Name:</td><td><input type=text name=name size=60></td></tr>"
                  Print #nkomp%, "<tr><td>Dauer ca.:</td><td><input type=text name=dauer size=6> Minuten</td></tr>"
                  Print #nkomp%, "<tr><td>Opusjahr:</td><td><input type=text size=6></td></tr>"
                  Print #nkomp%, "<tr><td>Opusjahr von:</td><td><input type=text size=6></td></tr>"
                  Print #nkomp%, "<tr><td>Opusjahr bis:</td><td><input type=text size=6></td></tr>"
                  Print #nkomp%, "<tr><td valign=top>Satzangaben:</td><td><textarea name=saetze cols=50 rows=5>"
                  Print #nkomp%, "</textarea></td></tr>"
                  Print #nkomp%, "<tr><td><hr></td><td><hr></td></tr>"
                  Print #nkomp%, "<tr><td>Herkunft der Daten:<br><small>(Verlag, Noten, CD, ...)<small></td><td valign=top><input type=text name=quelle></td></tr>"
                  Print #nkomp%, "<tr><td>Ihre Emailadresse:</td><td><input type=text name=srcmail></td></tr>"
                  Print #nkomp%, "<tr><td><input type=reset></td><td><input type=submit name=submit value=""Daten absenden""></td></tr>"
                  Print #nkomp%, "</table>"
                  Print #nkomp%, "</form>"
      Else
        If InStr(l$, "<meta name=""date"" content=") = 1 Then
          Print #nkomp%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
        Else
          Print #nkomp%, l$
        End If
      End If
    Wend
    Close #nkompv%
    Close #nkomp%
    'EOneuwerk
    Print #o%, "<p><a href=kompnews.htm>Neueste Änderungen</a></p>"
    For i% = 0 To 25
      kl$ = Chr$(Asc("a") + i%): uckl$ = UCase(kl$)
      v1$ = form1.vorlagenverzeichnis() + "\kompa.htm"
      f1% = FreeFile
      Open v1$ For Input As #f1%
      g% = FreeFile
      Open xdw$ & "/komp" & kl$ & ".htm" For Output As #g%
      While Not EOF(f1%)
        Line Input #f1%, l$
        If l$ = "lobgesang" Then l$ = lobgesang()
        If l$ = "<bkmk>index</bkmk>" Then
          Print #g%, "<p><a href=neukomp.htm><small>Neuer Komponist</small></a></p><p><a href=neuwerk.htm><small>Neues Werk</small></a></p><table border=0>"
          kompcount% = List1.ListCount
          For j% = 0 To List1.ListCount - 1
            If Left$(List1.List(j%), 1) = uckl$ And Text3(1).text <> "Pause" And Text3(1).text <> "oder" And Text3(2).text <> "Pause" And Text3(2).text <> "oder" Then
              List1.ListIndex = j%: DoEvents
              kid$ = Text3(0).text
              kdir$ = xdw$ & "\komp" & kid$
              On Error Resume Next
              MkDir kdir$
              On Error GoTo 0
              Print #g%, "<tr><td><p><a href=komp" & Text3(0).text & "/index.htm>" & Text3(1).text & ", " & Text3(2).text & "</a> " & "</p></td><td><p>(" & Text3(4).text & "-" & Text3(5).text & ")</p></td><td> &nbsp; </td><td><p>" & trm(List2.ListCount) & " Werke</p></td>"
'              Print #g%, "<td> &nbsp; </td><td><a href=""http://vdkw.dnsalias.net:5080/cgi-bin/noten/afunc.tcl?act=zw&wid=CHGK:" & kid$ & """>Daten ändern</a></td></tr>"
              Print #g%, "<td> &nbsp; </td><td> </td></tr>"
              v2$ = form1.vorlagenverzeichnis() + "\komp.htm"
              f2% = FreeFile
              Open v2$ For Input As #f2%
              wk$ = kdir$ & "\index.htm"
              wkf% = FreeFile
              Open wk$ For Output As #wkf%
              wkchg$ = kdir$ & "\chgkomp.htm"
              kompchg% = FreeFile
              Open wkchg$ For Output As #kompchg%
              kompname$ = Text3(2).text & " " & Text3(1).text
              zgvon$ = trm(Text3(4).text)
              zgbis$ = (Text3(5).text)
              If zgvon$ = "" Then zgvon$ = trm(apyear(Date) + 1)
              If zgbis$ = "" Then zgbis$ = trm(apyear(Date))
              While Not EOF(f2%)
                Line Input #f2%, l$
                If l$ = "lobgesang" Then l$ = lobgesang()
                If l$ = "<bkmk>index</bkmk>" Then
                  Print #kompchg%, "<form name=kchg action=/cgi-bin/chgdata><input type=hidden name=id value=" & kid$ & "><input type=hidden name=chgtyp value=komponist>"
                  Print #kompchg%, "<table><tr><td>Name:</td><td><input type=text name=name value=""" & Text3(1).text & """></td></tr>"
                  Print #kompchg%, "<tr><td>Vorname:</td><td><input type=text name=vorname value=""" & Text3(2).text & """></td></tr>"
                  Print #kompchg%, "<tr><td>Von:</td><td><input type=text name=von value=""" & Text3(4).text & """></td></tr>"
                  Print #kompchg%, "<tr><td>Bis:</td><td><input type=text name=bis value=""" & Text3(5).text & """></td></tr>"
                  Print #kompchg%, "<tr><td><hr></td><td><hr></td></tr>"
                  Print #kompchg%, "<tr><td>Herkunft der Daten:<br><small>(Verlag, Noten, CD, ...)<small></td><td valign=top><input type=text name=quelle></td></tr>"
                  Print #kompchg%, "<tr><td>Ihre Emailadresse:</td><td><input type=text name=srcmail></td></tr>"
                  Print #kompchg%, "<tr><td><input type=reset></td><td><input type=submit name=submit value=""Daten absenden""></td></tr>"
                  Print #kompchg%, "</table>"
                  Print #kompchg%, "</form>"
                  Print #wkf%, "<p><span class=""headblue"">" & kompname$ & " (" & Text3(3).text & ")</span> &nbsp; </p>"
                  Print #wkf%, "<p>"
                  Print #wkf%, "<small>erstellt: " & Date & "</small> "
                  If trm(Text3(6).text) <> "" Then
                    Print #wkf%, "<small>- Stand: " & datfromsql(Text3(6).text) & "</small> "
                  End If
                  Print #wkf%, "<a href=chgkomp.htm><small>Daten ändern</small></a>, <a href=neuwerk.htm><small>Neues Werk</small></a></p>"
                  Print #wkf%, "<p> &nbsp; </p>"

                    f3% = FreeFile
                    Open v2$ For Input As #f3%
                    werkchg% = FreeFile
                    Open kdir$ & "\neuwerk.htm" For Output As werkchg%
                    While Not EOF(f3%)
                      Line Input #f3%, l$
                      If l$ = "lobgesang" Then l$ = lobgesang()
                      If l$ = "<bkmk>index</bkmk>" Then
                  Print #werkchg%, "<p><span class=""headblue""><a href=index.htm>" & kompname$ & "</a> (" & Text3(3).text & ")</span>&nbsp;</p>"
                  Print #werkchg%, "<form name=wchg action=/cgi-bin/chgdata>"
                  Print #werkchg%, "<input type=hidden name=id value=""NEU"">"
                  Print #werkchg%, "<input type=hidden name=chgtyp value=werk>"
                  Print #werkchg%, "<input type=hidden name=kompname value=""" & kompname$ & """>"
                  Print #werkchg%, "<input type=hidden name=kompid value=""" & kid$ & """>"
                  Print #werkchg%, "<input type=hidden name=wid value=""NEU"">"
                  Print #werkchg%, "<table><tr><td>Name:</td><td><input type=text name=name size=60></td></tr>"
                  Print #werkchg%, "<tr><td>Dauer ca.:</td><td><input type=text name=dauer size=6> Minuten</td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr:</td><td><input type=text size=6 name=op></td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr von:</td><td><input type=text size=6></td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr bis:</td><td><input type=text size=6></td></tr>"
                  Print #werkchg%, "<tr><td valign=top>Satzangaben:</td><td><textarea name=saetze cols=50 rows=5>"
                  Print #werkchg%, "</textarea></td></tr>"
                  Print #werkchg%, "<tr><td><hr></td><td><hr></td></tr>"
                  Print #werkchg%, "<tr><td>Herkunft der Daten:<br><small>(Verlag, Noten, CD, ...)<small></td><td valign=top><input type=text name=quelle></td></tr>"
                  Print #werkchg%, "<tr><td>Ihre Emailadresse:</td><td><input type=text name=srcmail></td></tr>"
                  Print #werkchg%, "<tr><td><input type=reset></td><td><input type=submit name=submit value=""Daten absenden""></td></tr>"
                  Print #werkchg%, "</table>"
                  Print #werkchg%, "</form>"
                      Else
                        If InStr(l$, "<meta name=""date"" content=") = 1 Then
                          Print #werkchg%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
                        Else
                          Print #werkchg%, l$
                        End If
                      End If
                    Wend
                    Close #f3%
                    Close #werkchg%

                  For j1% = 0 To List2.ListCount - 1
                    List2.ListIndex = j1%: DoEvents
                  If InStr(LCase(trm(Text4(1).text)), "auswahl") = 0 And Left(trm(Text4(1).text), 1) <> "?" Then
                    Print #wkf%, "<p><span class=""boldblue""><a href=werk" & Text4(0).text & ".htm>" & Text4(1).text & "</a></span>&nbsp;</p>"


'*********************Hier das Werk*****************************

                    f3% = FreeFile
                    werkcount% = werkcount% + 1
                    Open v2$ For Input As #f3%
                    wrkf% = FreeFile
                    Open kdir$ & "\werk" & Text4(0).text & ".htm" For Output As wrkf%
                    werkchg% = FreeFile
                    Open kdir$ & "\chgwerk" & Text4(0).text & ".htm" For Output As werkchg%
                    While Not EOF(f3%)
                        Line Input #f3%, l$
                        If l$ = "lobgesang" Then l$ = lobgesang()
                        If l$ = "<bkmk>index</bkmk>" Then
                  Print #werkchg%, "<p><span class=""headblue""><a href=index.htm>" & kompname$ & "</a> (" & Text3(3).text & ")</span>&nbsp;</p>"
                  Print #werkchg%, "<form name=wchg action=/cgi-bin/chgdata>"
                  Print #werkchg%, "<input type=hidden name=id value=" & Text4(0).text & ">"
                  Print #werkchg%, "<input type=hidden name=chgtyp value=werk>"
                  Print #werkchg%, "<input type=hidden name=kompname value=""" & kompname$ & """>"
                  Print #werkchg%, "<input type=hidden name=kompid value=""" & kid$ & """>"
                  Print #werkchg%, "<input type=hidden name=wid value=""" & Text4(0).text & """>"
                  Print #werkchg%, "<table><tr><td>Name:</td><td><input type=text name=name size=60 value=""" & Text4(1).text & """></td></tr>"
                  Print #werkchg%, "<tr><td>Dauer ca.:</td><td><input type=text name=dauer size=6 value=""" & Text4(5).text & """> Minuten</td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr:</td><td><input type=text size=6 name=op value=""" & Text4(6).text & """></td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr von:</td><td><input type=text size=6 name=opvon value=""" & Text4(7).text & """></td></tr>"
                  Print #werkchg%, "<tr><td>Opusjahr bis:</td><td><input type=text size=6 name=opbis value=""" & Text4(8).text & """></td></tr>"
                  Print #werkchg%, "<tr><td valign=top>Satzangaben:</td><td><textarea name=saetze cols=50 rows=5>"
                  For j3% = 0 To List4.ListCount - 1
                    sz$ = trm(List4.List(j3%))
                    If InStr(LCase(sz$), "noten:") = 0 Then
                        szp% = InStr(sz$, "(ID:")
                        If szp% > 0 Then sz$ = trm(Left(sz$, szp% - 1))
                        Print #werkchg%, sz$
                    End If
                  Next j3%
                  Print #werkchg%, "</textarea></td></tr>"
                  Print #werkchg%, "<tr><td><hr></td><td><hr></td></tr>"
                  Print #werkchg%, "<tr><td>Herkunft der Daten:<br><small>(Verlag, Noten, CD, ...)<small></td><td valign=top><input type=text name=quelle></td></tr>"
                  Print #werkchg%, "<tr><td>Ihre Emailadresse:</td><td><input type=text name=srcmail></td></tr>"
                  Print #werkchg%, "<tr><td><input type=reset></td><td><input type=submit name=submit value=""Daten absenden""></td></tr>"
                  Print #werkchg%, "</table>"
                  Print #werkchg%, "</form>"
                            Print #wrkf%, "<p><span class=""headblue""><a href=index.htm>" & kompname$ & "</a> (" & Text3(3).text & ")</span>&nbsp;</p>"
                            Print #wrkf%, "<p>&nbsp;</p>"
                            Print #wrkf%, "<p><span class=""headblue"">" & Text4(1).text & "&nbsp; " & opjahr("(", trm(Text4(6).text), trm(Text4(7).text), trm(Text4(8).text), ")") & "</span></p>"
'                            If (trm(Text4(6).Text) <> "" And trm(Text4(6).Text) <> "0") _
'                                Or (trm(Text4(7).Text) <> "" And trm(Text4(7).Text) <> "0") _
'                                Or (trm(Text4(8).Text) <> "" And trm(Text4(8).Text) <> "0") Then
'                              Print #wrkf%, "<p>Opusjahr: " & Text4(6).Text & "&nbsp;"
'                              If trm(Text4(7).Text) <> "" Then
'                                Print #wrkf%, "von " & Text4(7).Text & "&nbsp;"
'                              End If
'                              If trm(Text4(8).Text) <> "" Then
'                                Print #wrkf%, "bis " & Text4(8).Text & "&nbsp;"
'                              End If
'                              Print #wrkf%, "</p>"
                            'End If
                            If trm(Text4(5).text) <> "" And trm(Text4(5).text) <> "0" Then
                              Print #wrkf%, "<p>Dauer ca. " & Text4(5).text & "'&nbsp;</p>"
                            End If
                            Print #wrkf%, "<p> &nbsp;</p>"
                            For j3% = 0 To List4.ListCount - 1
                              sz$ = trm(List4.List(j3%))
                              If InStr(LCase(sz$), "noten:") = 0 Then
                                szp% = InStr(sz$, "(ID:")
                                If szp% > 0 Then sz$ = trm(Left(sz$, szp% - 1))
                                Print #wrkf%, "<p><span class=""boldblue"">" & sz$ & "</span>&nbsp;</p>"
                              End If
                            Next j3%
                            wid$ = List2.List(List2.ListIndex)
                            If InStr(wid$, "(WID:") > 0 Then
                              wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)
                              Print #wrkf%, "<p>"
                            End If
                            Print #wrkf%, "<p>"
                            Print #wrkf%, "<p><?php"
                            Print #wrkf%, "include_once(""/home/www/web3/html/a.php"");"
                            Print #wrkf%, "werknoten(""" + kid$ + """,""" + wid$ + """);"
                            Print #wrkf%, "?></p>"
                            Print #wrkf%, "<p></p>"
                            Print #wrkf%, "<small>erstellt: " & Date & "</small> "
                            If trm(Text4(10).text) <> "" Then
                              Print #wrkf%, "<small>- Stand: " & datfromsql(Text4(10).text) & "</small> "
                            End If
                            Print #wrkf%, "<a href=chgwerk" & Text4(0).text & ".htm><small>Daten ändern</small></a></p>"
'                            Print #wrkf%, "Für die Änderung von Daten benötigen Sie einen Benutzernamen und ein Passwort.<br>"
'                            Print #wrkf%, "<a href=""http://vdkw.dnsalias.net:5080/cgi-bin/nfunc.tcl?NEWUID"">Anmeldung (oder Passwort vergessen) hier.</a>"
                        Else
                            If l$ = "<bkmk>suchw</bkmk>" Then
                                suchw$ = immersuch$ + ", " + kompname$ & ", Satzangaben, Sätze, " & Text4(1).text
                                Print #wrkf%, suchw$
                            Else
                                If InStr(l$, "<meta name=""date"" content=") = 1 Then
                                  Print #wrkf%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
                                Else
                                  Print #wrkf%, l$
                                  Print #werkchg%, l$
                                End If
                            End If
                        End If
                    Wend
                    Print
                    Close #werkchg%
                    Close #wrkf%
                    Close #f3%

'*********************Hier endet das Werk*****************************


                  Else
                    Debug.Print trm(Text4(1).text) + " nicht exportiert."
                  End If
                  Next j1%
                  Print #wkf%, "<p><?php"
                  Print #wkf%, "include_once(""/home/www/web3/html/a.php"");"
                  Print #wkf%, "kompnoten(""" + kid$ + """);"
                  Print #wkf%, "?></p>"
                  Print #wkf%, "<p><hr></p>"
                Else
                  If l$ = "<bkmk>suchw</bkmk>" Then
                    suchw$ = immersuch$ + ", " + kompname$
                    For j1% = 0 To List2.ListCount - 1
                        List2.ListIndex = j1%: DoEvents
                        If InStr(LCase(suchw$), LCase(Text4(1).text)) = 0 Then
                          suchw$ = suchw$ & ", "
                          suchw$ = suchw$ & Text4(1).text
                        End If
                    Next j1%
                    Print #wkf%, suchw$
                  Else
                    If InStr(l$, "<meta name=""date"" content=") = 1 Then
                      Print #wkf%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
                    Else
                      Print #wkf%, l$
                      Print #kompchg%, l$
                    End If
                  End If
                End If
              Wend

              Close #wkf%
              Close #kompchg%
              Close #f2%
            End If
          Next j%
          Print #g%, "</table>"
           Print #g%, "<p><br>"
        Else
          If l$ = "<bkmk>suchw</bkmk>" Then
            suchw$ = immersuch$ + ", "
            For j% = 0 To List1.ListCount - 1
              If Left$(List1.List(j%), 1) = uckl$ Then
                List1.ListIndex = j%: DoEvents
                If InStr(LCase(suchw$), LCase(Text3(1).text)) = 0 Then
                  If suchw$ <> "" Then suchw$ = suchw$ & ", "
                  suchw$ = suchw$ & Text3(1).text
                End If
                If InStr(LCase(suchw$), LCase(Text3(2).text)) = 0 Then
                  If suchw$ <> "" Then suchw$ = suchw$ & ", "
                  suchw$ = suchw$ & Text3(2).text
                End If
              End If
            Next j%
            Print #g%, suchw$
          Else
            If InStr(l$, "<meta name=""date"" content=") = 1 Then
              Print #g%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
            Else
              Print #g%, l$
            End If
          End If
        End If
      Wend
      Close #g%
      Close #f1%
      Print #o%, "<a href=komp" & kl$ & ".htm><img src=../bilder/lbl-" & UCase(kl$) & "-0.jpg></a> &nbsp; "
      If (i% + 1) Mod 6 = 0 Then Print #o%, "<br>"
    Next i%
'    Print #o%, "<p><hr>Für die Änderung von Daten benötigen Sie einen Benutzernamen und ein Passwort.<br>"
    Print #o%, "<p><br>"
    Print #o%, "<p><a href=neukomp.htm><small>Neuer Komponist</small></a></p><p><a href=neuwerk.htm><small>Neues Werk</small></a></p>"

'    Print #o%, "<a href=""http://vdkw.dnsalias.net:5080/cgi-bin/nfunc.tcl?NEWUID""><b>Anmeldung (oder Passwort vergessen) hier.</b></a>"
  Else
    If InStr(l$, "<meta name=""date"" content=") = 1 Then
      Print #o%, "<meta name=""date"" content=""" & datum2sql(Date) & "T" & Time & "+00:00"">"
    Else
      Print #o%, l$
    End If
  End If
Wend
Close #o%
Close #f%

End If '0=1

List2.Clear
fn$ = xdw$ & "\xp_k_loc.txt": List2.AddItem "Komponisten werden exportiert: " & fn$: List2.ListIndex = List2.ListCount - 1: DoEvents
On Error Resume Next
Kill fn$
On Error GoTo 0
Call form1.pg_xp("k_loc", fn$)
fn$ = xdw$ & "\xp_w_loc.txt": List2.AddItem "Werke werden exportiert: " & fn$: List2.ListIndex = List2.ListCount - 1: DoEvents
On Error Resume Next
Kill fn$
On Error GoTo 0
Call form1.pg_xp("w_loc", fn$)
fn$ = xdw$ & "\xp_sbz_loc.txt": List2.AddItem "Satzbezeichnungen werden exportiert: " & fn$: List2.ListIndex = List2.ListCount - 1: DoEvents
On Error Resume Next
Kill fn$
On Error GoTo 0
Call form1.pg_xp("sbz_loc", fn$)
MousePointer = 0: DoEvents
X = Shell("explorer.exe " & xdw$, vbNormalFocus)
d7z = form1.s0dir()
If Not nexist("c:\program files\7-zip\7z.exe") Then d7z = "c:\program files\7-zip"
If Not nexist(d7z & "\7z.exe") Then
  X = Shell(d7z & "\7z.exe a " & xd$ & "\werke.zip " & xdw$, vbNormalFocus)
End If
End Sub

Private Sub Command29_Click()
Dim up$, cmd$, i%, nflds As Integer, kid$, id$, rrr
Dim rtmp As ADODB.Recordset, nid$, l$, p%

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command29_Click"
nflds = 17
kid$ = Text3(0).text
If kid$ = "" Then Exit Sub
up$ = "insert into w_loc ("
For i% = 0 To nflds
  up$ = up$ + form1.sqla.TableDefs("w_loc").Fields(i%).name + ","
Next i%
up$ = Left$(up$, Len(up$) - 1) + ") values("
For i% = 0 To nflds
  If i% = 0 Then
    id$ = form1.newid("w_loc", "id", 10)
    Text4(i%).text = id$
  End If

  If Len(Text4(i%).text) = 0 Then
    up$ = up$ + "NULL,"
  Else
    up$ = up$ + "'" + Text4(i%).text + "',"
  End If
Next i%
up$ = Left$(up$, Len(up$) - 1)
cmd$ = up$ + ")"
Call form1.sqlqry(cmd$)
cmd$ = "update w_loc set bemerkung='" & Text4(17).text & "' where id ='" & id$ & "'"
Call form1.sqlqry(cmd$)
up$ = "insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('"
For i% = 0 To List4.ListCount - 1
  l$ = List4.List(i%)
  p% = InStr(l$, " (ID:")
  l$ = Mid$(l$, p% + 5): l$ = Left$(l$, Len(l$) - 1)
  cmd$ = "select * from sbz_loc where id='" + l$ + "'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    nid$ = form1.newid("sbz_loc", "id", 20)
    cmd$ = up$ & nid$ & "','" & id$ & "','" & rtmp!satzbezeichnung & "'," & rtmp!SatzNummer & ")"
    Call form1.sqlqry(cmd$)
    rtmp.MoveNext
  End If
Next i%

End Sub

Public Sub Command3_Click()
Dim i%, up$, cmd$, rtmp1 As ADODB.Recordset, rtmp As QueryDef, rrr
Dim nflds, id$, n$, V$, erg, beidenull

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command3_Click"
id$ = Text3(0).text
If id$ = "" Then Exit Sub
nflds = 10

Set rtmp1 = New ADODB.Recordset
rtmp1.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp1, "SELECT * FROM k_loc where id ='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

n$ = strrepl(Text3(1).text, "'", "´")
V$ = strrepl(Text3(2).text, "'", "´")
If List1.ListIndex >= 0 And InStr(List1.List(List1.ListIndex), "Komponist, Neuer ") = 1 Then List1.RemoveItem List1.ListIndex
For i% = 1 To nflds
  erg = (trm(rtmp1.Fields(i%)) <> trm(Text3(i%).text))
  beidenull = (IsNull(rtmp1.Fields(i%)) And Len(Text3(i%).text) = 0)
  If Not beidenull And (IsNull(erg) Or erg) Then
    If Len(Text3(i%).text) = 0 Then
      up$ = transo(Label1(i%).Caption) + "=NULL"
    Else
      up$ = transo(Label1(i%).Caption) + "= '" + strrepl(Text3(i%).text, "'", "´") + "'"
    End If
    cmd$ = "update k_loc set " + up$ + " where id= '" + id$ + "'"
    Call form1.sqlqry(cmd$)
  End If
Next i%
cmd$ = "update k_loc set stand='" + datum2sql(Date) + "' where id= '" + id$ + "'": Call form1.sqlqry(cmd$)
Text1.text = " ": DoEvents
Text1.text = n$ & ", " & V$
DoEvents
If List1.ListIndex = -1 Then
  List1.AddItem "" & n$ & ", " & V$ & Space$(80) & "  (ID:" + Text3(0).text + ")"
  DoEvents
  Call Text1_Change
End If
Call showkompdetail(id$)
Call showwerkdetail("-1")
BackColor = form1.cleancolor()
End Sub

Private Sub Command30_Click()
Dim fn$, X, d0$

'd2infile = "werkvz": d2insub = "Command30_Click"
MousePointer = 11: DoEvents
List2.Clear
d0$ = form1.s0dir()
fn$ = d0$ & "\xp_k_loc.txt": List2.AddItem "Komponisten werden exportiert: " & fn$: DoEvents
Call form1.pg_xp("k_loc", fn$)
fn$ = d0$ & "\xp_w_loc.txt": List2.AddItem "Werke werden exportiert: " & fn$: DoEvents
Call form1.pg_xp("w_loc", fn$)
fn$ = d0$ & "\xp_sbz_loc.txt": List2.AddItem "Satzbezeichnungen werden exportiert: " & fn$: DoEvents
Call form1.pg_xp("sbz_loc", fn$)
List2.AddItem "fertig"
X = Shell("explorer.exe " & d0$, vbNormalFocus)
MousePointer = 0
End Sub

Private Sub Command31_Click()
Dim kid$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command31_Click"
If List2.ListIndex < 0 Then Exit Sub

kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub

kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
Load dochist2
Call dochist2.setkrit("((Repert: " + kid$, "")
On Error Resume Next
Call dochist2.SetFocus
On Error GoTo 0

End Sub

Private Sub Command33_Click()
Dim i%, j%, k$, w$

'd2infile = "werkvz": d2insub = "Command33_Click"
Clipboard.Clear
j% = List2.ListIndex
If j% < 0 Then
  Command33.Enabled = False
  Exit Sub
End If
w$ = List2.List(j%)
i% = InStr(w$, "(KID:"): If i% > 0 Then w$ = trm(Left$(w$, i% - 1))
i% = InStr(w$, "(WID:"): If i% > 0 Then w$ = trm(Left$(w$, i% - 1))
i% = List1.ListIndex
k$ = ""
If i% < 0 And InStr(w$, " - ") > 0 Then
  k$ = w$
  Do
    i% = InStr(k$, " - ")
    k$ = trm(Mid$(k$, i% + 2))
  Loop Until InStr(k$, " - ") = 0
  w$ = trm(Left(w$, Len(w$) - (Len(k$) + 2)))
  For i% = 0 To List1.ListCount - 1
    If Left(List1.List(i%), Len(k$)) = k$ Then
      List1.ListIndex = i%
      DoEvents
      k$ = Text3(2).text & " " & Text3(1).text & " (" & Text3(3).text & ")"
      List1.ListIndex = -1
      DoEvents
      Exit For
    End If
  Next i%
Else
  k$ = Text3(2).text & " " & Text3(1).text & " (" & Text3(3).text & ")"
End If
Clipboard.settext k$ & ": " & w$
End Sub


Private Sub Command4_Click()
Dim i%, up$, cmd$, rtmp As QueryDef, id$

'd2infile = "werkvz": d2insub = "Command4_Click"
id$ = Text3(0).text
If id$ = "" Then Exit Sub

If List2.ListCount > 0 Then
  Command4.Visible = False
  Check1.value = 0
  MsgBox "Komponist hat Werke. Löschen nicht möglich."
  Exit Sub
End If

Check4.value = 0
Call Check4_Click
DoEvents
cmd$ = "delete from k_loc where id='" + id$ + "'"
Call form1.sqlqry(cmd$)

Call rlist1
Call showkompdetail("-1")
Call showwerkdetail("-1")
BackColor = form1.cleancolor()
End Sub

Private Sub Command5_Click()
Dim i%, up$, cmd$, rtmp As QueryDef
Dim id$, rlist$, plist$, ask%, c$, r As Recordset

'd2infile = "werkvz": d2insub = "Command5_Click"
id$ = Text4(0).text
If id$ = "" Then Exit Sub

If List4.ListCount > 0 Then
  Command5.Visible = False
  Check2.value = 0
  MsgBox "Werk hat Sätze. Löschen nicht möglich."
  Exit Sub
End If
rlist$ = "": plist$ = ""
If Not form1.isfieldmissing("opt_repertoire", "id") Then
c$ = "select * from opt_repertoire where wid='" + id$ + "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If InStr(rlist$, trm(r!vid)) = 0 Then
    If rlist$ <> "" Then rlist$ = rlist$ + ","
    rlist$ = rlist$ + " " + r!vid
  End If
  r.MoveNext
Wend
End If
c$ = "select * from programmliste where WerkID='" + id$ + "'"
Set r = form1.sqla.OpenRecordset(c$, dbOpenDynaset, dbReadOnly)
While Not r.EOF
  If InStr(plist$, trm(r!programmid)) = 0 Then
    If plist$ <> "" Then plist$ = plist$ + ","
    plist$ = plist$ + " " + r!programmid
  End If
  r.MoveNext
Wend
ask% = vbYes
If plist$ <> "" Or rlist$ <> "" Then
  If plist <> "" Then
    c$ = transe("Das Werk ist in folgenden Programmen enthalten:")
    c$ = c$ + vbCrLf + plist$
  End If
  If rlist <> "" Then
    c$ = c$ + vbCrLf + transe("Das Werk gehört zum Repertoire von:")
    c$ = c$ + vbCrLf + rlist$
  End If
  ask% = MsgBox(transe(c$), vbYesNo + vbCritical + vbDefaultButton2, transe("Wirklich löschen?"))
  If ask% = vbYes Then
    If rlist$ <> "" Then
      c$ = "delete from opt_repertoire where wid='" + id$ + "'": form1.sqlqry (c$)
    End If
    If plist$ <> "" Then
      c$ = "delete from programmliste where WerkID='" + id$ + "'": form1.sqlqry (c$)
    End If
  Else
    Exit Sub
  End If
End If

If ask% = vbYes Then
  cmd$ = "delete from w_loc where id='" + id$ + "'"
  Call form1.sqlqry(cmd$)
  Call List1_Click
End If

End Sub

Private Sub Command6_Click()
Dim i%, up$, cmd$, rtmp1 As ADODB.Recordset, rtmp As QueryDef, rrr
Dim nflds, id$, j%, erg, beidenull

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command6_Click"
id$ = Text4(0).text
If id$ = "" Then Exit Sub
nflds = 16
Set rtmp1 = New ADODB.Recordset
rtmp1.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp1, "SELECT * FROM w_loc where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

For i% = 1 To nflds + 1
  j% = i%: If i% = nflds + 1 Then j% = 31
'Debug.Print Label2(i%).Caption
'Debug.Print trm(rtmp1.Fields(j%)), trm(Text4(i%).Text)
  erg = (trm(rtmp1.Fields(j%)) <> trm(Text4(i%).text))
  beidenull = (IsNull(rtmp1.Fields(j%)) And Len(Text4(i%).text) = 0)
  If Not beidenull And (IsNull(erg) Or erg) Then
    If Len(Text4(i%).text) = 0 Then
      up$ = transo(Label2(i%).Caption) + "=NULL"
    Else
      If LCase(Label2(i%).Caption) = "number" Then
        up$ = "nummer='" + strrepl(Text4(i%).text, "'", "´") + "'"
      Else
        up$ = transo(Label2(i%).Caption) + "= '" + strrepl(Text4(i%).text, "'", "´") + "'"
      End If
    End If
    cmd$ = "update w_loc set " + up$ + " where id='" + id$ + "'"
    Call form1.sqlqry(cmd$)
  End If
Next i%
cmd$ = "update w_loc set stand='" + datum2sql(Date) + "' where id= '" + id$ + "'"
Call form1.sqlqry(cmd$)
'GEMA #
cmd$ = "update w_loc set s14='" + trm(Text5.text) + "' where id= '" + id$ + "'"
Call form1.sqlqry(cmd$)
cmd$ = "update w_loc set s13='" + trm(Text6.text) + "' where id= '" + id$ + "'"
Call form1.sqlqry(cmd$)
If Not form1.isfieldmissing("opt_stimmton", "id") Then
  cmd$ = "delete from opt_stimmton where id= '" + id$ + "'": Call form1.sqlqry(cmd$)
  If trm(txtStimmton.text) <> "" Then
    cmd$ = "insert into opt_stimmton (id,stimmton) values('" + id$ + "','" + trm(txtStimmton) + "')"
    Call form1.sqlqry(cmd$)
  End If
End If

Call List1_Click
Call showwerkdetail(id$)
list2updno = True
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), trm(Text4(1).text)) = 1 Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%
list2updno = False
End Sub

Private Sub Command7_Click()
Dim i%, up$, cmd$, rtmp As QueryDef, nflds As Integer, rrr
Dim stmp As ADODB.Recordset, kid$, id$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command7_Click"
nflds = 31
kid$ = Text3(0).text
If kid$ = "" Then Exit Sub

up$ = "insert into w_loc ("
For i% = 0 To nflds
  up$ = up$ + form1.sqla.TableDefs("w_loc").Fields(i%).name + ","
Next i%
up$ = Left$(up$, Len(up$) - 1) + ") values("
For i% = 0 To nflds
  Select Case i%
    Case 1: Text4(i%).text = "Neues Werk"
    Case 3: Text4(i%).text = kid$
    Case 0: Do
              id$ = form1.newid("w_loc", "id", 10)
              Set stmp = New ADODB.Recordset
              stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT id FROM w_loc where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            Loop Until stmp.EOF
            Text4(i%).text = id$
    Case Else: If i% < 18 Then Text4(i%).text = ""
  End Select

  If i% > 17 Then
    up$ = up$ + "NULL,"
  Else
    If Len(Text4(i%).text) = 0 Then
      up$ = up$ + "NULL,"
    Else
      up$ = up$ + "'" + Text4(i%).text + "',"
    End If
  End If
Next i%
up$ = Left$(up$, Len(up$) - 1)
cmd$ = up$ + ")"
Call form1.sqlqry(cmd$)

Call List1_Click
DoEvents
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), "Neues Werk") = 1 Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%
End Sub

Private Sub Command8_Click()
Dim cmd$, up$, werkid$, id$, i%, rrr
Dim stmp As ADODB.Recordset, eoffl As Boolean

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Command8_Click"
werkid$ = Text4(0).text
If werkid$ = "" Then Exit Sub

up$ = "insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values("
Do
    id$ = Left(trm(str$(Rnd)), 10) + form1.newid("sbz_loc", "id", 10)
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT id FROM sbz_loc where id='" + id$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    eoffl = stmp.EOF
    stmp.Close
Loop Until eoffl

up$ = up$ + "'" + id$ + "','" + werkid$ + "','Neuer Satz'," + Trim$(str$(List4.ListCount * 10))
cmd$ = up$ + ")"
Call form1.sqlqry(cmd$)
For i% = 0 To List2.ListCount - 1
  If InStr(List2.List(i%), trm(Text4(1).text)) = 1 Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%
Call List2_Click
For i% = 0 To List4.ListCount - 1
  If InStr(List4.List(i%), "Neuer Satz") = 1 Then
    List4.ListIndex = i%
    i% = List4.ListCount
    Call List4_DblClick
  End If
Next i%
End Sub

Private Sub Command9_Click()
Dim i%, p%, l$

'd2infile = "werkvz": d2insub = "Command9_Click"
i% = List4.ListIndex
If i% < 0 Then Exit Sub

l$ = List4.List(i%)

p% = InStr(l$, " (ID:")
l$ = Mid$(l$, p% + 5): l$ = Left$(l$, Len(l$) - 1)

List4.RemoveItem i%
Call form1.sqlqry("delete from sbz_loc where id='" + l$ + "'")
If List4.ListCount < 1 Then Command9.Enabled = False
End Sub

Private Sub Form_Load()
Dim s%, V$, i%, tr As String, xd$

'd2infile = "werkvz": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
list2updno = False
Randomize
searching = False
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
s% = form1.myfontsize()
List1.Font.Size = s%
List2.Font.Size = s%
List4.Font.Size = s%
Text1.Font.Size = s%
Text2.Font.Size = s%
Label12.Caption = ""
Label2(4).ForeColor = form1.lnkcolor
For i% = 0 To 10: Text3(i%).Font.Size = s%: Next i%
For i% = 0 To 16: Text4(i%).Font.Size = s%: Next i%
Command21.Picture = Picture3(0).Picture
Command22.Picture = Picture3(0).Picture
V$ = form1.getusersetting("komponistenverzeichnis", "")
On Error Resume Next
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
V$ = V$ & "\" & form1.mkkompdn(trm(Text3(1) & " " & Text3(2)))
tr = Dir(V$ & "\*.*")
If tr <> "" Then Command21.Picture = Picture3(1).Picture
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(1)))
tr = Dir(V$ & "\*.*")
On Error GoTo 0
If tr <> "" Then Command22.Picture = Picture3(1).Picture
xd$ = form1.getusersetting("htmlwerke", "")
If form1.getuserid() = "www" Then
  If xd$ <> "" Then
    Command28.Visible = True
    Command30.Visible = True
  End If
End If
If Not form1.isfieldmissing("opt_repertoire", "id") Then Command31.Visible = True
werkvz.Caption = transe("Werkeverzeichnis")
Command30.Caption = transe("&Export")
Command33.ToolTipText = transe("Markiertes Werk in die Zwischenablage kopieren")
Command26.ToolTipText = transe("Werk zu einem anderen Komponisten verschieben")
Command29.Caption = transe("Werk kopieren")
Command28.Caption = transe("&HTML")
Command27.Caption = transe("Abwahl")
Command25.Caption = transe("Hinweise in Datei speichern")
Command24.Caption = transe("Ende")
Command23.Caption = transe("Pos1")
Command22.ToolTipText = transe("Datenverzeichnis für dieses Werk im Explorer öffnen")
Command21.ToolTipText = transe("Komponistenverzeichnis im Explorer öffnen")
Command20.Caption = transe("Wrk.übertrgn")
Command19.ToolTipText = transe("Hilfeseite öffnen")
Command18.ToolTipText = transe("Suche das Werk bei bei Google")
Command17.ToolTipText = transe("Suche den Künstler bei Google")
Command16.Caption = transe("Notenverzeichnis")
Command16.ToolTipText = transe("Liste der Bezugsquellen (alle Werke/Komponisten)")
Command13.ToolTipText = transe("Notenverzeichnis öffnen / erstellen")
Command14.ToolTipText = transe("Notenverzeichnis öffnen / erstellen")
Command12.Caption = transe("ab")
Command12.ToolTipText = transe("Den markierten Satz um eine Position nach unten")
Command11.Caption = transe("auf")
Command11.ToolTipText = transe("Den markierten Satz um eine Position nach oben")
List4.ToolTipText = transe("Alle Sätze des Werkes")
Command9.ToolTipText = transe("Den markierten Satz löschen")
Command8.ToolTipText = transe("Neuen Satz anlegen")
Text4(10).ToolTipText = transe("Satz suchen")
Check3.Caption = transe("streng")
Command10.Caption = transe("wo gespielt?")
Command10.ToolTipText = transe("Liste der Orte, an denen dieses Werk schon gespielt wurde")
Command7.ToolTipText = transe("Neues Werk anlegen")
Text4(17).ToolTipText = transe("Interne Notizen")
Text4(16).ToolTipText = transe("Nummer im Werkverzeichnis")
Text4(15).ToolTipText = transe("Werkverzeichnis")
Text4(14).ToolTipText = transe("Namensergänzung, Bemerkung")
Text4(13).ToolTipText = transe("Offizieller Name des Werkes")
Text4(12).ToolTipText = transe("Nummer des Werkes")
Command6.ToolTipText = transe("Werk speichern")
Command5.ToolTipText = transe("löschen")
Text4(11).ToolTipText = transe("Tonart")
Text4(8).ToolTipText = transe("Jahr der Fertigstellung der Komposition")
Text4(7).ToolTipText = transe("Jahr des Beginns der Komposition")
Text4(6).ToolTipText = transe("Komplett in welchem Jahr geschrieben")
Text4(5).ToolTipText = transe("Dauer des Gesamtwerks")
Text4(4).ToolTipText = transe("Musikalische Besetzung")
Text4(1).ToolTipText = transe("wird automatisch zusammengesetzt")
Command4.ToolTipText = transe("löschen")
Command3.ToolTipText = transe("Komponisten speichern")
Text3(10).ToolTipText = transe("Todesdatum")
Text3(9).ToolTipText = transe("Geburtsdatum")
Text3(7).ToolTipText = transe("Andere Schreibweisen des Komponistennamens")
Text3(5).ToolTipText = transe("Todesjahr")
Text3(4).ToolTipText = transe("Geburtsjahr")
Text3(3).ToolTipText = transe("Biografische Daten")
Text3(2).ToolTipText = transe("Vorname(n)")
Text3(1).ToolTipText = transe("Name")
Text3(0).ToolTipText = transe("ID")
Text2.ToolTipText = transe("Werk suchen")
Command2.ToolTipText = transe("Neuer Komponist")
List2.ToolTipText = transe("Liste der Werke des markierten Komponisten")
Command1.ToolTipText = transe("schliessen")
Text1.ToolTipText = transe("Komponisten suchen")
List1.ToolTipText = transe("Ausgewählten anklicken")
Picture1.ToolTipText = transe("löschen verboten")
Picture2.ToolTipText = transe("löschen verboten")
Label11.Caption = transe("Stand:")
Label11.ToolTipText = transe("vollständig und korrekt eingegeben")
Label10.Caption = transe("autorisiert")
Label10.ToolTipText = transe("vollständig und korrekt eingegeben")
Label9.Caption = transe("autorisiert")
Label9.ToolTipText = transe("vollständig und korrekt eingegeben")
Label8.Caption = transe("Sätze")
Label7.Caption = transe("Suchen")
Label6.Caption = transe("Werke")
Label5.Caption = transe("Komponist")
Label4.Caption = transe("Suchen")
Label3.Caption = transe("Min.")
If Not form1.isfieldmissing("opt_cocomposers", "id") Then
  List3.Enabled = True
  Combo1.Enabled = True
  Label19(0).Visible = False
End If
If Not form1.isfieldmissing("opt_arranged", "id") Then
  List5.Enabled = True
  Combo2.Enabled = True
  Label19(1).Visible = False
End If
If Not form1.isfieldmissing("opt_published", "id") Then
  List6.Enabled = True
  Combo3.Enabled = True
  Label19(2).Visible = False
End If
If Not form1.isfieldmissing("opt_textdichter", "id") Then
  List7.Enabled = True
  Combo4.Enabled = True
  Label19(3).Visible = False
End If
If Not form1.isfieldmissing("opt_stimmton", "id") Then
  lbl_ston.Visible = True
  txtStimmton.Visible = True
End If
Show
isviz = True
callback$ = ""
Text1.text = ""
Text2.text = ""
Command4.Visible = False
Command5.Visible = False
Command6.Enabled = False
BackColor = form1.cleancolor()
Command10.Enabled = False
Command29.Enabled = False

Call rlist1
Call showkompdetail("-1")
Call showwerkdetail("-1")
BackColor = form1.cleancolor()

End Sub


Private Sub Form_Resize()
'd2infile = "werkvz": d2insub = "Form_Resize"
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

'd2infile = "werkvz": d2insub = "Form_Unload"
isviz = False
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0

End Sub

Private Sub Label1_DblClick(Index As Integer)
Clipboard.Clear
Clipboard.settext Text3(0).text
End Sub

Private Sub Label10_Click()
'd2infile = "werkvz": d2insub = "Label10_Click"
If Check5.value = 1 Then
  Check5.value = 0
Else
  Check5.value = 1
End If

End Sub

Private Sub Label19_Click(Index As Integer)
Dim i%

For i% = 0 To List2.ListCount - 1
  Select Case (Index)
    Case 0:
        If List2.List(i%) = "from Cocomposers:" Then
          List2.ListIndex = i%
          Exit Sub
        End If
    Case 1:
        If List2.List(i%) = "from Arrangement:" Then
          List2.ListIndex = i%
          Exit Sub
        End If
    Case 2:
        If List2.List(i%) = "from Publisher:" Then
          List2.ListIndex = i%
          Exit Sub
        End If
    Case 3:
        If List2.List(i%) = "from Librettist:" Then
          List2.ListIndex = i%
          Exit Sub
        End If
    Case Else:
  End Select
Next i%
End Sub

Private Sub Label2_DblClick(Index As Integer)
Dim rtmp As ADODB.Recordset, id$, rrr


Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Label2_DblClick"
If Index = 4 Then
  id$ = Text4(0).text
  If id$ = "" Then Exit Sub
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM b_loc where wid='" & id$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then
    MsgBox "Die Datenbankstruktur muss erst geändert werden. (b_loc fehlt)." & vbCrLf & "Bitte lontaktieren Sie Ihren Support"
    Exit Sub
  End If
  Load besetzung
  DoEvents
  besetzung.werkid.Caption = id$
End If

End Sub

Private Sub Label9_Click()
'd2infile = "werkvz": d2insub = "Label9_Click"
If Check4.value = 1 Then
  Check4.value = 0
Else
  Check4.value = 1
End If

End Sub

Private Sub List1_Click()
Dim kid$, i%, V$, tr, v0$

'd2infile = "werkvz": d2insub = "List1_Click"
Command17.Enabled = False
Command18.Enabled = False
Command33.Enabled = False
If List1.ListIndex < 0 Then Exit Sub
'If searching Then Exit Sub
Command17.Enabled = True
kid$ = List1.List(List1.ListIndex)
kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
kid$ = Left$(kid$, InStr(kid$, ")") - 1)
Command13.Picture = Picture4(0).Picture
Call rlist2(kid$)
Call showkompdetail(kid$)
Call showwerkdetail("-1")
Command10.Enabled = False
Command29.Enabled = False
For i% = 0 To List2.ListCount - 1
  If InStr(LCase$(List2.List(i%)), LCase$(Text2.text)) > 0 Then
    List2.ListIndex = i%
    i% = List2.ListCount
  End If
Next i%
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
Command14.Picture = Picture4(0).Picture
Command21.Picture = Picture3(0).Picture
v0$ = V$ & "\_KOMPONISTEN_"
V$ = v0$ & "\" & form1.mkkompdn(trm(Text3(1) & " " & Text3(2)))
Call form1.dbg2f("werkvz.list1click:" & V$)
On Error Resume Next
tr = Dir(V$ & "\*.*")
If tr <> "" Then Command21.Picture = Picture3(1).Picture
V$ = v0$ & "\" & form1.mkkompdn(trm(Text3(0)))
tr = Dir(V$ & "\*.*", vbDirectory)
If tr <> "" Then Command14.Picture = Picture4(1).Picture
On Error GoTo 0
BackColor = form1.cleancolor()
Timer1.Enabled = False
Timer1.Interval = 500
Timer1.Enabled = True

End Sub

Private Sub List2_Click()
Dim kid$, o%, i%, r As ADODB.Recordset, nsatz As Boolean, p%, q%, logf$, tbl$, id$
Dim l$, c$, sfnd As Boolean, cmd$, j%, X

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "List2_Click"
If list2updno Then Exit Sub
Command18.Enabled = False
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then Exit Sub
Command18.Enabled = True
kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
Call showwerkdetail(kid$)
Command10.Enabled = True

End Sub

Private Sub List2_DblClick()
Dim kid$, knr

'd2infile = "werkvz": d2insub = "List2_DblClick"
If List2.ListIndex < 0 Then Exit Sub
kid$ = List2.List(List2.ListIndex)
If InStr(kid$, "(WID:") = 0 Then
  If InStr(kid$, "(KID: ") <> 0 Then
    kid$ = Mid$(kid$, InStr(kid$, "(KID: ") + 6)
    kid$ = Left$(kid$, Len(kid$) - 1)
    Text1.text = form1.getkompnamebyid(kid$)
  End If
  Exit Sub
End If
kid$ = Mid$(kid$, InStr(kid$, "(WID:") + 5)
Text1.text = form1.getkompnamebywerkid(kid$)
If callback$ <> "" Then
  Select Case callback$
    Case "repertoire": Call repertoire.callback(kid$)
                     Call repertoire.SetFocus
    Case "prog": Call prog.callback(kid$)
                 Text2.text = ""
                 searching = False
                 If List1.ListIndex >= 0 Then
                   Call List1_Click
                 End If
                 Call prog.SetFocus
    Case Else:
  End Select
  callback$ = ""
  werkvz.Caption = "Werkeverzeichnis"
  isviz = False
  Hide
End If

End Sub

Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim id$, p%, wid$, c$, i%

If KeyCode = 46 Then

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

i% = List3.ListIndex: If i% < 0 Then Exit Sub
id$ = List3.List(i%)
p% = InStr(id$, "(KID:")
If p% = 0 Then Exit Sub
id$ = Mid$(id$, p% + 5)
c$ = "delete from opt_cocomposers where kid='" + id$ + "' and wid='" + wid$ + "'"
Call form1.sqlqry(c$)
Call rlist3

End If
End Sub

Private Sub List4_Click()

'd2infile = "werkvz": d2insub = "List4_Click"
Command9.Enabled = True

End Sub

Private Sub List4_DblClick()
Dim asatz$, bsatz$, i%, rtmp As ADODB.Recordset, l$, p%, cmd$, sbzid$, rrr

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "List4_DblClick"
i% = List4.ListIndex
l$ = List4.List(i%)
p% = InStr(l$, " (ID:")
l$ = Mid$(l$, p% + 5): l$ = Left$(l$, Len(l$) - 1)
sbzid$ = l$
cmd$ = "select * from sbz_loc where id='" + l$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
asatz$ = trm(rtmp!satzbezeichnung)

If callback$ <> "" Then
  Select Case callback$
    Case "prog": Call prog.callback("SBZ:" + sbzid$)
                 Call prog.SetFocus
    Case Else
  End Select
  callback$ = ""
  werkvz.Caption = "Werkeverzeichnis"
  isviz = False
  Hide
  Exit Sub
End If

bsatz$ = trm(InputBox(transe("Satzangabe bearbeiten"), transe("Satzangabe bearbeiten"), asatz$))
bsatz$ = strrepl(bsatz, "'", "´")
If asatz$ <> bsatz$ And trm(bsatz$) <> "" Then
  cmd$ = "update sbz_loc set Satzbezeichnung='" & bsatz$ & "'"
  cmd$ = cmd$ + " where id='" + l$ + "'"
  Call form1.sqlqry(cmd$)
  Call List2_Click
End If

End Sub

Private Sub List4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fn$, nf%, rrr, o%, l$, up$, cnt%, bemmode As Boolean, werkid$, bem$

'd2infile = "werkvz": d2insub = "List4_OLEDragDrop"
werkid$ = Text4(0).text
If werkid$ = "" Then Exit Sub

If List4.ListCount > 0 Then
  MsgBox "DragDrop ist nur bei leerer Liste möglich."
  Exit Sub
End If
bemmode = False
cnt% = 1
If Data.GetFormat(vbCFFiles) Then
  nf% = 1
  Do
    On Error Resume Next
    fn$ = Data.Files(nf%)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      up$ = "insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & _
             form1.newid("sbz_loc", "id", 40) & "','" & werkid$ & "','Noten: hier',0)"
      Call form1.sqlqry(up$)
      List4.AddItem l$
      o% = FreeFile
      Open fn$ For Input As #o%
      If Text4(1).text = "Neues Werk" Then
        Line Input #o%, l$
        l$ = strrepl(trm(l$), "'", "´")
        l$ = strrepl(l$, """", "´´")
        Text4(1).text = l$
        Text4(13).text = l$
      End If
      While (Not EOF(o%)) And (Not bemmode)
        Line Input #o%, l$
        If InStr(l$, "UNICORN") > 0 Or InStr(l$, "BEMERKUNG") > 0 Or InStr(l$, "HINWEISE") > 0 Then
          bemmode = True
        Else
          l$ = strrepl(trm(l$), "'", "´")
          If l$ <> "" Then
            l$ = strrepl(l$, """", "´´")
            cnt% = cnt% + 1
            up$ = "insert into sbz_loc (id,wid,satzbezeichnung,satznummer) values('" & _
               form1.newid("sbz_loc", "id", 40) & "','" & werkid$ & "','" & l$ & "'," & trm(str$(cnt)) & ")"
            Call form1.sqlqry(up$)
            List4.AddItem l$
            List4.ListIndex = List4.ListCount - 1
            DoEvents
          End If
        End If
      Wend
      If bemmode Then
        bem$ = l$
        While (Not EOF(o%))
          Line Input #o%, l$
          l$ = strrepl(trm(l$), "'", "´")
          If l$ <> "" Then
            l$ = strrepl(l$, """", "´´")
            bem$ = bem$ & vbCrLf & l$
          End If
        Wend
        If Text4(17).text <> "" Then Text4(17).text = Text4(17).text & vbCrLf
        Text4(17).text = Text4(14).text & bem$
      End If
      Close #o%
    End If
    nf% = nf% + 1
  Loop Until rrr <> 0
End If

End Sub

Private Sub List5_dblClick()
Dim i%, id$

i% = List5.ListIndex
If i% < 0 Then Exit Sub
id$ = List5.List(i%)
Load shwAdrDetail
Call shwAdrDetail.refreshadrdetail(id$, "-1")
On Error Resume Next
Call shwAdrDetail.SetFocus
On Error GoTo 0

End Sub

Private Sub List5_KeyUp(KeyCode As Integer, Shift As Integer)
Dim id$, p%, wid$, c$, i%

If KeyCode = 46 Then

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

i% = List5.ListIndex: If i% < 0 Then Exit Sub
id$ = List5.List(i%)
c$ = "delete from opt_arranged where aid='" + id$ + "' and wid='" + wid$ + "'"
Call form1.sqlqry(c$)
Call rlist5

End If
End Sub

Private Sub List6_DblClick()
Dim i%, id$

i% = List6.ListIndex
If i% < 0 Then Exit Sub
id$ = List6.List(i%)
Load shwAdrDetail
Call shwAdrDetail.refreshadrdetail(id$, "-1")
On Error Resume Next
Call shwAdrDetail.SetFocus
On Error GoTo 0
End Sub

Private Sub List7_DblClick()
Dim i%, id$

i% = List7.ListIndex
If i% < 0 Then Exit Sub
id$ = List7.List(i%)
Load shwAdrDetail
Call shwAdrDetail.refreshadrdetail(id$, "-1")
On Error Resume Next
Call shwAdrDetail.SetFocus
On Error GoTo 0
End Sub

Private Sub List6_KeyUp(KeyCode As Integer, Shift As Integer)
Dim id$, p%, wid$, c$, i%

If KeyCode = 46 Then

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

i% = List6.ListIndex: If i% < 0 Then Exit Sub
id$ = List6.List(i%)
c$ = "delete from opt_published where aid='" + id$ + "' and wid='" + wid$ + "'"
Call form1.sqlqry(c$)
Call rlist6

End If

End Sub

Private Sub List7_KeyUp(KeyCode As Integer, Shift As Integer)
Dim id$, p%, wid$, c$, i%

If KeyCode = 46 Then

wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

i% = List7.ListIndex: If i% < 0 Then Exit Sub
id$ = List7.List(i%)
c$ = "delete from opt_textdichter where aid='" + id$ + "' and wid='" + wid$ + "'"
Call form1.sqlqry(c$)
Call rlist7

End If

End Sub

Private Sub Text1_Change()
'd2infile = "werkvz": d2insub = "Text1_Change"
Command17.Enabled = False
Command18.Enabled = False
searching = True
Call timerreset

End Sub

Public Sub Text2_Change()

'd2infile = "werkvz": d2insub = "Text2_Change"
Timer1.Enabled = False
Timer1.Interval = 500
Timer1.Enabled = True

End Sub

Private Sub Text3_Change(Index As Integer)
Dim i%

'd2infile = "werkvz": d2insub = "Text3_Change"
i% = Index

Command3.Enabled = True
If i% = 4 Or i% = 5 Then Text3(3).text = Text3(4).text & "-" & Text3(5).text
BackColor = form1.dirtycolor()

End Sub

Private Sub Text4_change(Index As Integer)
Dim i%, t1$, t$, streng%

'd2infile = "werkvz": d2insub = "Text4_change"
Command6.Enabled = True
BackColor = form1.dirtycolor()
streng% = Check3.value
If Index > 10 And Index < 17 And streng% <> 0 Then
  i% = Index
  t1$ = "": If trm(Text4(12).text) <> "" Then t1$ = " Nr. " + trm(Text4(12).text)
  t$ = trm(Text4(13).text) + t1$ + " " + _
       trm(Text4(11).text) + " " + trm(Text4(15).text) + " " + _
       trm(Text4(16).text) + " " + trm(Text4(14).text)
  While InStr(t$, "  ") > 0: t$ = strrepl(t$, "  ", " "): Wend
  Text4(1).text = t$
End If

End Sub

Private Sub Text5_Change()
Command6.Enabled = True
BackColor = form1.dirtycolor()

End Sub

Private Sub Text6_Change()
Call Text5_Change
End Sub

Private Sub Timer1_Timer()
Dim s$, rtmp As ADODB.Recordset, i%, c$, ca$, j%, cb$, kid$, rrr, sworte As Integer, sorder$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "Timer1_Timer"
Call form1.dbg2f("werkvz Timer1 start")
Timer1.Enabled = False
sorder$ = form1.getusersetting("sortierewerke", "name")
s$ = LCase(trm(Text2.text))
If s$ = "" Then
  Call form1.dbg2f("werkvz Timer1 exit")
  Exit Sub
End If
List2.Clear
For i% = 0 To 4: swrd$(i%) = "": Next i%
For i% = 0 To 3: Label19(i%).Visible = False: Next i%
DoEvents
i% = 0
While Len(s$) > 0 And i% < 5
  swrd$(i%) = word1(s$)
  s$ = word2bis(s$)
  i% = i% + 1
Wend
i% = i% - 1: sworte = i%
c$ = "SELECT id,name,KomponistenNummer FROM w_loc where ("
ca$ = ""
For j% = 0 To i%
  If ca$ <> "" Then ca$ = ca$ & "and "
  ca$ = ca$ & "(instr(lcase(name),'" + swrd$(j%) + "')>0 or instr(lcase(s14),'" + swrd$(j%) + "')>0 or instr(lcase(s13),'" + swrd$(j%) + "')>0) "
Next j%
cb$ = ""
kid$ = ""
If List1.ListIndex >= 0 Then
  kid$ = List1.List(List1.ListIndex)
  kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
  kid$ = Left$(kid$, InStr(kid$, ")") - 1)
  cb$ = " and (komponistennummer='" & kid$ & "') "
End If
c$ = c$ & ca$ & cb$ & ") order by " & sorder$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  rtmp.MoveFirst
  i% = 0
  While Not rtmp.EOF And i% < 1999
    i% = i% + 1
    If Not IsNull(rtmp!name) Then List2.AddItem rtmp!name + " - " + form1.getkompnamebyid(rtmp!KomponistenNummer) & Space$(160) & " (KID: " + rtmp!KomponistenNummer + ")" & Space$(160) & "(WID:" + rtmp!id
    rtmp.MoveNext
  Wend
End If
rtmp.Close

If Not form1.isfieldmissing("opt_cocomposers", "id") Then
  c$ = "SELECT w_loc.id,w_loc.name,k_loc.Name as kname,KomponistenNummer from (k_loc INNER JOIN opt_cocomposers ON k_loc.id = opt_cocomposers.kid) INNER JOIN w_loc ON opt_cocomposers.wid = w_loc.id where ("
  ca$ = ""
  For j% = 0 To sworte
    If ca$ <> "" Then ca$ = ca$ & "and "
    ca$ = ca$ & "(instr(lcase(k_loc.Name),'" + swrd$(j%) + "')>0 or instr(lcase(vornamen),'" + swrd$(j%) + "')>0) "
  Next j%
  cb$ = ""
  kid$ = ""
  If List1.ListIndex >= 0 Then
    kid$ = List1.List(List1.ListIndex)
    kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
    kid$ = Left$(kid$, InStr(kid$, ")") - 1)
    cb$ = " and (komponistennummer='" & kid$ & "') "
  End If
  c$ = c$ & ca$ & cb$ & ") order by k_loc.Name,w_loc.id;"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 And Not rtmp.EOF Then
    rtmp.MoveFirst
    List2.AddItem "from Cocomposers:"
    Label19(0).Visible = True
    i% = 0
    While Not rtmp.EOF And i% < 1999
      i% = i% + 1
      If Not IsNull(rtmp!name) Then List2.AddItem rtmp!name + " - " + form1.getkompnamebyid(rtmp!KomponistenNummer) & Space$(160) & " (KID: " + rtmp!KomponistenNummer + ")" & Space$(160) & "(WID:" + rtmp!id
      rtmp.MoveNext
    Wend
  End If
  rtmp.Close
End If

If Not form1.isfieldmissing("opt_arranged", "id") Then
  c$ = "SELECT w_loc.id,w_loc.name,adresse.Name as kname,KomponistenNummer from (adresse INNER JOIN opt_arranged ON adresse.ID = opt_arranged.aid) INNER JOIN w_loc ON opt_arranged.wid = w_loc.id where ("
  ca$ = ""
  For j% = 0 To sworte
    If ca$ <> "" Then ca$ = ca$ & "and "
    ca$ = ca$ & "(instr(lcase(adresse.Name),'" + swrd$(j%) + "')>0 or instr(lcase(adresse.id),'" + swrd$(j%) + "')>0) "
  Next j%
  cb$ = ""
  kid$ = ""
  If List1.ListIndex >= 0 Then
    kid$ = List1.List(List1.ListIndex)
    kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
    kid$ = Left$(kid$, InStr(kid$, ")") - 1)
    cb$ = " and (komponistennummer='" & kid$ & "') "
  End If
  c$ = c$ & ca$ & cb$ & ") order by adresse.Name,w_loc.id;"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 And Not rtmp.EOF Then
    rtmp.MoveFirst
    List2.AddItem "from Arrangement:"
    Label19(1).Visible = True
    i% = 0
    While Not rtmp.EOF And i% < 1999
      i% = i% + 1
      If Not IsNull(rtmp!name) Then List2.AddItem rtmp!name + " - " + form1.getkompnamebyid(rtmp!KomponistenNummer) & Space$(160) & " (KID: " + rtmp!KomponistenNummer + ")" & Space$(160) & "(WID:" + rtmp!id
      rtmp.MoveNext
    Wend
  End If
  rtmp.Close
End If
If Not form1.isfieldmissing("opt_published", "id") Then
  c$ = "SELECT w_loc.id,w_loc.name,adresse.Name as kname,KomponistenNummer from (adresse INNER JOIN opt_published ON adresse.ID = opt_published.aid) INNER JOIN w_loc ON opt_published.wid = w_loc.id where ("
  ca$ = ""
  For j% = 0 To sworte
    If ca$ <> "" Then ca$ = ca$ & "and "
    ca$ = ca$ & "(instr(lcase(adresse.Name),'" + swrd$(j%) + "')>0 or instr(lcase(adresse.id),'" + swrd$(j%) + "')>0) "
  Next j%
  cb$ = ""
  kid$ = ""
  If List1.ListIndex >= 0 Then
    kid$ = List1.List(List1.ListIndex)
    kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
    kid$ = Left$(kid$, InStr(kid$, ")") - 1)
    cb$ = " and (komponistennummer='" & kid$ & "') "
  End If
  c$ = c$ & ca$ & cb$ & ") order by adresse.Name,w_loc.id;"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 And Not rtmp.EOF Then
    rtmp.MoveFirst
    List2.AddItem "from Publisher:"
    Label19(2).Visible = True
    i% = 0
    While Not rtmp.EOF And i% < 1999
      i% = i% + 1
      If Not IsNull(rtmp!name) Then List2.AddItem rtmp!name + " - " + form1.getkompnamebyid(rtmp!KomponistenNummer) & Space$(160) & " (KID: " + rtmp!KomponistenNummer + ")" & Space$(160) & "(WID:" + rtmp!id
      rtmp.MoveNext
    Wend
  End If
  rtmp.Close
End If
If Not form1.isfieldmissing("opt_textdichter", "id") Then
  c$ = "SELECT w_loc.id,w_loc.name,adresse.Name as kname,KomponistenNummer from (adresse INNER JOIN opt_textdichter ON adresse.ID = opt_textdichter.aid) INNER JOIN w_loc ON opt_textdichter.wid = w_loc.id where ("
  ca$ = ""
  For j% = 0 To sworte
    If ca$ <> "" Then ca$ = ca$ & "and "
    ca$ = ca$ & "(instr(lcase(adresse.Name),'" + swrd$(j%) + "')>0 or instr(lcase(adresse.id),'" + swrd$(j%) + "')>0) "
  Next j%
  cb$ = ""
  kid$ = ""
  If List1.ListIndex >= 0 Then
    kid$ = List1.List(List1.ListIndex)
    kid$ = Mid$(kid$, InStr(kid$, "(ID:") + 4)
    kid$ = Left$(kid$, InStr(kid$, ")") - 1)
    cb$ = " and (komponistennummer='" & kid$ & "') "
  End If
  c$ = c$ & ca$ & cb$ & ") order by adresse.Name,w_loc.id;"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 And Not rtmp.EOF Then
    rtmp.MoveFirst
    List2.AddItem "from Librettist:"
    Label19(3).Visible = True
    i% = 0
    While Not rtmp.EOF And i% < 1999
      i% = i% + 1
      If Not IsNull(rtmp!name) Then List2.AddItem rtmp!name + " - " + form1.getkompnamebyid(rtmp!KomponistenNummer) & Space$(160) & " (KID: " + rtmp!KomponistenNummer + ")" & Space$(160) & "(WID:" + rtmp!id
      rtmp.MoveNext
    Wend
  End If
  rtmp.Close
End If

Call form1.dbg2f("werkvz Timer1 exit")
End Sub
Public Sub showkompdetailbyname(uId$)

'd2infile = "werkvz": d2insub = "showkompdetailbyname"
Text1.text = uId$
DoEvents

End Sub

Public Sub showkompdetail(uId$)
Dim nflds, rtmp As ADODB.Recordset, i%, V$, tr As String, w$, rrr

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "showkompdetail"
nflds = 10
i% = 0

For i% = 0 To nflds
  Label1(i%).Caption = transe(form1.sqla.TableDefs("k_loc").Fields(i%).name)
  Text3(i%).text = ""
Next i%
Command21.Picture = Picture3(0).Picture

Command3.Enabled = False
Check1.value = 1
Command4.Visible = False
Command7.Enabled = False
Check4.value = 0
If uId$ = "-1" Then Exit Sub

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM k_loc where id ='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  For i% = 0 To nflds
    On Error Resume Next
    w$ = trm(rtmp.Fields(i%))
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      Text3(i%).text = strrepl(trm(rtmp.Fields(i%)), "Null", "")
      If Text3(i%).text = "-" Then Text3(i%).text = ""
    End If
  Next i%
End If
Command3.Enabled = False
Check1.value = 1
Command4.Visible = False
Command7.Enabled = True
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM aut_werke where tabelle ='k_loc' and tabid='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
If Not rtmp.EOF Then
  Check4.value = 1
End If
End If
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
V$ = V$ & "\_KOMPONISTEN_"
V$ = V$ & "\" & form1.mkkompdn(trm(Text3(1) & " " & Text3(2)))
Call form1.dbg2f("werkvz:showkompdetail: V=" + V$)
On Error Resume Next
tr = Dir(V$ & "\*.*")
On Error GoTo 0
If tr <> "" Then Command21.Picture = Picture3(1).Picture
End Sub

Public Sub showwerkdetail(uId$)
Dim nflds, streng%, rtmp As ADODB.Recordset, i%, renum%, sbz$, id$, V$, tr As String, v0$
Dim rrr, kn$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "showwerkdetail"
Command22.Picture = Picture3(0).Picture
Command13.Picture = Picture4(0).Picture
streng% = Check3.value
nflds = 16
i% = 0

For i% = 0 To nflds
  Label2(i%).Caption = transe(form1.sqla.TableDefs("w_loc").Fields(i%).name)
  Text4(i%).text = ""
Next i%
Label2(17).Caption = transe(form1.sqla.TableDefs("w_loc").Fields(31).name)
Text4(17).text = ""
txtStimmton.text = ""

List4.Clear
Check5.value = 0
Check2.value = 1
Command5.Visible = False
Command6.Enabled = False
BackColor = form1.cleancolor()
Command9.Enabled = False
Command10.Enabled = False
Command29.Enabled = False
If uId$ = "-1" Then Exit Sub
Command10.Enabled = True
Command29.Enabled = True
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM w_loc where id ='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  For i% = 0 To nflds
    Text4(i%).text = ""
    If Not IsNull(rtmp.Fields(i%)) Then Text4(i%).text = trm(rtmp.Fields(i%))
    If Text4(i%).text = "Null" Then Text4(i%).text = ""
  Next i%
  Text4(17).text = ""
  i% = 31: If Not IsNull(rtmp.Fields(i%)) Then Text4(17).text = rtmp.Fields(i%)
  If Text4(17).text = "Null" Then Text4(17).text = ""
End If
Text5.text = trm(rtmp!s14)
Text6.text = trm(rtmp!s13)
renum% = 0
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM sbz_loc where wid ='" + uId$ + "' order by satznummer", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  sbz$ = "": If Not IsNull(rtmp!satzbezeichnung) Then sbz$ = rtmp!satzbezeichnung
  List4.AddItem sbz$ + Space$(80) + " (ID:" + rtmp!id + ")"
  If rtmp!SatzNummer <> List4.ListCount Then renum% = 1
  rtmp.MoveNext
Wend
If renum% = 1 Then
  For i% = 0 To List4.ListCount - 1
    id$ = List4.List(i%)
    id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
    id$ = Left$(id$, Len(id$) - 1)
    form1.sqlqry ("update sbz_loc set satznummer=" & i% + 1 & " where id='" & id$ & "'")
  Next i%
End If
If Not form1.isfieldmissing("opt_stimmton", "id") Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT * FROM opt_stimmton where id='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
  If Not rtmp.EOF Then
    txtStimmton.text = trm(rtmp!stimmton)
  End If
  End If
End If
Check2.value = 1
Command5.Visible = False
Command6.Enabled = False
BackColor = form1.cleancolor()
Command9.Enabled = False
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM aut_werke where tabelle ='w_loc' and tabid='" + uId$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
If Not rtmp.EOF Then
  Check5.value = 1
End If
End If
If streng% = 1 Then
  Text4(1).Enabled = False
Else
  Text4(1).Enabled = True
End If
Call rlist3
Call rlist5
Call rlist6
Call rlist7
V$ = form1.getusersetting("komponistenverzeichnis")
If V$ = "" Then V$ = form1.s0dir() & "\" & form1.docs()
v0$ = V$ & "\_KOMPONISTEN_"
kn$ = form1.getkompnamebyid(form1.getkompidbywerkid(trm(Text4(0).text)))
V$ = v0$ & "\" & form1.mkkompdn(strrepl(kn$, ",", ""))
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(1).text))
Call form1.dbg2f("werkvz:showwerkdetail: V=" + V$)
On Error Resume Next
tr = Dir(V$ & "\*.*")
If tr <> "" Then Command22.Picture = Picture3(1).Picture
V$ = v0$ & "\" & form1.mkkompdn(form1.getkompidbywerkid(trm(Text4(0).text)))
V$ = V$ & "\" & form1.mkkompdn(trm(Text4(0).text))
tr = Dir(V$ & "/*.*")
On Error GoTo 0
If tr <> "" Then Command13.Picture = Picture4(1).Picture
End Sub


Public Sub callbackinit(frm$)
'd2infile = "werkvz": d2insub = "callbackinit"
callback$ = frm$
werkvz.Caption = "Werkeverzeichnis - Bitte ein Werk auswählen (Doppelklick)"

End Sub
Sub timerreset()
'd2infile = "werkvz": d2insub = "timerreset"
Timer2.Enabled = False
tm_brk% = 1
DoEvents
Timer2.Interval = form1.getsuchvz()
Timer2.Enabled = True

End Sub

Public Sub Timer2_Timer()
Dim i%, s$, l%

'd2infile = "werkvz": d2insub = "Timer2_Timer"
Timer2.Enabled = False
tm_brk% = 0
'DoEvents

s$ = trm(Text1.text)
l% = Len(s$)
If l% = 0 Then Exit Sub
Call form1.dbg2f("werkvz Timer2 start")
For i% = 0 To List1.ListCount - 1
  If tm_brk% = 1 Then Exit For
  If Left(List1.List(i%), l%) = s$ Then
    If i% < List1.ListCount - 5 Then List1.ListIndex = i% + 4
    searching = False
    List1.ListIndex = i%
    Call form1.dbg2f("werkvz Timer2 exit")
    Exit Sub
  End If
Next i%
For i% = 0 To List1.ListCount - 1
  If tm_brk% = 1 Then Exit For
  If LCase(Left(trm(List1.List(i%)), l%)) = LCase(s$) Then
    If i% < List1.ListCount - 5 Then List1.ListIndex = i% + 4
    searching = False
    List1.ListIndex = i%
    Call form1.dbg2f("werkvz Timer2 exit")
    Exit Sub
  End If
Next i%

tm_brk% = 0
Call form1.dbg2f("werkvz Timer2 exit")
End Sub

Function lobgesang() As String
'd2infile = "werkvz": d2insub = "lobgesang"
lobgesang = lobgesang & "<table><tr><td valign=top><a href=http://www.sks-russ.de target=_blank><img src=/site/bilder/SKS3.JPG></a></td>"
lobgesang = lobgesang & "<td valign=top>"
lobgesang = lobgesang & "<p><span class=""boldblue"">Wir begrüßen alle Besucher/innen dieser und <a href=http://www.sks-russ.de target=_blank>unserer Webseite</a>.</span></p><br>"
lobgesang = lobgesang & "<p>Die hier befindlichen Informationen sind ein Ergebnis jahrelanger Beschäftigung mit der Materie als Mitarbeiter im Konzertwesen. Die Angaben erheben weder Anspruch auf wissenschaftliche Authentizität, noch auf korrekte Schreibweise der Namen oder gar Vollständigkeit.</p>"
lobgesang = lobgesang & "<p>Von Tippfehlern abgesehen, haben wir uns um Einheitlichkeit der Schreibweise in Deutsch bemüht, wobei gewisse Fremdsprachenkenntnis (auch russisch, deshalb Rachmaninow, und nicht Rachmaninoff oder Rakhmaninov oder …) hilfreich ist. Den Religionskrieg zum Thema Dur-dur/Moll-moll haben wir auf unsere Weise beendet.</p>"
lobgesang = lobgesang & "</td></tr></table><table><tr><td valign=top><p>Glücklich wären wir über Ihre Mithilfe beim Erweitern der Werksangaben. Schicken Sie uns Ihre Informationen, möglichst mit Zeitangaben, auf jeden Fall aber mit Quellenangabe zu. Wir werden die Angaben prüfen und zeitnah einpflegen.</p>"
lobgesang = lobgesang & "<p>Über zufriedene Besucher, Anregungen, Ergänzungen und Tipps freuen wir uns immer.</p>"
lobgesang = lobgesang & "<p><a href=mailto:conrad.haas@sks-russ.de><b>Conrad Haas</b></a></p>"
lobgesang = lobgesang & "<p>und alle Helfer/innen</p>"
'lobgesang = lobgesang & "<br><br><br><p><small>Stand: " & Date & " " & Time & "</small></p></td></tr></table>"
lobgesang = lobgesang & "<br><br></td></tr></table>"
End Function

Private Sub txtStimmton_Change()
Command6.Enabled = True
BackColor = form1.dirtycolor()
End Sub

Sub rlist5()
Dim rtmp As ADODB.Recordset, rrr, na$, wid$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist5"
List5.Clear
If form1.isfieldmissing("opt_arranged", "id") Then Exit Sub
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT aid,wid from opt_arranged where wid='" + wid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  na$ = trm(rtmp!aid)
  List5.AddItem na$
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Sub rlist6()
Dim rtmp As ADODB.Recordset, rrr, na$, wid$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist6"
List6.Clear
If form1.isfieldmissing("opt_arranged", "id") Then Exit Sub
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT aid,wid from opt_published where wid='" + wid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  na$ = trm(rtmp!aid)
  List6.AddItem na$
  rtmp.MoveNext
Wend
rtmp.Close

End Sub

Sub rlist7()
Dim rtmp As ADODB.Recordset, rrr, na$, wid$

Dim d2infile As String, d2insub As String
d2infile = "werkvz": d2insub = "rlist7"
List7.Clear
If form1.isfieldmissing("opt_textdichter", "id") Then Exit Sub
wid$ = List2.List(List2.ListIndex)
If InStr(wid$, "(WID:") = 0 Then Exit Sub
wid$ = Mid$(wid$, InStr(wid$, "(WID:") + 5)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT aid,wid from opt_textdichter where wid='" + wid$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Sub
rtmp.MoveFirst
While Not rtmp.EOF
  na$ = trm(rtmp!aid)
  List7.AddItem na$
  rtmp.MoveNext
Wend
rtmp.Close

End Sub


