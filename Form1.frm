VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "   Haupt-Formular - AgencyProf - "
   ClientHeight    =   4395
   ClientLeft      =   1095
   ClientTop       =   1860
   ClientWidth     =   9840
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   9840
   Visible         =   0   'False
   Begin VB.PictureBox cb4 
      Height          =   135
      Left            =   6240
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   108
      Top             =   3900
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox cb3 
      Height          =   135
      Left            =   6240
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   107
      Top             =   4140
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   615
      Left            =   5160
      TabIndex        =   103
      ToolTipText     =   "CPU-Auslastung"
      Top             =   2160
      Visible         =   0   'False
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   180
      Left            =   5160
      TabIndex        =   104
      ToolTipText     =   "CPU-Auslastung"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox cloudupds 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      IntegralHeight  =   0   'False
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   100
      ToolTipText     =   "Alles für heute Relevante aus dem Kalender"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.ListBox hordex 
      Height          =   1530
      IntegralHeight  =   0   'False
      Left            =   6360
      TabIndex        =   102
      Top             =   5640
      Width           =   6855
   End
   Begin VB.Timer cldpusher 
      Enabled         =   0   'False
      Left            =   5640
      Top             =   1800
   End
   Begin VB.Timer tmrcld 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   1800
   End
   Begin VB.CommandButton btncld 
      Caption         =   "Cloud"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   101
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   "Alles für heute Relevante aus dem Kalender"
      Top             =   2760
      Width           =   6255
   End
   Begin MSComctlLib.ProgressBar pbg1 
      Height          =   375
      Left            =   10800
      TabIndex        =   99
      ToolTipText     =   "CPU-Auslastung"
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox altbvorl 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0CCA
      Left            =   3240
      List            =   "Form1.frx":0CCC
      TabIndex        =   55
      ToolTipText     =   "zu öffnendes Verzeichnis"
      Top             =   3840
      Width           =   2895
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   14
      IntegralHeight  =   0   'False
      Left            =   4680
      TabIndex        =   98
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   13
      IntegralHeight  =   0   'False
      Left            =   4800
      TabIndex        =   97
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   12
      IntegralHeight  =   0   'False
      Left            =   4920
      TabIndex        =   96
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   11
      IntegralHeight  =   0   'False
      Left            =   5040
      TabIndex        =   95
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   10
      IntegralHeight  =   0   'False
      Left            =   4920
      TabIndex        =   94
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   14
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   93
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   13
      IntegralHeight  =   0   'False
      Left            =   3360
      TabIndex        =   92
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   12
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   91
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   11
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   90
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   10
      IntegralHeight  =   0   'False
      Left            =   3000
      TabIndex        =   89
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   9
      IntegralHeight  =   0   'False
      Left            =   4800
      TabIndex        =   88
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   8
      IntegralHeight  =   0   'False
      Left            =   4920
      TabIndex        =   87
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   7
      IntegralHeight  =   0   'False
      Left            =   5040
      TabIndex        =   86
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   6
      IntegralHeight  =   0   'False
      Left            =   4920
      TabIndex        =   85
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   5
      IntegralHeight  =   0   'False
      Left            =   4800
      TabIndex        =   84
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   9
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   83
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   8
      IntegralHeight  =   0   'False
      Left            =   3360
      TabIndex        =   82
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   7
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   81
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   6
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   80
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   5
      IntegralHeight  =   0   'False
      Left            =   3000
      TabIndex        =   79
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox cbi 
      Height          =   135
      Left            =   2040
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   78
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      Picture         =   "Form1.frx":0CCE
      Style           =   1  'Grafisch
      TabIndex        =   76
      ToolTipText     =   "In Taskleiste minimieren"
      Top             =   3420
      Width           =   375
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Listen"
      Height          =   375
      Left            =   8040
      TabIndex        =   75
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox altsuch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      IntegralHeight  =   0   'False
      Left            =   9960
      TabIndex        =   74
      ToolTipText     =   "Suche in allen Feldern"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Picture         =   "Form1.frx":139C
      Style           =   1  'Grafisch
      TabIndex        =   73
      ToolTipText     =   "Neuen Termin erstellen"
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox cb2 
      Height          =   135
      Left            =   2040
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   72
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox cb1 
      Height          =   135
      Left            =   1800
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   71
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2520
      Picture         =   "Form1.frx":172E
      Style           =   1  'Grafisch
      TabIndex        =   69
      ToolTipText     =   "Kalender öffnen"
      Top             =   2760
      Width           =   375
   End
   Begin VB.ListBox deldoclist 
      Height          =   450
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   68
      Top             =   5520
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   240
      Picture         =   "Form1.frx":18B8
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   67
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   4
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   66
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   3
      IntegralHeight  =   0   'False
      Left            =   3360
      TabIndex        =   65
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   64
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   63
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   4
      IntegralHeight  =   0   'False
      Left            =   4680
      TabIndex        =   62
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   3
      IntegralHeight  =   0   'False
      Left            =   4560
      TabIndex        =   61
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   4440
      TabIndex        =   60
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   4320
      TabIndex        =   59
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox sortlist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      IntegralHeight  =   0   'False
      Left            =   8880
      Sorted          =   -1  'True
      TabIndex        =   58
      ToolTipText     =   "Alles für heute Relevante aus dem Kalender"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2880
      Picture         =   "Form1.frx":3ABA
      Style           =   1  'Grafisch
      TabIndex        =   57
      ToolTipText     =   "Geburtstagsliste"
      Top             =   3120
      Width           =   375
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   3960
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2760
      Picture         =   "Form1.frx":3C44
      Style           =   1  'Grafisch
      TabIndex        =   53
      ToolTipText     =   "Ihr Dokumentenverzeichnis öffnen"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Height          =   375
      Left            =   2760
      Picture         =   "Form1.frx":426E
      Style           =   1  'Grafisch
      TabIndex        =   47
      ToolTipText     =   "Mailsafe"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9000
      Picture         =   "Form1.frx":43F8
      Style           =   1  'Grafisch
      TabIndex        =   54
      ToolTipText     =   "Schliesst alle Formulare"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6360
      Picture         =   "Form1.frx":4A22
      Style           =   1  'Grafisch
      TabIndex        =   51
      ToolTipText     =   "gespeicherte Selektionen"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   7680
      Picture         =   "Form1.frx":4BAC
      Style           =   1  'Grafisch
      TabIndex        =   49
      ToolTipText     =   "Saalpläne"
      Top             =   4440
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
      Left            =   600
      TabIndex        =   48
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox pin 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "erweitert"
      Height          =   195
      Left            =   3120
      TabIndex        =   45
      ToolTipText     =   "Auch in Hinweisen suchen"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   4080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":500E
      Style           =   1  'Grafisch
      TabIndex        =   44
      ToolTipText     =   "InternetBrowser öffnen"
      Top             =   2160
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.PictureBox mlstat 
      Height          =   375
      Index           =   2
      Left            =   3000
      Picture         =   "Form1.frx":5292
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   40
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox mlstat 
      Height          =   375
      Index           =   1
      Left            =   2280
      Picture         =   "Form1.frx":5C58
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox mlstat 
      Height          =   375
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":661E
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":6FE4
      Style           =   1  'Grafisch
      TabIndex        =   37
      ToolTipText     =   "Email empfangen"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2520
      Picture         =   "Form1.frx":79AA
      Style           =   1  'Grafisch
      TabIndex        =   36
      ToolTipText     =   "Ihre Dokumente anzeigen"
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":7E5D
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":89DF
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1800
      Picture         =   "Form1.frx":9561
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "Cross Over-Projekte"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1320
      Picture         =   "Form1.frx":965E
      Style           =   1  'Grafisch
      TabIndex        =   32
      ToolTipText     =   "Kammermusik-Projekte"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1800
      Picture         =   "Form1.frx":9A6A
      Style           =   1  'Grafisch
      TabIndex        =   30
      ToolTipText     =   "Cross Over-Tourneeangebote"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1320
      Picture         =   "Form1.frx":9B67
      Style           =   1  'Grafisch
      TabIndex        =   29
      ToolTipText     =   "Kammermusik-Tourneeangebote"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      Picture         =   "Form1.frx":9F73
      Style           =   1  'Grafisch
      TabIndex        =   27
      ToolTipText     =   "Künstler-Tourneeangebote"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2880
      Picture         =   "Form1.frx":A330
      Style           =   1  'Grafisch
      TabIndex        =   25
      ToolTipText     =   "Kalender öffnen"
      Top             =   2760
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Left            =   2280
      Top             =   2280
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8640
      Picture         =   "Form1.frx":A430
      Style           =   1  'Grafisch
      TabIndex        =   24
      ToolTipText     =   "Verwaltungsfunktionen"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":A4CF
      Style           =   1  'Grafisch
      TabIndex        =   18
      ToolTipText     =   "To Do-Liste öffnen"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   59000
      Left            =   240
      Top             =   0
   End
   Begin VB.ListBox waehrungen 
      Height          =   450
      Left            =   3960
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox todo2 
      Height          =   450
      Index           =   0
      Left            =   4200
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox todo 
      Height          =   450
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   360
      Picture         =   "Form1.frx":A575
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Künstler-Projekte"
      Top             =   2160
      Width           =   375
   End
   Begin VB.ListBox dochistlist 
      Height          =   255
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7560
      Top             =   120
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Programme"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Programme öffnen"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Termine u. &Auftritte"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Termine und Auftritte öffnen"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   840
      Picture         =   "Form1.frx":A932
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Orchester-Projekte"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   840
      Picture         =   "Form1.frx":A9A3
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Orchester-Tourneeangebote"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Werke"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Werkeverzeichnis öffnen"
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   6360
      TabIndex        =   4
      ToolTipText     =   "Suche nach Kontaktpersonen"
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "Form1.frx":AA14
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Auf Wiedersehen!"
      Top             =   3840
      Width           =   375
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Suche in allen Feldern"
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton sqlmess 
      Caption         =   "SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton errmess 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Fehler"
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
      MaskColor       =   &H0080C0FF&
      TabIndex        =   19
      ToolTipText     =   "Achtung! Fehler"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":AC64
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Email schreiben"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Adressen"
      Height          =   735
      Left            =   1320
      Picture         =   "Form1.frx":AD16
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "Formular Adressen öffnen"
      Top             =   240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   9360
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   109
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "queued:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   106
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Replication: none"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   105
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet: "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1200
      TabIndex        =   77
      Top             =   3780
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Prioritäten"
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
      Left            =   1320
      TabIndex        =   70
      ToolTipText     =   "Alle Projekte, Doppelklick öffnet Projektübersicht"
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label fallbackq 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   56
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label uuid 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "UserID"
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
      Height          =   255
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Ihre Benutzerkennung"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   52
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image outlk 
      Height          =   375
      Left            =   7800
      Picture         =   "Form1.frx":B21A
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label pinlbl 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
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
      Left            =   8040
      TabIndex        =   46
      ToolTipText     =   "Aktuelles Datum mit Uhrzeit"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label isregdat 
      Caption         =   "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   41
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label isregdat 
      Caption         =   "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label isregdat 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "  Projekte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Alle Projekte, Doppelklick öffnet Projektübersicht"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   6840
      Picture         =   "Form1.frx":BB8C
      ToolTipText     =   "In Kontakten suchen"
      Top             =   240
      Width           =   315
   End
   Begin VB.Image Image5 
      Height          =   330
      Left            =   2760
      Picture         =   "Form1.frx":BC5F
      ToolTipText     =   "In Adressen suchen"
      Top             =   240
      Width           =   345
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "   Tourneeangebote"
      Height          =   735
      Left            =   240
      TabIndex        =   28
      ToolTipText     =   "Alle Tourneeangebote"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "datum_uhr"
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
      TabIndex        =   17
      ToolTipText     =   "Aktuelles Datum mit Uhrzeit"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Heute aktuell"
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
      Left            =   3240
      TabIndex        =   22
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Suche in Kontakten:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      ToolTipText     =   "In Kontakten suchen"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Suchen"
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
      Left            =   3240
      TabIndex        =   21
      ToolTipText     =   "In Adressen suchen"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4215
      Left            =   2520
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   4215
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Wiederherstellen"
         Enabled         =   0   'False
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&beenden"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu dat 
      Caption         =   "&Datei"
      Begin VB.Menu edt_myfiles 
         Caption         =   "&Eigene Agencyprofdateien"
         Shortcut        =   ^O
      End
      Begin VB.Menu ruler1 
         Caption         =   "----------"
         Enabled         =   0   'False
      End
      Begin VB.Menu dat_clsall 
         Caption         =   "&Alle Fenster schließen"
         Shortcut        =   ^H
      End
      Begin VB.Menu dat_end 
         Caption         =   "&beenden"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu edt 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu edt_adrsuch 
         Caption         =   "&Adressen suchen"
         Shortcut        =   ^A
      End
      Begin VB.Menu edt_prio 
         Caption         =   "&Prioritäten"
         Shortcut        =   ^P
      End
      Begin VB.Menu eml 
         Caption         =   "&Email ..."
         Begin VB.Menu edt_erecv 
            Caption         =   "&empfangen"
            Shortcut        =   ^R
         End
         Begin VB.Menu edt_esend 
            Caption         =   "&senden"
            Shortcut        =   ^S
         End
         Begin VB.Menu edt_msafe 
            Caption         =   "&Mailsafe"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu edt_st 
         Caption         =   "Eins&tellungen ..."
         Begin VB.Menu edt_tools 
            Caption         =   "&Werkzeuge"
            Shortcut        =   ^T
         End
         Begin VB.Menu edt_set 
            Caption         =   "&Benutzerinstellungen"
            Shortcut        =   ^E
         End
      End
   End
   Begin VB.Menu trmn 
      Caption         =   "&Termine"
      Begin VB.Menu ta 
         Caption         =   "&Angebote ..."
         Begin VB.Menu ta_kstl 
            Caption         =   "&Künstler"
         End
         Begin VB.Menu ta_orch 
            Caption         =   "&Orchester"
         End
         Begin VB.Menu ta_chamb 
            Caption         =   "Ka&mmermusik"
         End
         Begin VB.Menu ta_cross 
            Caption         =   "&Crossover"
         End
      End
      Begin VB.Menu prj 
         Caption         =   "&Projekte ..."
         Begin VB.Menu prj_kstl 
            Caption         =   "&Künstler"
         End
         Begin VB.Menu prj_orch 
            Caption         =   "&Orchester"
         End
         Begin VB.Menu prj_chamb 
            Caption         =   "Ka&mmermusik"
         End
         Begin VB.Menu prj_cross 
            Caption         =   "&Crossover"
         End
         Begin VB.Menu ruler2 
            Caption         =   "----------"
            Enabled         =   0   'False
         End
         Begin VB.Menu prj_cal 
            Caption         =   "Ü&bersicht"
         End
      End
      Begin VB.Menu trmne 
         Caption         =   "&Termine ..."
         Begin VB.Menu trmn_dayvw 
            Caption         =   "&Tageskalender"
         End
         Begin VB.Menu trmn_cal 
            Caption         =   "&Kalenderübersicht"
         End
         Begin VB.Menu trmn_list 
            Caption         =   "&Terminliste"
         End
         Begin VB.Menu trm_todolist 
            Caption         =   "To&Do-Liste"
            Shortcut        =   ^D
         End
         Begin VB.Menu trmn_akt 
            Caption         =   "&Heute aktuell"
         End
         Begin VB.Menu trmn_brthd 
            Caption         =   "&Geburtstagsliste"
            Shortcut        =   ^G
         End
         Begin VB.Menu ruler4 
            Caption         =   "----------"
         End
         Begin VB.Menu trmn_new 
            Caption         =   "&neuer Termin"
            Shortcut        =   ^N
         End
         Begin VB.Menu trm_todo 
            Caption         =   "n&eues ToDo"
         End
      End
   End
   Begin VB.Menu wrk 
      Caption         =   "&Werkeverzeichnis"
      Begin VB.Menu wrk_open 
         Caption         =   "&Werke"
      End
      Begin VB.Menu wrk_prgopn 
         Caption         =   "&Programme"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "&?"
      Begin VB.Menu hlp_help 
         Caption         =   "&Hilfe"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindow Lib "user32" _
  (ByVal hWnd As Long, ByVal wCmd As Long) As Long
  
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Public dochistisopen As Boolean, auftrittisopen As Boolean, todoisopen As Boolean, brwhidden As Boolean
Public adoc As ADODB.Connection, geodbok As Boolean, kaldbok As Boolean, tzoffset As Long
Public cloud As Boolean, cloudserver$, clouduser$, cloudpass$, hordexlock As Boolean
Public kaldb As ADODB.Connection, geodb As ADODB.Connection, cloudmanager As String, cloudstaff As String, supershares_krono As String
Public alertdbo As ADODB.Connection, alertdbok As Boolean, isp3home As String
Public dttrenn As String, odbcdriver As String, dochistlock As Boolean
Public lnkcolor As Long, wawipara$, ttmode As String, ostype
Dim shwled As Boolean, ismin As Boolean, warnmeondata As Boolean
Public clddb As ADODB.Connection, tlopen As Boolean
Dim use_adrsuchtimer As Boolean, usealtsuch As Boolean, dir_mailoutbox As String
Public ihavemail As Boolean, alertdbuid As String, hppth$, aplibok As Boolean, libist As Long, libsoll As Long
Dim wrkJet As Workspace, granttab(199) As String, grantptr As Integer
Public pwentered As String, usemenu As String, noalarms As Boolean, alertdbpara$
Public sqla As Database, currentlanguage As String, uselimitinsql As Boolean, weckerpresent As Boolean
Dim adopara$, starting As Boolean, missingfields As String, internalkey As String
Dim dbname$, dbpara$, dbpsswd$, dbserver$, fallbackserver$, dbuid$, dbport$
Public fallbackserverpath$, dbfpara$, menoquit As Boolean, globcount As Long, poplock As Boolean
Public pub_sqla As Database, meinesprache, tdlistisopen As Boolean
Public dbg2file%, usrprofile$, msec As Double
Dim uId$, ufaxrtf$, ubrfrtf$, ufax$, ueditor$, uclog$, usuchvz As Double, doc0dir$
Dim ubrowse$, umysqld$, udochiscore%, udochiscore1%, uemlhiscore%, uinbhiscore%, uwantstooltips%, umysqlhost$, uwpool$
Dim uphedit$, netscape47inbox$, mailclient$, usavealways$, ucalalways$, ufdow$, uhomepath$
Dim umsrvout$, umailadr$, uexpi$, umchk$, ufsze%, m0dir$
Dim upop$, upopid$, upoppsswd$, upopport%, umsgwait%, AuftrittsdruckFuerAdresse$
Dim bkmstart$, bkmend$, snotb4 As Double, docdup$, docequiv1$, docequiv2$, ttabptr%, spmlst$(99)
Dim dayname(1 To 7), longdayname(1 To 7), break%, pspath$, backslashhandler
Dim bkmlist$(1, 299), bkmlcount%(1), bkmrflag$(299)
Dim mwst As Double, provision As Double, ichspreche As String
Dim Honorarliste$(6, 499), Honorarwaehrung$(499)
Public currentconfmode$
Public listenhauptperson, dbpasswd As String
Public honorarlcount%, computername$, myip$, mydemoid$, dbswitch As Boolean, skip1del As Boolean
Dim datchgmode As String, crlfrepl As String
Public mustfield$, fastsave_copy As Boolean
Dim honorarkurs$(499), SelectedDate$, selectedcolor As Long, dirtcol As Long
Dim honorarsumme1_brutto$, honorarsumme1_netto$, mwstsumme$, honvalid, hontrue As Boolean
Dim provisionssumme_brutto$, provisionssumme_netto$, provisionssumme_mwst$, honerr$
Dim memono%, t3tick%, t2tick%
Dim noliccnt%, ehsc%, errsh%, fallbackdir$
Dim provtyp%      '0=fix, 1=%
Dim colorcacheid$(99), colorcache%(99, 2), colorcachepointer%
Dim atabkzcacheid$(99), atabkz$(99), atabkzcachepointer%
Public d0t0y As Long, d0t0m As Long
Public sys_mwst As Double, autocheckmail As Boolean
Dim thismwst As Double     'mwst für provisionsrechnung
Dim d1t0y As Long, d1t0m As Long
Dim aKey() As Byte, mycookie$, s0d$, s00d$
Dim exceldelim$, convertcolor As Long
Public vorlagencache As String
Public adrmerkid$, err_dupok%, uname$, kalopen As Boolean, dayvopen As Boolean, priosopen As Boolean, anredeuser$
Dim aliasfeld$(99, 2), aliastext$(99, 2), adrfeldcache$(199)
Dim statusfarbe(9) As Long, statusname$(9), mnams$(1 To 12), mnams_engl$(1 To 12)
Dim transtab(2, 1500), auftrittsdruck_currvorlage$, auftrittsdruck_currfeld$
Dim h_netto As Double, h_mwst As Double, h_brutto As Double
Dim srchhigru As String
Dim POPTaskID As Long, localdir As String
Dim alertdb As String, alertdbuser As String, alertdbpsswd As String, alertdbhost As String
Dim adruckmerkwert(10) As String
Dim usr_setting(0 To 1, 0 To 199) As String, useusrcache As String, usr_set_hits(199) As Integer
Dim xkurs_publicratestring
Dim adostats_tsum As Long, adostats_samples As Long
      
Dim dbg_prvtlnkcount As String
      
      'constants required by Shell_NotifyIcon API call:
      Const NIM_ADD = &H0
      Const NIM_MODIFY = &H1
      Const NIM_DELETE = &H2
      Const NIF_MESSAGE = &H1
      Const NIF_ICON = &H2
      Const NIF_TIP = &H4
      Const WM_MOUSEMOVE = &H200
      Const WM_LBUTTONDOWN = &H201     'Button down
      Const WM_LBUTTONUP = &H202       'Button up
      Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Const WM_RBUTTONDOWN = &H204     'Button down
      Const WM_RBUTTONUP = &H205       'Button up
      Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hWnd As Long) As Long
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      
      Private Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hicon As Long
       szTip As String * 64
      End Type

      Private nid As NOTIFYICONDATA


Public Function s0dir() As String

s0dir = s0d$

End Function

Public Function s00dir() As String

s00dir = s00d$

End Function

Public Function wavdir() As String

wavdir = ichspreche

End Function

Function getuseremail(u$) As String
Dim r As ADODB.Recordset, cmd$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getuseremail"
getuseremail = ""

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT email FROM benutzerdaten where id='" + u$ + "'"
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
If r.EOF Then Exit Function
If Not IsNull(r!email) Then getuseremail = r!email
r.Close
End If

End Function
Sub setdbpara(dbn$, dbp$, hpt$, dbs$, adop$, wawip$)
Dim p%

dbname$ = dbn$
dbpara$ = dbp$
dbserver$ = dbs$
p% = InStr(dbserver$, ":")
dbport$ = "3306"
If p% > 0 Then
  dbport$ = Mid$(dbserver$, p% + 1)
  dbserver$ = Left(dbserver$, p% - 1)
End If
adopara$ = adop$
wawipara$ = wawip$
If InStr(hpt$, "\") > 0 Then uhomepath$ = hpt$

End Sub

Public Function new2do(von$, an$, betreff$, nachricht$, trgdatum$, trgzeit$, perdata%, perunit$, pmal%) As String
Dim id$, trgf$, neubetreff$, abetr$, f$, p_in%, p%, sid$

'd2infile = "Form1": d2insub = "new2do"
id$ = GUID()
neubetreff$ = betreff$
new2do = id$
f$ = betreff$
p_in% = InStr(f$, "[Wiedervorlage] Voicemail:")
If p_in% > 0 Then
  abetr$ = ""
  If p_in% > 1 Then abetr$ = Left$(f$, p_in% - 1)
  f$ = Mid$(f$, p_in% + Len("[Wiedervorlage] Voicemail:"))
  sid$ = Mid$(f$, p% + 1)
  If exist(sid$) <> 0 Then
    trgf$ = myuniqueinboxname() & ".wav"
    Call FileCopy(sid$, trgf$)
    neubetreff$ = abetr$ & "[Wiedervorlage] voicemail:" & trgf$
  End If
End If
Call sqlqry("insert into todolist (id) values('" & id$ & "');")
Call sqlqry("update todolist set von='" & von$ & "' where id='" & id$ & "';")
Call sqlqry("update todolist set an='" & an$ & "' where id='" & id$ & "';")
Call sqlqry("update todolist set betreff='" & neubetreff$ & "' where id='" & id$ & "';")
Call sqlqry("update todolist set nachricht='" & trm(nachricht$) & " ' where id='" & id$ & "';")
Call sqlqry("update todolist set status='neu' where id='" & id$ & "';")
Call sqlqry("update todolist set datum='" & trgdatum$ & "' where id='" & id$ & "';")
Call sqlqry("update todolist set zeit='" & trgzeit$ & "' where id='" & id$ & "';")
Call sqlqry("update todolist set pdelta='" & trm(perdata%) & "' where id='" & id$ & "';")
Call sqlqry("update todolist set poft='" & trm(perdata%) & "' where id='" & id$ & "';")
'sqlqry ("insert into todolist (id,von,an,betreff,nachricht,status,datum,zeit,pdelta,poft) " + _
'       "values('" + id$ + "'," + _
'              "'" + von$ + "'," + _
'              "'" + an$ + "'," + _
'              "'" + neubetreff$ + "'," + _
'              "'" + nachricht$ + "'," + _
'              "'neu'," + _
'              "'" + trgdatum$ + "'," + _
'              "'" + trgzeit$ + "'," + _
'               trm(perdata%) + "," + _
'               trm(perdata%) + ")")

If perunit$ <> "" Then sqlqry ("update todolist set pdeltaunit='" & perunit$ & "' where id='" & id$ & "';")
If tdlistisopen Then Call todolist.Command4_Click

End Function
Public Function num1(l$)
Dim f%, i%

f% = 0
For i% = 1 To Len(l$)
  If InStr("1234567890.,", Mid$(l$, i%, 1)) = 0 Then
    f% = i%
    i% = Len(l$)
  End If
Next i%

If f% > 0 Then
  num1 = Left$(l$, f% - 1)
Else
  num1 = l$
End If
End Function

Private Sub recalc_honorarliste(cmwst%)
Dim s As Double, i%, rrr
Dim prov As Double, cmws As Double, provstr$, amws As Double
Dim k As Double, pn As Double, pm As Double
Dim s1 As Double, formst$, provmwst As Double

'd2infile = "Form1": d2insub = "recalc_honorarliste"
If honvalid = 1 Then Exit Sub
h_netto = 0: h_mwst = 0: h_brutto = 0
s = 0: pn = 0: pm = 0
cmws = var2dbl(cmwst%) / 100 / 100
prov = provision
For i% = 0 To honorarlcount% - 1
  amws = var2dbl(Val(Honorarliste$(4, i%))) / 100 / 100
  Call dbg2f("recalc_honorarliste: i=" & str$(i%) & " 0(honorar)=" & Honorarliste$(0, i%) & " 1=" & Honorarliste$(1, i%) & " 2=" & Honorarliste$(2, i%) & " 3=" & Honorarliste$(3, i%) & " 4=" & Honorarliste$(4, i%) & " 5=" & Honorarliste$(5, i%))
  If Honorarliste$(0, i%) <> "" Then
    provmwst = var2dbl(Honorarliste$(6, i%)) * 100
    Honorarliste$(2, i%) = nurdiewaehrung(Honorarliste$(0, i%))
    honorarkurs$(i%) = kursvom(Honorarliste$(2, i%), Honorarliste$(5, i%))
    k = var2dbl(strrepl(honorarkurs$(i%), ".", ","))
    If k = 0 Then k = 1000000
    On Error Resume Next
    s1 = CCur(ohnewaehrung(Honorarliste$(0, i%))) / k
    rrr = Err
    On Error GoTo 0
    s = s + s1
    If rrr <> 0 Then
      hontrue = False
      honerr$ = "mindestens 1 Honorar ungültig am " + Honorarliste$(5, i%)
    End If
    On Error Resume Next
    Honorarliste$(1, i%) = CCur(ohnewaehrung(Honorarliste$(0, i%))) / k
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      hontrue = False
      honerr$ = "mindestens 1 Honorar ungültig am " + Honorarliste$(5, i%)
    End If
    Call dbg2f("recalc_honorarliste: i=" & str$(i%) & " 0=" & Honorarliste$(0, i%) & " 1=" & Honorarliste$(1, i%) & " 2=" & Honorarliste$(2, i%) & " 3=" & Honorarliste$(3, i%) & " 4=" & Honorarliste$(4, i%) & " 5=" & Honorarliste$(5, i%))
'Debug.Print "recalc_honorarliste: i=" & str$(i%) & " 0=" & Honorarliste$(0, i%) & " 1=" & Honorarliste$(1, i%) & " 2=" & Honorarliste$(2, i%) & " 3=" & Honorarliste$(3, i%) & " 4=" & Honorarliste$(4, i%) & " 5=" & Honorarliste$(5, i%)
    provstr$ = trm(Honorarliste$(3, i%))
    provtyp = 0: If Right$(provstr$, 1) = "%" Then provtyp = 1
    h_netto = h_netto + s1
    h_mwst = h_mwst + s1 * amws
    h_brutto = h_netto + h_mwst
    If provtyp = 1 Then
      prov = var2dbl(trm(Left(provstr$, Len(provstr$) - 1))) / 100
      pn = pn + ((s1 + s1 * amws) * prov) / k
      pm = pm + (((s1 + s1 * amws) * prov) * (provmwst / 10000)) / k
      provisionssumme_netto$ = fixeur(pn)
      provisionssumme_mwst$ = fixeur(pm)
      provisionssumme_brutto$ = fixeur(pn + pm)
    Else
      On Error Resume Next
      prov = CCur(ohnewaehrung(Honorarliste$(3, i%)))
      rrr = Err
      On Error GoTo 0
      If rrr <> 0 Then prov = 0
      pn = pn + prov / k
      pm = pm + prov * (provmwst / 10000) / k
      provisionssumme_netto$ = fixeur(pn)
      provisionssumme_mwst$ = fixeur(pm)
      provisionssumme_brutto$ = fixeur(pn + pm)
    End If
  End If
Next i%
formst$ = "0.00"
If hontrue Then
  honorarsumme1_netto$ = fixeur(h_netto)
  honorarsumme1_brutto$ = fixeur(h_brutto)
  mwstsumme$ = fixeur(h_mwst)
Else
  honorarsumme1_netto$ = honerr$
  honorarsumme1_brutto$ = honerr$
  mwstsumme$ = honerr$
  provisionssumme_netto$ = honerr$
  provisionssumme_mwst$ = honerr$
  provisionssumme_brutto$ = honerr$
  provisionssumme_netto$ = honerr$
  provisionssumme_mwst$ = honerr$
  provisionssumme_brutto$ = honerr$
End If
honvalid = 1

End Sub

Sub hgradrsuch(s$)
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, rrr, cmd$, i9%, na$, l$
Dim nsuch As String, addit As Boolean, rcnt%, whkrit$, id$
Dim w1 As String, rest As String, p1%, ktid As String

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "hgradrsuch"
'On Error GoTo errhdl

If poplock Then Exit Sub

cmd$ = "SELECT LCase(FeldDaten) AS flddat, auftritthigru.auftrittsid as aid, auftritthigru.auftrittstyp, auftritthigru.FeldName, adresse.ID as adrid " + _
       "FROM auftritthigru, adresse " + _
       "WHERE instr(LCase(FeldDaten),'" + LCase(s$) + "')>0 AND InStr(auftrittsid,adresse.id)=1;"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
rcnt% = 0
If Not rtmp.EOF And rrr = 0 Then
  rtmp.MoveFirst
  While Not rtmp.EOF And break% = 0 And rcnt% < 50
    w1 = Mid(trm(rtmp!aid), Len(rtmp!adrid) + 1)
    If w1 = "" Then
      cmd$ = "select id,name,' ' as position from adresse where id='" + rtmp!aid + "'"
    Else
      cmd$ = "select id,name,position from kontakt where id='" + w1 + "'"
    End If
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF And rrr = 0 Then
      If w1 <> "" Then
        na$ = trm(r!name)
        If trm(r!Position) <> "" Then na$ = na$ + " (" + trm(r!Position) + ")"
        na$ = na$ + " (" + rtmp!auftrittstyp + ", " + rtmp!feldname + ")"
      Else
        na$ = trm(r!id)
        na$ = na$ + "::" + rtmp!auftrittstyp + ", " + rtmp!feldname
      End If
      If w1 = "" Then
        List1.AddItem trm(crlffake(na$) + Space$(80) + "ID:" + r!id)
      Else
        List2.AddItem crlffake(na$) + Space$(80) + "ID:" + r!id
      End If
    End If
    rtmp.MoveNext
    DoEvents
    If break% > 0 Then
      break% = 0
      Exit Sub
    End If
  Wend
End If
rtmp.Close
End Sub

Sub rlist1(s$)
Dim rtmp As ADODB.Recordset, rrr, cmd$, i9%, na$, l$
Dim nsuch As String, addit As Boolean, rcnt%, whkrit$, id$
Dim w1 As String, rest As String, p1%, ktid As String, als$, ij%
Dim s1$, ca$, cb$, z$, swrd$(6), i%
Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "rlist1"
'On Error GoTo errhdl

If poplock Then Exit Sub
List1.Clear
List2.Clear
nsuch = getusersetting("adresstypnichtsuchen", "")
shwAdrDetail.Combo3.Clear

s$ = strrepl(s$, "'", "´")
s1$ = LCase(s$)
cmd$ = "SELECT * FROM adresse where (instr(telfaxhandy,'" + s$ + "')>0) or "
If getusersetting("adressensuchen", "erweitert") = "erweitert" Then
  For i% = 0 To 4: swrd$(i%) = "": Next i%
  i% = 0
  While Len(s1$) > 0 And i% < 5
    swrd$(i%) = word1(s1$)
    s1$ = word2bis(s1$)
    i% = i% + 1
  Wend
  cb$ = ""
  If Len(s1$) > 0 Then swrd$(i%) = s1$
  For i% = 0 To 5
    If swrd$(i%) <> "" Then
      If cb$ <> "" Then cb$ = cb$ + " and "
      If Check1.value = 1 Then
        cb$ = cb$ + "( (" + _
           "instr(lcase(strasse),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(ort),'" + swrd$(i%) + "')>0) or (" + _
           "instr(plz,'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(id),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(url),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(hinweise),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(email),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(name),'" + swrd$(i%) + "')>0) )"
      Else
        cb$ = cb$ + "( (" + _
           "instr(lcase(strasse),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(ort),'" + swrd$(i%) + "')>0) or (" + _
           "instr(plz,'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(id),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(url),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(email),'" + swrd$(i%) + "')>0) or (" + _
           "instr(lcase(name),'" + swrd$(i%) + "')>0) )"
      End If
    End If
  Next i%
  ca$ = " ( " + cb$ + " ) "
Else
  If Check1.value = 1 Then
    ca$ = "(" + _
       "instr(lcase(strasse),'" + s1$ + "')>0) or (" + _
       "instr(lcase(ort),'" + s1$ + "')>0) or (" + _
       "instr(plz,'" + s1$ + "')>0) or (" + _
       "instr(lcase(id),'" + s1$ + "')>0) or (" + _
       "instr(lcase(url),'" + s1$ + "')>0) or (" + _
       "instr(lcase(hinweise),'" + s1$ + "')>0) or (" + _
       "instr(lcase(email),'" + s1$ + "')>0) or (" + _
       "instr(lcase(name),'" + s1$ + "')>0)"
  Else
    ca$ = "(" + _
       "instr(lcase(strasse),'" + s1$ + "')>0) or (" + _
       "instr(lcase(ort),'" + s1$ + "')>0) or (" + _
       "instr(plz,'" + s1$ + "')>0) or (" + _
       "instr(lcase(id),'" + s1$ + "')>0) or (" + _
       "instr(lcase(url),'" + s1$ + "')>0) or (" + _
       "instr(lcase(email),'" + s1$ + "')>0) or (" + _
       "instr(lcase(name),'" + s1$ + "')>0)"
  End If
End If
If getusersetting("erweiterteumlautsuche", "ja") = "ja" Then
  cb$ = ca$
  z$ = "ä": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ae")
  z$ = "ö": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "oe")
  z$ = "ü": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ue")
  z$ = "ß": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ss")
  ca$ = cb$
End If
cmd$ = cmd$ + ca$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
rcnt% = 0
If Not rtmp.EOF And rrr = 0 Then
  rtmp.MoveFirst
  While Not rtmp.EOF And break% = 0 And rcnt% < 50
    addit = True
    If nsuch <> "" Then
      rest = trm(nsuch)
      While rest <> "" And addit
        w1 = cut_d1(rest, "|")
        rest = cut_d2bis(rest, "|")
        If isoftype(rtmp!id, w1) <> "-1" Then addit = False
      Wend
    End If
    If addit Then
      rcnt% = rcnt% + 1
      If trm(rtmp!id) <> trm(rtmp!name) Then
        als$ = rtmp!id & "::" + rtmp!name
      Else
        als$ = rtmp!id
      End If
      Call l1a(als$)
      For ij% = 0 To altsuch.ListCount - 1
        If altsuch.List(ij%) = als$ Then Exit For
      Next ij%
      If ij% >= altsuch.ListCount Then altsuch.AddItem als$
    End If
    shwAdrDetail.Combo3.AddItem rtmp!id
    rtmp.MoveNext
    DoEvents
    If break% > 0 Then
      break% = 0
      Exit Sub
    End If
  Wend
End If
rtmp.Close
rcnt% = 0
cmd$ = "SELECT ID,name,position FROM kontakt where ((instr(telfaxhandy,'" + s$ + "')>0) ) or "

If getusersetting("adressensuchen", "erweitert") = "erweitert" Then
  cb$ = ""
  For i% = 0 To 5
    If swrd$(i%) <> "" Then
      If cb$ <> "" Then cb$ = cb$ + " and "
      cb$ = cb$ + "( (instr(lcase(name),'" + swrd$(i%) + "')>0) or " + _
                  " (instr(email,'" + swrd$(i%) + "')>0) )"
    End If
  Next i%
  ca$ = cb$
Else
  ca$ = "( (instr(lcase(name),'" + LCase(s$) + "')>0) or " + _
         " (instr(email,'" + LCase(s$) + "')>0) )"
End If
If getusersetting("erweiterteumlautsuche", "ja") = "ja" Then
  cb$ = ca$
  z$ = "ä": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ae")
  z$ = "ö": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "oe")
  z$ = "ü": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ue")
  z$ = "ß": If InStr(ca$, z$) > 0 Then cb$ = cb$ + " or " + strrepl(ca$, z$, "ss")
  ca$ = cb$
End If
cmd$ = cmd$ + "(" + ca$ + ")"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not rtmp.EOF Then
  rtmp.MoveFirst
  i9% = 0
  While Not rtmp.EOF And i9% < 99 And break% = 0 And rcnt% < 50
    na$ = trm(rtmp!name)
    ktid = getadridbykontaktid(rtmp!id)
    addit = True
    If nsuch <> "" Then
      rest = trm(nsuch)
      While rest <> "" And addit
        w1 = cut_d1(rest, "|")
        rest = cut_d2bis(rest, "|")
        If isoftype(na$ + "{" + ktid + "}", w1) <> "-1" Then addit = False
      Wend
    End If
    If addit Then
      rcnt% = rcnt% + 1
      i9% = i9% + 1
      If trm(rtmp!Position) <> "" Then na$ = na$ + " (" + rtmp!Position + ")"
      List2.AddItem form1.crlffake(na$) & Space$(80) & "ID:" & rtmp!id
    End If
    rtmp.MoveNext
    DoEvents
    If break% > 0 Then
      break% = 0
      Exit Sub
    End If
  Wend
End If
rtmp.Close
If Check1.value = 1 And srchhigru <> "" Then
  l$ = srchhigru
  whkrit$ = ""
  While l$ <> ""
    na$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
    If whkrit$ <> "" Then whkrit$ = whkrit$ + " or "
    whkrit$ = whkrit$ + "Feldname='" + na$ + "'"
  Wend
  cmd$ = "SELECT adresse.ID,adresse.name FROM adresse INNER JOIN auftritthigru ON adresse.ID = auftritthigru.auftrittsid where "
  cmd$ = cmd$ + "((" + whkrit$ + ") and instr(lcase(felddaten),'" + LCase(s$) + "')>0)"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

  If Not rtmp.EOF Then
    While Not rtmp.EOF And i9% < 99 And break% = 0 And rcnt% < 50
      na$ = trm(rtmp!name)
      id$ = trm(rtmp!id)
      rcnt% = rcnt% + 1
      i9% = i9% + 1
      Call l1a(trm(rtmp!id) + "::" + na$)
      rtmp.MoveNext
      DoEvents
      If break% > 0 Then
        break% = 0
        Exit Sub
      End If
    Wend
  End If
  rtmp.Close
End If
If usealtsuch Then
  If altsuch.ListCount > 0 Then
    For ij% = 0 To altsuch.ListCount - 1
      For i% = 0 To List1.ListCount - 1
        If List1.List(i) = altsuch.List(ij%) Then Exit For
      Next i%
      If i% >= List1.ListCount Then Call l1a(altsuch.List(ij%))
    Next ij%
  End If
End If
List1.AddItem "--" & form1.inmylanguage("Aktuelle Adressen") & "--"
If Check1.value = 1 Then Call hgradrsuch(s$)
Call showprios
End Sub

Private Sub Beenden_Click()
Call Command1_Click
End Sub

Private Sub commannd19_updtooltip()
Dim tx$

tx$ = trm(altbvorl.text)
If tx$ = "" Then tx$ = "Ihre Dokumente"
Command19.ToolTipText = tx$ & " " + transe("im Exporer öffnen")
End Sub

Private Sub altbvorl_Change()

Call commannd19_updtooltip

End Sub

Private Sub altbvorl_Click()
Call commannd19_updtooltip
End Sub

Private Sub altbvorl_DropDown()
Dim rtmp As ADODB.Recordset, drvl As String
Dim uId$, rrr, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "altbvorl_DropDown"
altbvorl.Clear

uId$ = form1.getuserid()
altbvorl.AddItem s00d$ & "\" & docs() & "\" & uId$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
cmd$ = "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_DokumentenVerzeichnis')>0 or instr(owner,'sysvar_system_DokumentenVerzeichnis')>0"
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
While Not rtmp.EOF
  'o$ = rtmp!Owner
  'o$ = Mid$(o$, InStr(o$, uId$) + Len(uId$) + 1)
  Call dbg2f("wert=" + trm(rtmp!wert))
  altbvorl.AddItem rtmp!wert
  rtmp.MoveNext
Wend
End If
drvl = trm(GetDriveStrings())
While drvl <> ""
  altbvorl.AddItem cut_d1(drvl, Chr$(0))
  drvl = trm(cut_d2bis(drvl, Chr$(0)))
Wend
End Sub

Private Sub btncld_Click()
Dim rtmp As ADODB.Recordset
Dim c1 As ADODB.Connection
Dim c$, ttt$, rrr, cbp$, ask As Integer, X
Dim cuser$, cpass$, cserver$

MousePointer = 11
DoEvents
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "select FeldDaten from auftritthigru where auftrittstyp='webcal' and FeldName='cloud'"
On Error Resume Next
rtmp.Open c$, adoc, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  ttt$ = "Agencyprof can be configured to export data to a Horde groupware server that is able to sync adresses and calendars on handheld devices like smart- or iPhones." + vbCrLf
  ttt$ = ttt$ + "Step 1: Configure addresses (do not use contacts) that may connect:" + vbCrLf
  ttt$ = ttt$ + "Create an addresstype webcal with at least one property: cloud" + vbCrLf
  ttt$ = ttt$ + "Make an address you want to get 'its' data exported to be of 'webcal'" + vbCrLf
  ttt$ = ttt$ + "Enter the username of this address to connect to the hordeserver in 'cloud'" + vbCrLf
  ttt$ = ttt$ + vbCrLf
  ttt$ = ttt$ + "To hide this button permanently add a usersetting:" + vbCrLf + "cloud=no"
  MousePointer = 0: DoEvents
  MsgBox ttt$
  Exit Sub
End If

cserver$ = trm(getusersetting("cloud", ""))
cserver$ = strrepl(cserver$, ":", ";PORT=")
If cserver$ = "" Then
  ttt$ = "Agencyprof can be configured to export data to a Horde groupware server that is able to sync adresses and calendars on handheld devices like smart- or iPhones." + vbCrLf
  ttt$ = ttt$ + "Step 2: you need to configure the Horde database to connect to:" + vbCrLf
  ttt$ = ttt$ + "In the usersettings use the user 'system'." + vbCrLf
  ttt$ = ttt$ + "Set usersetting: cloud=<name or ip of server>" + vbCrLf
  ttt$ = ttt$ + "Set usersetting: clouduser=<username needed>" + vbCrLf
  ttt$ = ttt$ + "Set usersetting: cloudpass=<password for this user>" + vbCrLf + vbCrLf
  ttt$ = ttt$ + "To hide this button permanently add a usersetting:" + vbCrLf + "cloud=no"
  MousePointer = 0: DoEvents
  MsgBox ttt$
  Exit Sub
Else
  cuser$ = trm(getusersetting("clouduser", ""))
  cpass$ = trm(getusersetting("cloudpass", ""))
  cbp$ = "DATABASE=horde;SERVER=" + cserver$ + ";DRIVER=" + odbcdriver + ";UID=" + cuser$ + ";PWD=" + cpass$ + ";DSN="
  Set c1 = New ADODB.Connection
  c1.ConnectionString = cbp$
  On Error Resume Next
  c1.Open
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
  ttt$ = "Agencyprof can be configured to export data to a Horde groupware server that is able to sync adresses and calendars on handheld devices like smart- or iPhones." + vbCrLf
    ttt$ = ttt$ + "A connection to the database 'horde' at " + cserver$ + " failed." + vbCrLf
    ttt$ = ttt$ + "Most likely one of these is wrong - or the server cannot be reached. Check:" + vbCrLf
    ttt$ = ttt$ + "Set usersetting: cloud=" + cserver$ + vbCrLf
    ttt$ = ttt$ + "Set usersetting: clouduser=" + cuser$ + vbCrLf
    ttt$ = ttt$ + "Set usersetting: cloudpass=<hidden>" + vbCrLf + vbCrLf
    ttt$ = ttt$ + "To hide this permanently set usersetting:" + vbCrLf + "cloud=no"
    MsgBox ttt$
    MousePointer = 0: DoEvents
    Exit Sub
  Else
    c1.Close
    Load hordeinfo
    Call hordeinfo.SetFocus
    DoEvents
    hordeinfo.List3.AddItem "database horde@" + cserver$ + ": ok." + vbCrLf
    MousePointer = 0: DoEvents
  End If
End If

End Sub

Private Sub Check1_Click()
Dim s$

s$ = Combo1.text
If s$ <> "" Then
  break% = 0
  Call rlist1(s$)
End If

End Sub

Private Sub cldpusher_Timer()
Dim nq As String, nqd As String, nqn As String, i As Integer, o%, l$, rrr
Dim c$, berg As Boolean

If form1.hordexlock Then Exit Sub
cldpusher.Enabled = False
If Not cloud Then Exit Sub
Call dbg2f("Timer cldpush start")
DoEvents
If hordex.ListCount > 0 Then
  pb2.Visible = True
  If pb2.Max < hordex.ListCount Then pb2.Max = hordex.ListCount
End If
pb2.value = hordex.ListCount
DoEvents
i = 0
nq = newcloudqfile()
If nq = "" Then
  btncld.Caption = "queue access error"
  cldpusher.Interval = 57024
  cldpusher.Enabled = True
  Exit Sub
End If
nqd = DirName(nq)
cloudupds.Clear
nq = Dir(nqd + "\*.sql")
While nq <> "" And i < 10000
  i = i + 1
  cloudupds.AddItem nq
  nq = Dir()
Wend
i = 0
While cloudupds.ListCount > 0
  nq = cloudupds.List(0)
  If Not nexist(nqd + "\" + nq) Then
    o% = FreeFile()
    Open nqd + "\" + nq For Input As #o%
    c$ = ""
    While Not EOF(o%)
      Line Input #o%, l$
      If l$ <> "" Then
        c$ = c$ + " " + l$
        DoEvents
      End If
    Wend
    Close #o%
    DoEvents
    Call xhorde(c$)
  End If
  DoEvents
  On Error Resume Next
  Kill nqd + "\" + nq
  rrr = Err
  On Error GoTo 0
  DoEvents
  If rrr <> 0 And rrr <> 53 Then
    btncld.Caption = "queue access error:" + trm(rrr)
    cldpusher.Interval = 57024
    cldpusher.Enabled = True
    Call dbg2f("Timer cld_push exit")
    Exit Sub
  End If
  cloudupds.RemoveItem 0
  DoEvents
  i = i + 1
  If (i Mod 10) = 0 Then
    btncld.Caption = "Q: " + trm(cloudupds.ListCount)
    If hordex.ListCount > 0 Then btncld.Caption = btncld.Caption + vbCrLf + "address q: " + trm(hordex.ListCount)
  End If
  DoEvents
Wend
If nq = "" Then
  i = 0
  If hordex.ListCount > 0 Then
    If pb2.Max < hordex.ListCount Then
      pb2.Max = hordex.ListCount
    End If
    Do
      l$ = hordex.List(0)
      c$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
      berg = form1.cloudcreateadr(c$, "", l$)
      DoEvents
      Call form1.adr2cloud(c$)
      hordex.RemoveItem 0
      btncld.Caption = "Q: " + trm(cloudupds.ListCount)
      If hordex.ListCount > 0 Then btncld.Caption = btncld.Caption + vbCrLf + "address q: " + trm(hordex.ListCount)
      If pb2.Max < hordex.ListCount Then pb2.Max = hordex.ListCount
      pb2.value = hordex.ListCount
      DoEvents
      i = i + 1
'    Loop Until berg Or i > 5 Or hordex.ListCount = 0
      Loop Until hordex.ListCount = 0
    nq = "x"
  End If
End If
If hordex.ListCount = 0 Then pb2.Visible = False
If nq <> "" Then
  cldpusher.Interval = 1000
Else
  cldpusher.Interval = 15000
End If
cldpusher.Enabled = True
Call dbg2f("Timer cldpush exit")

End Sub

Private Sub Combo1_Click()
Dim d0 As Variant, sid$

DoEvents
List1.Clear
If Combo1.ListCount >= 0 Then
  Call combo1_Change
End If
sid$ = trm(Combo1.text)
If InStr(sid$, "(") > 1 Then sid$ = Left$(sid$, InStr(sid$, "(") - 1)
Load shwAdrDetail
Call shwAdrDetail.savecheck
Call shwAdrDetail.refreshadrdetail(sid$, "")
List1.AddItem "--Aktuelle Adressen--"
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s$
altsuch.Clear
If Not use_adrsuchtimer Then
  If KeyCode = 13 Then
    s$ = Combo1.text
    If s$ <> "" Then
      break% = 0
      Call rlist1(s$)
    End If
  Else
    altsuch.Clear
    usealtsuch = False
  End If
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

shwAdrDetail.srchit% = 0

End Sub

Public Sub Command1_Click()

Unload form1

End Sub

Private Sub Command10_Click()

break% = 1

Load prog
Call prog.SetFocus

End Sub

Private Sub Command11_Click()

'Command11.Caption = "Todo"
Load todolist
Call todolist.SetFocus
Command11.Picture = Picture2.Picture
Timer3.Enabled = False
End Sub

Private Sub Command12_Click()

Load msafe
End Sub

Private Sub Command13_Click()

On Error Resume Next
Unload shwAdrDetail
DoEvents
On Error GoTo 0
Load shwAdrDetail
On Error Resume Next
Call shwAdrDetail.SetFocus
On Error GoTo 0
End Sub

Private Sub Command14_Click()

Load verwalt_public
On Error Resume Next
Call verwalt_public.SetFocus
On Error GoTo 0

End Sub

Private Sub Command15_Click()

Load kc
Call kc.settag0(Date)
If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then kc.Text1.text = 1
On Error Resume Next
Call kc.SetFocus
Call k3.SetFocus
On Error GoTo 0

End Sub

Private Sub Command16_Click()
Dim o%, l$, rrr

Load smtp
smtp.Visible = True
Call smtp.SetFocus
smtp.txtServer.Enabled = False
smtp.txtMailFrom.Enabled = False
Call signaturinclude

End Sub


Private Sub Command17_Click()
Dim rrr

break% = 1
On Error Resume Next
Load taliste
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Call taliste.SetFocus
  Call taliste.Command15_Click
  Call taliste.presel("Künstler")
End If
End Sub

Private Sub Command18_Click()

Call handbuchcall("04-Hauptformular.htm")

End Sub
Public Sub handbuchcall(p$)
Dim u$, p1$
Dim brw$, X

p1$ = transe(trm(p$))
If p1$ = "" Then p1$ = "index.html"
Unload frmBrowser
DoEvents
u$ = form1.getusersetting("Handbuch", transe("http://www.agencyprof.de/tutorial"))
If u$ = "" Then
  u$ = "file:///" & form1.s0dir() & "/handbuch/" & p1$
Else
   u$ = u$ & "/" & p1$
End If
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & u$, 1)
Else
  frmBrowser.StartingAddress = u$
  Load frmBrowser
End If

End Sub

Private Sub Command19_Click()
Dim fn$, X

fn$ = mydir()
If trm(altbvorl.text) <> "" Then fn$ = trm(altbvorl.text)


X = Shell("explorer.exe " & fn$, vbNormalFocus)

End Sub

Private Sub Command20_Click()
Dim rrr

break% = 1
On Error Resume Next
Load taliste
rrr = Err
If rrr = 0 Then
  Call taliste.SetFocus
  Call taliste.Command15_Click
  Call taliste.presel("Kammermusik")
End If

End Sub

Private Sub Command21_Click()
Dim rrr

break% = 1
On Error Resume Next
Load taliste
rrr = Err
If rrr = 0 Then
  Call taliste.SetFocus
  Call taliste.Command15_Click
  Call taliste.presel("Kammermusik")
End If
End Sub

Private Sub Command22_Click()

Exit Sub
Load splan
On Error Resume Next
Call splan.SetFocus
On Error GoTo 0

End Sub

Private Sub Command23_Click()

break% = 1
Load tplan
tplan.setcaption (transe("Kammermusik - Projekte"))
Call tplan.SetFocus

End Sub

Private Sub Command24_Click()

break% = 1
Load tplan
tplan.setcaption (transe("Crossover- Projekte"))
Call tplan.SetFocus

End Sub

Private Sub Command25_Click()

Load sels
On Error Resume Next
Call sels.SetFocus
On Error GoTo 0

End Sub

Private Sub Command27_Click()

Call Form_DblClick
End Sub

Private Sub Command28_Click()
Dim c$, r As ADODB.Recordset, s As ADODB.Recordset, rrr, na$, p%, rl$, ady$, trgd$
Dim dd$, mm$, rfd As String, s0y As Integer, backd%, vid$, svid As String, sid As String
Dim age%, dtx, diffdays, dirdt$, delfl As Boolean, ddays, trgdx$, o%, sp1 As String, sp2 As String
Dim sp3 As String, sp4 As String, typlist As String, typgo As Boolean, ttyp As String, twert As String
Dim xact As Boolean, tanz As Integer, t1anz As Integer

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "Command28_Click"
List3.Clear
List3.ToolTipText = "Geburtstagsliste"
sortlist.Clear
typlist = ""
backd% = -Val(getusersetting("VergangeneGeburtstage", "10"))
'c$ = "select * from auftritthigru where feldname='Geburtstag' and felddaten<>''"
c$ = "select count(*) as anz from auftritthigru where feldname='Geburtstag' and length(felddaten)>0"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
pbg1.Top = altbvorl.Top
pbg1.Left = altbvorl.Left
pbg1.Visible = True
MousePointer = 11: DoEvents: tanz = 0
pbg1.value = 0
If Not r.EOF Then
  tanz = r!anz
End If
pbg1.Max = imax(tanz, 1)
r.Close
t1anz = 0
c$ = "select * from auftritthigru where feldname='Geburtstag' and length(felddaten)>0"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub
While Not r.EOF
  pbg1.value = pbg1.value + 1
  DoEvents
  vid$ = r!auftrittsid
  na$ = form1.getnamebyid(vid$)
  typgo = False
  If typlist = "" Then typgo = True
  c$ = typlist
  If na$ <> "" Then
    While c$ <> "" And typgo = False
      ttyp = cut_d1(c$, "|")
      c$ = cut_d2bis(c$, "|")
      If InStr(ttyp, "=") > 0 Then
        If InStr(ttyp, "==") > 0 Then
          xact = True
          ttyp = strrepl(ttyp, "==", "=")
        Else
          xact = False
        End If
        twert = cut_d2bis(ttyp, "=")
        ttyp = strrepl(ttyp, "=", "|")
      End If
      If Not xact Then
        If isoftype(vid, ttyp) <> "-1" Then typgo = True
      Else
        If LCase(isoftype(vid, ttyp)) = twert Then
          typgo = True
        End If
      End If
    Wend
  Else
    While c$ <> "" And typgo = False
      ttyp = cut_d1(c$, "|")
      c$ = cut_d2bis(c$, "|")
      If InStr(ttyp, "=") > 0 Then
        If InStr(ttyp, "==") > 0 Then
          xact = True
          ttyp = strrepl(ttyp, "==", "=")
        Else
          xact = False
        End If
        twert = cut_d2bis(ttyp, "=")
        ttyp = strrepl(ttyp, "=", "|")
      End If
      If Not xact Then
        If kisoftype(vid, ttyp) <> "-1" Then typgo = True
      Else
        If LCase(kisoftype(vid, ttyp)) = twert Then typgo = True
      End If
    Wend
    If typgo Then
    'c$ = "select name,vid,id from kontakt where vid + id='" & vid$ & "'"
      c$ = "select name,vid,id from kontakt where (instr('" & vid$ & "',vid)=1) and (instr('" & vid$ & "',id)>0)"
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not s.EOF Then
        na$ = trm(s!name)
        svid = trm(s!id)
        sid = trm(s!id)
        s.MoveNext
      End If
      s.Close
    End If
  End If
  If na$ <> "" And typgo Then
    rfd = trm(r!felddaten)
    dd$ = cut_d1(rfd, "."): p% = InStr(rfd, ".") + 1
    rl$ = Mid$(rfd, p%)
    mm$ = cut_d1(rl$, "."): p% = InStr(rl$, ".") + 1
    rl$ = Mid$(rl$, p%)
    ady$ = ""
    If p% > 1 Then ady$ = word1(rl$)
    trgd$ = dd$ & "." & mm$ & "." & trm(str$(apyear(now)))
    On Error Resume Next
    ddays = Int(CDate(trgd$) - now) + 1
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      dd$ = "28"
      trgd$ = dd$ & "." & mm$ & "." & trm(str$(apyear(now)))
    Else
      trgdx$ = dd$ & "." & mm$ & "." & trm(str$(apyear(now) + 1))
      On Error Resume Next
      ddays = Int(CDate(trgdx$) - now) + 1
      rrr = Err
      On Error GoTo 0
    End If
    If rrr <> 0 Then
      dd$ = "28"
      trgd$ = dd$ & "." & mm$ & "." & trm(str$(apyear(now)))
    End If
    s0y = 0
    On Error Resume Next
    s0y = Val("0" & trm(ady$))
    rrr = Err
    On Error GoTo 0
    If s0y < 100 Then s0y = s0y + 1900
    age% = apyear(now) - s0y
    On Error Resume Next
    dtx = CDate(trgd$)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      If dtx - now < backd% Then
        trgd$ = dd$ & "." & mm$ & "." & trm(str$(apyear(now) + 1))
        age% = age% + 1
      End If
      If ady <> "" Then ady$ = " (" & trm(str$(age%)) & ")"
      On Error Resume Next
      diffdays = Int(CDate(trm(trgd$)) - now) + 1
      rrr = Err
      On Error GoTo 0
      If rrr = 0 Then
        dirdt$ = transe("in:")
        If diffdays < 0 Then
          diffdays = Abs(diffdays)
          dirdt$ = transe("vor:")
        End If
      Else
        dirdt$ = transe("Fehler:")
        diffdays = 0
      End If
      'List3.AddItem datum2sql(trgd$) & ": " & na$ & ", " & dirdt$ & " " & diffdays & " " + transe("Tagen") + " " & ady$ & Space$(160) & "(ID:" & na$
      sortlist.AddItem datum2sql(trgd$) & ": " & na$ & ", " & dirdt$ & " " & diffdays & " " + transe("Tagen") & ady$ & Space$(160) & ":" & rfd & " (ID:" & na$
    End If
  End If
  r.MoveNext
Wend
r.Close
List3.Clear
pbg1.Visible = False
While sortlist.ListCount > 0
  delfl = True
  While sortlist.ListCount > 1 And delfl = True
    If sortlist.List(0) = sortlist.List(1) Then
      sortlist.RemoveItem 1
    Else
      delfl = False
    End If
  Wend
  List3.AddItem sortlist.List(0)
  sortlist.RemoveItem 0
Wend
c$ = mydir() + "\geburtstage.csv"
o% = FreeFile
On Error Resume Next
Open c$ For Output As #o%
rrr = Err
On Error GoTo 0
c$ = getusersetting("exceldelimiter", ";")
If rrr = 0 Then
  For p% = 0 To List3.ListCount - 1
    sp1 = cut_d1(List3.List(p%), ":")
    rl$ = cut_d2bis(List3.List(p%), ":")
    sp2 = cut_d1(rl, ":")
    rl$ = cut_d2bis(rl$, ":")
    sp3 = trm(cut_d1(rl, ":"))
    rl$ = cut_d2bis(rl$, ":")
    sp4 = trm(strrepl(cut_d1(rl, ":"), "(ID", ""))
  Print #o%, """" + sp1$ + """" + c$ + """" + sp2$ + """" + c$ + """" + sp3$ + """" + c$ + """'" + sp4$ + """"
  Next p%
  Close #o%
'  rrr = Shell("explorer.exe " & mydir(), vbNormalFocus)
End If
MousePointer = 0

End Sub

Private Sub Command29_Click()
'd2infile = "Form1": d2insub = "Command29_Click"
Load dayvw
On Error Resume Next
Call dayvw.SetFocus
On Error GoTo 0
dayvw.Text1.text = word1(Label1.Caption)

End Sub

Private Sub Command3_Click()
'd2infile = "Form1": d2insub = "Command3_Click"
break% = 1

Load werkvz
werkvz.Visible = True
Call werkvz.SetFocus

End Sub

Private Sub Command30_Click()
Dim nid$, tpid$, d0

Dim d2infile As String, d2insub As String
d2infile = "form1": d2insub = "Command30_Click"
tpid$ = "-1"
MousePointer = 11: DoEvents
d0 = CDate(Date)
nid$ = form1.newid("auftritt", "id", 20)
form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
               nid$ & "','" + tpid$ + _
               "','Neuer Auftritt','Neuer Auftritt','" + _
               datum2sql(CDate(d0)) & "')")
Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(nid$, 0)
MousePointer = 0

End Sub

Private Sub Command31_Click()
Call Form_DblClick
Unload shwAdrDetail
mPopupSys.Enabled = True
mPopupSys.Visible = True
mPopExit.Enabled = True
mPopRestore.Enabled = True
       With nid
        .cbSize = Len(nid)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hicon = Me.Icon
        .szTip = Me.Caption & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid
ismin = True
Hide
End Sub

Private Sub Command32_Click()
Exit Sub
If form1.isfieldmissing("opt_listen", "id") Then Exit Sub
Load listen
On Error Resume Next
Call listen.SetFocus
On Error GoTo 0

End Sub

Private Sub Command4_Click()
Dim rrr

'd2infile = "Form1": d2insub = "Command4_Click"
break% = 1
On Error Resume Next
Load taliste
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Call taliste.SetFocus
  Call taliste.Command15_Click
  Call taliste.presel("Orchester")
End If
End Sub

Private Sub Command5_Click()
'd2infile = "Form1": d2insub = "Command5_Click"
break% = 1
Load tplan
tplan.setcaption (transe("Orchester - Projekte"))
On Error Resume Next
Call tplan.SetFocus
On Error GoTo 0
End Sub


Private Sub Command6_Click()
'd2infile = "Form1": d2insub = "Command6_Click"
break% = 1

On Error GoTo exmec6
Load auftrittselect
On Error Resume Next
auftrittselect.SetFocus
On Error GoTo 0

exmec6:
On Error GoTo 0

End Sub


Private Sub Command7_Click()
'd2infile = "Form1": d2insub = "Command7_Click"
break% = 1
Load tplan
tplan.setcaption (transe("Künstler - Projekte"))
On Error Resume Next
Call tplan.SetFocus
On Error GoTo 0

End Sub


Private Sub Command8_Click()

ihavemail = False
frmMain.Visible = True
frmMain.SetFocus
frmMain.txtUserName = upopid$
frmMain.txtServer = upop$
frmMain.txtPassword = upoppsswd$
frmMain.txtPort = trm(upopport%)
DoEvents
Command8.Picture = mlstat(0).Picture
Call frmMain.chkm

End Sub

Private Sub Command9_Click()
Dim brw$, X

'd2infile = "shwAdrDetail": d2insub = "Label11_Click"
Unload frmBrowser
DoEvents
brw$ = form1.UseBrowser()
If brw$ <> "" Then
  X = Shell(brw$ & " " & getmyhomepg(), 1)
Else
  frmBrowser.StartingAddress = getmyhomepg()
  Load frmBrowser
End If

End Sub

Private Sub dat_clsall_Click()
Call Form_DblClick
End Sub

Private Sub dat_end_Click()
Call Command1_Click
End Sub

Private Sub edt_adrsuch_Click()
Call Label3_DblClick
End Sub

Private Sub edt_erecv_Click()
Call Command8_Click
End Sub

Private Sub edt_esend_Click()
Call Command16_Click
End Sub

Private Sub edt_msafe_Click()
Call Command12_Click
End Sub

Private Sub edt_myfiles_Click()
altbvorl.text = ""
Call Command19_Click
End Sub

Private Sub edt_prio_Click()
Call Label7_DblClick
End Sub

Private Sub edt_set_Click()
Call uuid_DblClick
End Sub

Private Sub edt_tools_Click()
Call Command14_Click
End Sub

Private Sub errmess_Click()
Dim o%, stmp As ADODB.Recordset, logid$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "errmess_Click"
If InStr(errmess.Caption, transe("&Fehler")) > 0 Then
  Load fehler
End If
End Sub

Private Sub fallbackq_DblClick(Index As Integer)
'd2infile = "Form1": d2insub = "fallbackq_DblClick"
If fallbackdir$ <> "" Or fallbackserver$ <> "" Then
  Load Datenreplikator
  Datenreplikator.Show
End If

End Sub

Public Sub Form_DblClick()
'd2infile = "Form1": d2insub = "Form_DblClick"
Call unloadall

End Sub
Public Sub setAuftrittsdruckFuerAdresse(a$)
'd2infile = "Form1": d2insub = "setAuftrittsdruckFuerAdresse"
AuftrittsdruckFuerAdresse = a$
End Sub

Private Sub Form_Load()
Dim prpLoop As Property, o%, s1d$, idc As Long, rrr, geodbpara$, clddbpara$, upw$, ldtg As Double
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, c$, i%, l$, fallbackserverdatenbank$, ww$, cldsrv$
Dim fallbackserverusername$, fallbackserverpasswort$, geodbserver$, mew, meh, sid$, ttt$
Dim outl$, tr As String, astf As Long, lupd, ddiff, X, say$, resetfb As Boolean, connectfb As Boolean

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "Form_Load"
axsResizer1.SaveControlPositions

Hide
dbg_prvtlnkcount = ""
libsoll = hexstring2dec("0x10002")
adostats_tsum = 0
adostats_samples = 0
dbswitch = False
tlopen = False
brwhidden = False
resetfb = False
connectfb = False
shwled = False
hordexlock = False
dochistlock = False
fastsave_copy = False
starting = True
mydemoid$ = ""
missingfields = ""
pb2.Min = 0: pb2.Max = 1
dbg2file% = 0
autocheckmail = True
uselimitinsql = True
exceldelim$ = ","
skip1del = False
ttabptr% = 0
listenhauptperson = ""
dochistisopen = False
auftrittisopen = False
todoisopen = False
currentconfmode$ = ""
s00d$ = CurDir
s0d$ = CurDir
If InStr(s0d$, "\ldrv\") > 0 Then
  i% = InStr(s0d$, "\ldrv\") + 6
  s0d$ = Mid$(s0d$, i%)
  s0d$ = "l:\" + s0d$
End If
localdir = s0d$
If nexist("startlog.yes") Then starting = False
weckerpresent = False
If Not nexist(s0d$ + "\wecker.exe") Then weckerpresent = True
usrprofile$ = ""
lnkcolor = &H8000000D
'lnkcolor = RGB(255, 0, 0)
ihavemail = False
kalopen = False
dayvopen = False
priosopen = False
menoquit = False
globcount = 1000000
pspath$ = ""
poplock = False
pin.Visible = False
pinlbl.Visible = False
errsh% = 1
auftrittsdruck_currvorlage$ = "": auftrittsdruck_currfeld$ = ""
adrmerkid$ = ""
dirtcol = RGB(211, 238, 249)
AuftrittsdruckFuerAdresse = ""
mnams$(1) = "Januar"
mnams$(2) = "Februar"
mnams$(3) = "März"
mnams$(4) = "April"
mnams$(5) = "Mai"
mnams$(6) = "Juni"
mnams$(7) = "Juli"
mnams$(8) = "August"
mnams$(9) = "September"
mnams$(10) = "Oktober"
mnams$(11) = "November"
mnams$(12) = "Dezember"
dayname(1) = "So"
dayname(2) = "Mo"
dayname(3) = "Di"
dayname(4) = "Mi"
dayname(5) = "Do"
dayname(6) = "Fr"
dayname(7) = "Sa"
longdayname(1) = "Sonntag"
longdayname(2) = "Montag"
longdayname(3) = "Dienstag"
longdayname(4) = "Mittwoch"
longdayname(5) = "Donnerstag"
longdayname(6) = "Freitag"
longdayname(7) = "Samstag"
noalarms = False
err_dupok% = 0
Label10.Caption = "": Label10.Visible = False
i% = 0: statusname$(i%) = "interessiert": statusfarbe(i%) = RGB(200, 200, 200)
i% = 1: statusname$(i%) = "geplant": statusfarbe(i%) = RGB(140, 140, 140)
i% = 2: statusname$(i%) = "bestätigt": statusfarbe(i%) = RGB(0, 150, 230)
i% = 3: statusname$(i%) = "hat stattgefunden": statusfarbe(i%) = RGB(200, 250, 0)
i% = 4: statusname$(i%) = "abgesagt": statusfarbe(i%) = RGB(255, 255, 255)
For i% = 5 To 9
  statusname$(i%) = "kein Status": statusfarbe(i%) = RGB(255, 255, 255)
Next i%
mwst = 0.19
memono% = 0
t3tick% = 0
d0t0m = 6
t2tick% = 0
umsgwait% = 0
ehsc% = 0
SelectedDate$ = ""
selectedcolor = -1
d0t0y = 2060
Randomize
msec = 1# / (24# * 3600# * 1000#)

Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
Combo1.text = ""
Timer3.Enabled = False
List1.Clear
List2.Clear

uId$ = ""
errmess.Visible = False
noliccnt% = 0
myip$ = GetIPAddress()
Load login
While uId$ = "": DoEvents: DoEvents: Wend
Call form1.startlog(uId$, "unloading login")
Unload login
If uId$ = "" Or uId$ = "_LOGOUT_" Then
  Call Command1_Click
  MsgBox iml("Benutzer") + " " & uId$ & " " + transe("kann nicht angemeldet werden")
  Exit Sub
End If
internalkey = "kzJfuz5vFRiuZ9oui974kJHbkGf"
Call startup.SetFocus: DoEvents
If dbpara$ <> "msaccessmdb" Then
  startup.List1.AddItem iml("Datenbank (DAO) wird geöffnet"): DoEvents
  Set sqla = wrkJet.OpenDatabase(dbname$, dbDriverCompleteRequired, False, dbpara$)
  Set pub_sqla = wrkJet.OpenDatabase(dbname$, dbDriverCompleteRequired, False, dbpara$)
Else
  startup.List1.AddItem iml("Datenbank (MDB) wird geöffnet"): DoEvents
  Set sqla = wrkJet.OpenDatabase(dbname$, False, False)
  Set pub_sqla = wrkJet.OpenDatabase(dbname$, False, False)
  uselimitinsql = False
End If
Call form1.startlog(uId$, "DAO-Open beendet")

startup.List1.AddItem iml("Datenbank (ADO) wird geöffnet"): DoEvents
Set adoc = New ADODB.Connection
adoc.ConnectionString = adopara$
adoc.Open
Call form1.startlog(uId$, "ADO-Open beendet")
aKey = "jhfzer6498"
useusrcache = getusersetting("cachesettings", "ja")
upw$ = getusersetting("appasswort", "")
pwentered = ""
If upw$ <> "" Then
  Load pwenter
  While pwentered = ""
    DoEvents
  Wend
  If "passwo2345rteingabeabge4356brochen11223476" = pwentered Then
    Call Command1_Click
    End
  End If
  If upw$ <> pwentered Then
    MsgBox "Falsches Passwort."
    Call Command1_Click
    End
  End If
End If
c$ = "select ID from benutzerdaten where ID='" + uId + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
On Error Resume Next
rtmp.Open c$, adoc, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  If Not rtmp.EOF Then
    If uId <> trm(rtmp!id) Then
      MsgBox "User '" + uId + "' unknown, changed to '" + trm(rtmp!id) + "'."
      uId = trm(rtmp!id)
    End If
  End If
End If
Call form1.startlog(uId$, "Benutzereinstellungen laden")
startup.List1.AddItem iml("Benutzereinstellungen werden geladen"): DoEvents
s0d$ = getusersetting("agencyprof", s00d$)
If nexist(s0d$ + "\Agencyprof1.exe") Then
  c$ = Left(s00d$, 1) + Mid$(s0d$, 2)
  If LCase(s0d$) <> LCase(c$) Then
    MsgBox (s0d$ + " " + transe("nicht gefunden, benutzt wird") + " " + c$)
    s0d$ = c$
  End If
End If
localdir = getusersetting("localdir", s0d$)
On Error Resume Next
MkDir localdir
rrr = Err
On Error GoTo 0
If rrr <> 0 And rrr <> 75 Then localdir = s0d$
Call rereadsomesysvars
For i% = 0 To 9
  ttt$ = form1.getusersetting("astatusname" + trm(i%), "")
  If ttt$ <> "" Then
    statusname$(i%) = ttt$
    astf = RGB(255, 255, 255)
    ttt$ = form1.getusersetting("astatusfarbe" + trm(i%), "")
    If ttt$ <> "" Then astf = CLng(ttt$)
    statusfarbe(i%) = astf
  End If
Next i%
Shape2.BackColor = getusersetting("shapecolor", "12632256")
Shape3.BackColor = getusersetting("shapecolor", "12632256")
On Error Resume Next
convertcolor = Val(getusersetting("convertcolor", "16434176"))
rrr = Err
On Error GoTo 0
If rrr <> 0 Then convertcolor = 16434176
If libist > 1 Then
  If getusersetting("debuglib2file", "0") = "0" Then Call bas_APLibWriteLog(0)
End If
Call form1.startlog(uId$, "cloud detection")
startup.List1.AddItem "Cloud Detection": DoEvents
cloud = False: c$ = " ok": cloudmanager = "": cloudstaff = ""
cloudserver$ = trm(getusersetting("cloud", ""))
If cloudserver$ = "no" Then
  cloud = False
  btncld.Enabled = False
  btncld.Visible = False

Else

supershares_krono = ""
If cloudserver$ = "" Then
  ttt$ = " not configured"
Else
  clouduser$ = trm(getusersetting("clouduser", ""))
  cloudpass$ = trm(getusersetting("cloudpass", ""))
  cldsrv$ = strrepl(cloudserver$, ":", ";PORT=")
  clddbpara$ = "DATABASE=horde;SERVER=" + cldsrv$ + ";DRIVER=" + odbcdriver + ";UID=" + clouduser$ + ";PWD=" + cloudpass$ + ";DSN="
  Set clddb = New ADODB.Connection
  clddb.ConnectionString = clddbpara$
  On Error Resume Next
  clddb.Open
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    ttt$ = " no database"
  Else
    cloud = True
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    c$ = "select FeldDaten from auftritthigru where auftrittstyp='webcal' and FeldName='cloud'"
    On Error Resume Next
    rtmp.Open c$, adoc, adOpenDynamic, adLockReadOnly
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      cloud = False
      ttt$ = " webcal missing"
    Else
      rtmp.Close
      c$ = "select auftrittsid,FeldDaten from auftritthigru where (FeldDaten='Manager' or FeldDaten='Mitarbeiter') and FeldName='Benutzergruppe' and auftrittstyp='webcal'"
      On Error Resume Next
      rtmp.Open c$, adoc, adOpenDynamic, adLockReadOnly
      rrr = Err
      On Error GoTo 0
      While Not rtmp.EOF
        c$ = "select FeldDaten as wert from auftritthigru where auftrittsid='" + trm(rtmp!auftrittsid) + "' and FeldName='cloud' and auftrittstyp='webcal'"
        If LCase(trm(rtmp!felddaten)) = "manager" Then
          c$ = get1erg(c$)
          cloudmanager = cloudmanager + "|" + c$ + "|"
          c$ = "select share_name as wert from kronolith_sharesng where share_owner='" + c$ + "'"
          supershares_krono = supershares_krono + "|" + get1hordeerg(c$) + "|"
        End If
        If LCase(trm(rtmp!felddaten)) = "mitarbeiter" Then
          c$ = get1erg(c$)
          cloudstaff = cloudstaff + "|" + c$ + "|"
          c$ = "select share_name as wert from kronolith_sharesng where share_owner='" + c$ + "'"
          supershares_krono = supershares_krono + "|" + get1hordeerg(c$)
        End If
        rtmp.MoveNext
      Wend
      cloudmanager = strrepl(cloudmanager, "||", "|")
      supershares_krono = strrepl(supershares_krono, "||", "|")
    End If
  End If
End If
End If
cldpusher.Enabled = False: isp3home = ""
If Not cloud Then
  btncld.Caption = "no cloud:" + vbCrLf + ttt$
  tmrcld.Enabled = False
Else
  btncld.Caption = "cloud starting"
  Call tmrcld_Timer
  cldpusher.Interval = 10000
  cldpusher.Enabled = True
  isp3home = getusersetting("isp3home", "https://" + cloudserver$ + ":8080")
  c$ = "select share_id,share_name,attribute_name from turba_sharesng"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    btncld.Caption = "no cloud:" + vbCrLf + ttt$
    tmrcld.Enabled = False
    cloud = False
  Else
    r.Close
  End If
  DoEvents
End If
Set geodb = New ADODB.Connection
geodbserver$ = trm(getusersetting("geodbserver", "localhost"))
geodbok = False

If geodbserver$ <> "nein" And geodbserver$ <> "" Then
  say$ = "GeoDB (" + geodbserver$ + ") " + iml("wird geöffnet")
  Call form1.startlog(uId$, say$)
  startup.List1.AddItem say$: DoEvents
  i% = InStr(adopara$, ";DRIVER=")
  c$ = ""
  If i% > 0 Then c$ = Mid$(adopara$, i%)
  geodbpara$ = "DATABASE=opengeodb;SERVER=" + geodbserver$ + c$
  Call form1.startlog(uId$, geodbpara$)
  geodb.ConnectionString = geodbpara$
  On Error Resume Next
  geodb.Open
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    Call form1.startlog(uId$, "geodb.open ok")
    geodbok = True
    c$ = "select * from geodb_type_names limit 0,1;"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    On Error Resume Next
    rtmp.Open c$, geodb, adOpenDynamic, adLockReadOnly
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then geodbok = False
  Else
    Call form1.startlog(uId$, "Fehler " + trm(rrr) + " bei geodb.Open")
  End If
Else
  startup.List1.AddItem iml("GeoDB nicht konfiguriert"): DoEvents
  Call form1.startlog(uId$, "GeoDB nicht konfiguriert")
End If

On Error Resume Next
MkDir s0d$ & "\tmp"
On Error GoTo 0
Call form1.startlog(uId$, "Ermitteln der Fallbackserver")
fallbackserver$ = getusersetting("fallbackserver", "")
If fallbackserver$ = "nein" Or fallbackserver = "none" Then fallbackserver = ""
Call fieldcheck("opt_repliken", "id")
If Not form1.isfieldmissing("opt_repliken", "id") Then
  If fallbackserver$ <> "" And (dbserver$ = "localhost" Or dbserver$ = "127.0.0.1") Then
    c$ = "select count(lfdnr) as rc from opt_repliken"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If rrr = 0 Then
      If r!rc > 0 Then
'MsgBox "i should not have opt_repliken filled"
Debug.Print "i should not have opt_repliken filled reset fallbackserver?"
      End If
    End If
  End If
End If
If iamdemo() Then
  If fallbackserver = "" Then
    If fallbackserverpasswort$ = "" Then
      resetfb = True
      fallbackserver = "apdemo" + Right$(form1.mydemoid, 1)
      Call form1.setusersetting("fallbackserver", fallbackserver)
      Call form1.setusersetting("fallbackserverusername", "ap")
      Call form1.setusersetting("fallbackserverdatenbank", "example")
      c$ = "decrypt:" + encrypt(dbpasswd, form1.getinternalkey())
      Call form1.setusersetting("fallbackserverpasswort", c$)
    End If
  Else
    resetfb = False
    If nexist(replicationfilename(form1.computername)) Then resetfb = True
    connectfb = True
  End If
End If
fallbackserverusername$ = getusersetting("fallbackserverusername", "")
fallbackserverpasswort$ = getusersetting("fallbackserverpasswort", "")
fallbackserverdatenbank$ = getusersetting("fallbackserverdatenbank", "")
If fallbackserver$ <> "" Then
  dbfpara$ = "DATABASE=" & fallbackserverdatenbank$
  dbfpara$ = dbfpara$ + ";SERVER=" & strrepl(fallbackserver$, ":", ";PORT=")
  dbfpara$ = dbfpara$ + ";DRIVER=" + form1.odbcdriver
  dbfpara$ = dbfpara$ + ";UID=" & fallbackserverusername$
  dbfpara$ = dbfpara$ + ";PWD=" & fallbackserverpasswort$
  If InStr(dbpara$, dbfpara$) > 0 Then
    dbfpara$ = ""
    fallbackserver$ = ""
    MsgBox ("Die Datenbank scheint mit der Masterdatenbank identisch zu sein." + vbCrLf + "Eine Datenreplikation findet in dieser Sitzung nicht statt.")
  End If
  On Error Resume Next
  If fallbackserver$ <> "" Then fallbackq(0).Caption = fallbackserver
  MkDir s00d$
  MkDir s00d$ & "\fallbackserver"
  MkDir s00d$ & "\fallbackserver\" & strrepl(fallbackserver$, ":", "_")
  On Error GoTo 0
  fallbackserverpath$ = s00d$ & "\fallbackserver\" & strrepl(fallbackserver$, ":", "_")
  fallbackserverpath$ = getusersetting("fallbackserverpath", fallbackserverpath$)
Else
  dbfpara$ = ""
End If
fallbackdir$ = getusersetting("fallbackdir", "")
If fallbackdir$ <> "" Then
  If InStr(LCase(fallbackdir$), ":\") = 0 Then fallbackdir$ = s00d$ & "\" & fallbackdir$
  On Error Resume Next
  MkDir fallbackdir$
  On Error GoTo 0
  fallbackdir$ = fallbackdir$ & "\" & dbname$
  On Error Resume Next
  MkDir fallbackdir$
  On Error GoTo 0
End If
Call form1.startlog(uId$, "Formulare ...")
startup.List1.AddItem iml("Formulare vorbereiten und öffnen"): DoEvents
On Error Resume Next
MkDir mydatadir()
MkDir mydatadir() + "\positions"
MkDir mylocaldatadir()
MkDir mylocaldatadir() + "\positions"
On Error GoTo 0
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
mew = form1.mylastwidth(Me.name, 0)
meh = form1.mylastheight(Me.name, 0)
If meh > 0 And mew > 0 Then
  Me.Width = mew
  Me.Height = meh
End If
Call form1.formpos(Me)
Me.Caption = transe("Haupt-Formular - AgencyProf - ")
Me.Caption = Me.Caption + " " + transe("Datenbank") + ": " & dbname$ & " auf " + dbserver$ & ", " & s0d$
ufsze% = Val(getusersetting("fontsize"))
If ufsze% < 8 Or ufsze% > 12 Then ufsze% = 8
List1.Font.Size = ufsze%
List2.Font.Size = ufsze%
List3.Font.Size = ufsze%
Combo1.Font.Size = ufsze%

say$ = iml("Feldertest, weitere Benutzerdaten")
Call form1.startlog(uId$, say$)
startup.List1.AddItem say$: DoEvents
use_adrsuchtimer = True
If getusersetting("adresssuchtimer", "an") = "aus" Then use_adrsuchtimer = False
anredeuser$ = getusersetting("Anreden", uId$)
alertdbuid = uId$ + "|" & dbserver$ + "|" + dbname$
Call fieldcheck("opt_othertplans", "id")
Call fieldcheck("opt_repertoire", "id")
Call fieldcheck("auftrittsfelder", "opthordeshare")
Call fieldcheck("auftrittsfelder", "opthordesharewhat")
Call fieldcheck("adresse", "optinternal")
Call fieldcheck("opt_allenummern", "vid")
Call fieldcheck("adresse", "opttel")
Call fieldcheck("mailsafe", "optcc")
Call fieldcheck("mailsafe", "optan")
Call fieldcheck("kontakt", "opttel")
Call fieldcheck("auftritthigru", "opt_kid")
Call fieldcheck("auftritt", "optkalcolor")
Call fieldcheck("opt_adresspool", "id")
Call fieldcheck("opt_checks", "id")
Call fieldcheck("opt_topics", "id")
Call fieldcheck("opt_checklists", "id")
Call fieldcheck("opt_numbers", "id")
Call fieldcheck("opt_prios", "id")
Call fieldcheck("opt_listen", "id")
Call fieldcheck("opt_talisted1", "id")
Call fieldcheck("opt_vnr", "id")
Call fieldcheck("opt_stimmton", "id")
Call fieldcheck("opt_cocomposers", "id")
Call fieldcheck("opt_arranged", "id")
Call fieldcheck("opt_published", "id")
Call fieldcheck("opt_textdichter", "id")
Call fieldcheck("kontakt", "opt_kpos")
If Not isfieldmissing("opt_checklists", "id") Then
  c$ = "select * from opt_checklists where id=''"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  On Error Resume Next
  rtmp.Open c$, adoc, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    If rtmp.EOF Then
      c$ = "insert into opt_checklists (id) values('')"
      Call sqlqry(c$)
    End If
  End If
End If
On Error GoTo errex
uuid.Caption = uId$
On Error GoTo 0
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM benutzerdaten where id ='" + uId$ + "'"
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then
  Call form1.startlog(uId$, "neuer Benutzer")
  ww$ = "write.exe"
  Load dbupgrade
  dbupgrade.List1.Clear
  dbupgrade.List1.AddItem iml("Word wird gesucht...")
  MousePointer = 11
  DoEvents
  ueditor$ = getsystemsetting("defaulteditor")
  If exist(ueditor$) = 0 Then
    ueditor = ""
    'ueditor$ = wordsuchen()
    ueditor = GetWordPath()
  End If
  Unload dbupgrade
  MousePointer = 0
  DoEvents
  If ueditor$ <> "" Then ww$ = ueditor$
  sqlqry ("INSERT INTO benutzerdaten (ID,name,fax,tel,faxvorlage,briefvorlage,editor) " + _
           "VALUES('" & uId$ & "','Neuer Benutzer','Ihre Faxnummer','Ihre Telefonnummer','fax.rtf','" + "brief.rtf','" + ww$ + "')")
  On Error Resume Next
  Load einstellungen
  Call einstellungen.SetFocus
  On Error GoTo 0
  'Exit Sub
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM benutzerdaten where id ='" + uId$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rtmp.EOF Then Exit Sub
End If
exceldelim$ = getusersetting("exceldelimiter", ",")
dbg2file% = 0
Call dbupgrd

uname$ = "unbekannt"
If Not IsNull(rtmp!name) Then uname$ = rtmp!name
uuid.Caption = uId$
Combo1.text = ""
rtmp.Close
Call rlist3

Call form1.startlog(uId$, "Termintypen vorbereiten")
startup.List1.AddItem iml("Termintypen vorbereiten"): DoEvents
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM auftrittstypen"
Call form1.startlog(uId$, c$)
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
colorcachepointer% = 0
atabkzcachepointer% = 0
While Not rtmp.EOF
  Call form1.startlog(uId$, "atyp:" + trm(rtmp!id))
  colorcacheid$(colorcachepointer%) = trm(rtmp!id)
  On Error Resume Next
  atabkzcacheid$(atabkzcachepointer%) = trm(rtmp!abkz)
  atabkz$(atabkzcachepointer%) = trm(rtmp!id)
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then atabkzcachepointer% = atabkzcachepointer% + 1
  For i% = 0 To 2
    colorcache%(colorcachepointer%, i%) = Val("0" & trm(rtmp.Fields(2 + i%).value))
  Next i%
  colorcachepointer% = colorcachepointer% + 1
  rtmp.MoveNext
Wend
On Error Resume Next
rtmp.Close
On Error GoTo 0
Call form1.startlog(uId$, "Währungen lesen")
startup.List1.AddItem iml("Währungen lesen"): DoEvents
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT id,fixed,name FROM waehrung order by position"
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Call form1.startlog(uId$, c$)
waehrungen.Clear
While Not rtmp.EOF
  c$ = trm(rtmp!id) & ":" & trm(rtmp!fixed) & ":" & trm(rtmp!name)
  Call form1.startlog(uId$, c$)
  waehrungen.AddItem c$
  rtmp.MoveNext
Wend
On Error Resume Next
rtmp.Close
On Error GoTo 0
startup.List1.AddItem iml("Mailempfang konfigurieren"): DoEvents
'upop$ = getusersetting("POPServer")
'upopid$ = getusersetting("POP_Username")
'upoppsswd$ = getusersetting("POP_Passwort")
'upopport% = Val(getusersetting("POP_Port"))
dir_mailoutbox = ""
umsrvout$ = "dir:Outbox"
Call form1.startlog(uId$, "mailserver=" + umsrvout$)
If LCase(upop$) = "dir:inbox" Then
  dir_mailoutbox = form1.s0dir() + "\" + form1.docs() + "\" + uId$ + "\mail\outbox"
  Call xmysettings
  If Not isTitleRunning("Agencyprof - POPClient") Then
    Call mailboxinit
    POPTaskID = Shell(s0d + "\AgencyprofPOPClient.exe", vbNormalFocus)
  End If
End If
dir_mailoutbox = form1.s0dir() + "\" + form1.docs() + "\" + uId$ + "\mail\outbox"
umchk$ = "no"
If autocheckmail And upop$ <> "" And upopid$ <> "" And upoppsswd$ <> "" And upopport% <> 0 Then umchk$ = "yes"
If umchk$ = "no" And LCase(upop$) = "dir:inbox" Then umchk$ = "yes"
startup.List1.AddItem iml("Hauptformular initialisieren"): DoEvents

Call Timer2_Timer
Call myuniquedocname("noask")

o% = FreeFile
On Error Resume Next
Open mylocaldatadir() + "\hisl.log" For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  While Not EOF(o%)
    Line Input #o%, l$
    List1.AddItem trm(l$)
  Wend
  Close #o%
End If
List1.AddItem "--" & form1.inmylanguage("Aktuelle Adressen") & "--"
Call showprios
startup.List1.AddItem iml("Adressformular initialisieren"): DoEvents
Call form1.startlog(uId$, "Load shwAdrDetail")
Load shwAdrDetail
Call form1.startlog(uId$, "hide shwAdrDetail")
'shwAdrDetail.Hide
shwAdrDetail.srchit% = 0
sid$ = ""
Call form1.startlog(uId$, "runnig shwAdrDetail.savecheck")
Call shwAdrDetail.savecheck
Call form1.startlog(uId$, "runnig shwAdrDetail.refreshadrdetail(" + trm(sid$) + ")")
Call shwAdrDetail.refreshadrdetail(sid$, "")
shwAdrDetail.Combo3.text = sid$
'Call shwAdrDetail.SetFocus
shwAdrDetail.srchit% = 1
Call form1.startlog(uId$, "reading poplist")
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id FROM poplist where instr(id,'" + uId$ + "_')=1", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
'If Not r.EOF And umchk$ = "yes" Then
If Not r.EOF Then
  pin.Visible = True
  pinlbl.Visible = True
End If
r.Close
End If
'Call setbenutzerdaten("editor", ueditor$)
outl$ = ""
For i% = 1 To 12
  mnams_engl$(i%) = dictionarylookup(mnams$(i%))
  Call form1.startlog(uId$, mnams_engl$(i%) + "=" + mnams$(i%))
Next i%
Label5.ForeColor = lnkcolor
Label7.ForeColor = lnkcolor
Call form1.startlog(uId$, "outlook?")
outl$ = form1.getmyoutlook()
Call form1.startlog(uId$, "pspath")
pspath$ = form1.getusersetting("pspath")
Call form1.startlog(uId$, "docdup")
docdup$ = getusersetting("docdup", "")
docequiv1$ = getusersetting("docequiv", "")
If docequiv1$ = "" Then
  docequiv2$ = ""
Else
  docequiv2$ = Mid(docequiv1$, InStr(docequiv1$, ",") + 1)
  docequiv1$ = Left(docequiv1$, InStr(docequiv1$, ",") - 1)
End If
meinesprache = "de"
If outl$ <> "" Then outlk.Visible = True
Call form1.startlog(uId$, "Externe Mail überprüfen")
startup.List1.AddItem iml("Externe Mail überprüfen"): DoEvents
Set rtmp = New ADODB.Recordset
dirtcol = Val(getusersetting("dirtycolor", "-1"))
If dirtcol = -1 Then dirtcol = RGB(211, 238, 249)
startup.List1.AddItem iml("Übersetzungen durchführen"): DoEvents

shwAdrDetail.Show
Command19.ToolTipText = form1.inmylanguage("Ihr Dokumentenverzeichnis im Explorer öffnen")
Command12.ToolTipText = form1.inmylanguage("Mailsafe")
Command15.ToolTipText = form1.inmylanguage("Kalender öffnen")
altbvorl.ToolTipText = form1.inmylanguage("zu öffnendes Verzeichnis")
Command27.ToolTipText = form1.inmylanguage("Schliesst alle Formulare")
Command29.ToolTipText = form1.inmylanguage("Tageskalender öffnen")
Command25.ToolTipText = form1.inmylanguage("gespeicherte Selektionen")
Command22.ToolTipText = form1.inmylanguage("Saalpläne")
Command18.Caption = "?"
Command18.ToolTipText = form1.inmylanguage("Hilfeseite öffnen")
Check1.Caption = form1.inmylanguage("erweitert")
Check1.ToolTipText = form1.inmylanguage("Auch in Hinweisen suchen")
Command9.ToolTipText = form1.inmylanguage("InternetBrowser öffnen")
Command8.ToolTipText = form1.inmylanguage("Email empfangen")
Command2.ToolTipText = form1.inmylanguage("Ihre Dokumente anzeigen")
Command24.ToolTipText = form1.inmylanguage("Cross Over-Projekte")
Command23.ToolTipText = form1.inmylanguage("Kammermusik-Projekte")
Command20.ToolTipText = form1.inmylanguage("Cross Over-Tourneeangebote")
Command21.ToolTipText = form1.inmylanguage("Kammermusik-Tourneeangebote")
Command28.ToolTipText = form1.inmylanguage("Geburtstagsliste")
Command17.ToolTipText = form1.inmylanguage("Künstler-Tourneeangebote")
'Picture1.ToolTipText = form1.inmylanguage("Neueste Änderungen, Fragen und Antworten")
Command16.ToolTipText = form1.inmylanguage("Email schreiben")
Command15.ToolTipText = form1.inmylanguage("Kalender öffnen")
Command29.ToolTipText = form1.inmylanguage("Tageskalender öffnen")
Command14.ToolTipText = form1.inmylanguage("Verwaltungsfunktionen")
Command13.Caption = form1.inmylanguage("&Adressen")
Command13.ToolTipText = form1.inmylanguage("Formular Adressen öffnen")
Command11.ToolTipText = form1.inmylanguage("To Do-Liste öffnen")
Command7.ToolTipText = form1.inmylanguage("Künstler-Projekte")
List3.ToolTipText = form1.inmylanguage("Alles für heute Relevante aus dem Kalender")
Command10.Caption = form1.inmylanguage("&Programme")
Command10.ToolTipText = form1.inmylanguage("Programme öffnen")
Command6.Caption = form1.inmylanguage("Termine u. &Auftritte")
Command6.ToolTipText = form1.inmylanguage("Termine und Auftritte öffnen")
Command5.ToolTipText = form1.inmylanguage("Orchester-Projekte")
Command4.ToolTipText = form1.inmylanguage("Orchester-Tourneeangebote")
Command3.Caption = form1.inmylanguage("&Werke")
Command3.ToolTipText = form1.inmylanguage("Werkeverzeichnis öffnen")
List2.ToolTipText = form1.inmylanguage("Suche nach Kontaktpersonen")
Command1.ToolTipText = form1.inmylanguage("Auf Wiedersehen!")
List1.ToolTipText = form1.inmylanguage("Suche in allen Feldern")
'sqlmess.Caption = form1.inmylanguage("SQL")
errmess.Caption = form1.inmylanguage("&Fehler")
errmess.ToolTipText = form1.inmylanguage("Achtung! Fehler")
uuid.ToolTipText = form1.inmylanguage("Ihre Benutzerkennung")
pinlbl.Caption = form1.inmylanguage("PIN:")
pinlbl.ToolTipText = form1.inmylanguage("Aktuelles Datum mit Uhrzeit")
Label5.Caption = form1.inmylanguage("   Projekte")
Label5.ToolTipText = form1.inmylanguage("Alle Projekte, Doppelklick öffnet Projektübersicht")
Image6.ToolTipText = form1.inmylanguage("In Kontakten suchen")
Image5.ToolTipText = form1.inmylanguage("In Adressen suchen")
Label7.Caption = form1.inmylanguage("Prioritäten")
Label2.Caption = form1.inmylanguage("   Tourneeangebote")
Label2.ToolTipText = form1.inmylanguage("Alle Tourneeangebote")
Label1.ToolTipText = form1.inmylanguage("Aktuelles Datum mit Uhrzeit")
Label8.Caption = form1.inmylanguage("Heute aktuell")
Label3.Caption = form1.inmylanguage("Suche in Kontakten:")
Label3.ToolTipText = form1.inmylanguage("In Kontakten suchen")
Label6.Caption = form1.inmylanguage("Suchen")
Label6.ToolTipText = form1.inmylanguage("In Adressen suchen")
Label1.Caption = trm(Date & " " & Left(Time, 5))
If getusersetting("adostats", "nein") = "ja" Then Label10.Visible = True
i% = 0
l$ = "-1xxxx"
c$ = "select feldname from auftrittsfelder where instr(feldname,'adrselect.')>0 order by feldname"
Call form1.startlog(uId$, c$)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If r!feldname <> l$ Then
    Call form1.startlog(uId$, l$)
    adrfeldcache$(i%) = LCase(cut_d1(Mid$(r!feldname, 11), "."))
    l$ = (r!feldname)
    i% = i% + 1
  End If
  r.MoveNext
Wend
r.Close
Call form1.startlog(uId$, "Lese spamlist")
Call read_spmlst

Me.Show
On Error Resume Next
form1.SetFocus
On Error GoTo 0
alertdbok = False
Call form1.startlog(uId$, trm(ttabptr%) & " Übersetzungspaare")

startup.List1.AddItem iml("fertig."): DoEvents
If isfieldmissing("opt_prios", "id") Then
  Label7.Visible = False
  edt_prio.Enabled = False
End If
If Not nexist(form1.mylocaldatadir() + "\positions\dayvw.aut") Then Call formload("dayvw")
c$ = s0d$ & "\debug2file_" & uId$ & ".txt"
If Not nexist(c$) Then
  ldtg = (Date + Time) * 86400
  sid$ = basename(c$, ".txt")
  sid$ = s0d$ & "\alt_" & sid$ & "_" & trm(ldtg) & ".txt"
  On Error Resume Next
  Name c$ As sid$
  On Error GoTo 0
End If
On Error Resume Next
Call form1.startlog(uId$, "starte logging debug2file")
Kill c$
On Error GoTo 0
Call form1.startlog(uId$, "strtlog beendet")
starting = False
If form1.getusersetting("debug2file", "nein") = "ja" Then
  dbg2file% = 1
  Call debugwarn
End If
Call dbg2f(App.EXEName & " (" + trm(App.Major) + "." + trm(App.Minor) + " Build #" + trm(App.Revision) + ") gestartet: " & Date & " " & Time)
If fallbackdir$ <> "" Or fallbackserver$ <> "" Then
  Load Datenreplikator
  cb3.Visible = True
  cb4.Visible = True
Else
  fallbackq(1).Visible = False
End If
Timer3.Enabled = True
uuid.ForeColor = form1.lnkcolor
Label6.ForeColor = form1.lnkcolor
Label3.ForeColor = form1.lnkcolor
Label8.ForeColor = form1.lnkcolor
usealtsuch = False
Call dbg2f("Timer4 wird gestartet")
Timer4.Interval = 10000
Timer4.Enabled = True
On Error Resume Next
Call Me.SetFocus
On Error GoTo 0
If usemenu <> "ja" Then
  dat.Visible = False
  edt.Visible = False
  trmn.Visible = False
  wrk.Visible = False
  hlp.Visible = False
End If
If resetfb Then
  Load Datenreplikator
  Call Datenreplikator.Command4_Click
  DoEvents
  Call Datenreplikator.Command27_Click
End If
If connectfb Then
  Load Datenreplikator
  DoEvents
  Call Datenreplikator.Command2_Click
  Call Datenreplikator.Command27_Click
End If
Call dbg2f("form1.load beendet")
Exit Sub

errex:
On Error GoTo 0
Call dbg2f("form1.load Fehlerausgang")
Call Command1_Click

End Sub

Public Sub killlogonexit()
'd2infile = "Form1": d2insub = "killlogonexit"
If getusersetting("killlogonexit", "ja") = "ja" Then
  On Error Resume Next
  Kill s0d$ & "\debug2file_" & uId$ & ".txt"
  Kill s0d$ & "\debug2file_" & uId$ & "_*.txt"
  On Error GoTo 0
End If
End Sub
Public Sub unloadall()
Dim hProc As Long

'd2infile = "Form1": d2insub = "unloadall"
Call killlogonexit
On Error Resume Next
Call shwAdrDetail.mynormsize
Unload bezlist
Unload hordeinfo
Unload listen
Unload ttform
Unload login
Unload landwahl
Unload remedit
Unload prios
Unload startup
Unload higruselect
Unload higruselect2
Unload pplan
Unload sels
Unload dayvw
Unload mexplore
Unload tabkalk
Unload AutoAnwahl
Unload bplan
Unload splan
Unload kalku
Unload imgs
Unload abos
Unload besetzung
Unload msafe
Unload kbuch
Unload fdet
Unload einstellungen
Unload werkvz
Unload trvw
Unload taliste
Unload vwr
Unload dialselect
Unload kvk
Unload adrtypselector
Unload rtfview
Unload frmMain
Unload frmTrace
Unload dochist2
Unload frmBrowser
Unload tplan
Unload tpzoom
Unload adrselect
Unload waehrung
Unload prog
Unload auftrittshintergrund
Unload auftritt
Unload auftrittrepeat
Unload verwaltung
Unload verwalt_public
Unload verwalt_sbf
Unload auftrittselect
Unload multilineinput
Unload todolist
Unload alarmlist
Unload import
Unload fehler
Unload colorsel
Unload kc
Unload k3
Unload create2do
Unload smtp
Unload fselect
Unload emailadrselect
Unload dbupgrade
Unload memoview
Unload agx
Unload kurse
Unload zusinf
Unload repertoire
On Error GoTo 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'this procedure receives the callbacks from the System Tray icon.
If Not ismin Then Exit Sub

      Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
       End If
       Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
         Call mPopRestore_Click
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Call mPopRestore_Click
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hWnd)
         Me.PopupMenu Me.mPopupSys
       End Select
End Sub

Private Sub Form_Resize()
axsResizer1.Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rrr
Dim tr$, o%, i%

'd2infile = "Form1": d2insub = "Form_Unload"
break% = 1
'Call k2.setbreak
Call killxmysettings
If Not ismin Then
  Call unloadall
  Unload shwAdrDetail
  Unload Datenreplikator
End If
Shell_NotifyIcon NIM_DELETE, nid
DoEvents

o% = FreeFile
On Error Resume Next
Open mylocaldatadir() + "\hisl.log" For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 And Combo1.ListCount > 0 Then
  For i% = 0 To Combo1.ListCount - 1
    If Left(Combo1.List(i%), 2) <> "--" Then Print #o%, Combo1.List(i%)
  Next i%
  Close #o%
End If
tr$ = Dir("sql???.*")
While tr$ <> ""
  If InStr(tr$, ".bat") > 0 Or InStr(tr$, ".txt") > 0 Then
    On Error Resume Next
    Kill tr$
    On Error GoTo 0
  End If
  tr$ = Dir
Wend
On Error Resume Next
Kill s0d$ & "\" & uId$ & ".run"
On Error GoTo 0
Call killlogonexit
Hide
On Error GoTo exuld
If Not ismin Then
  Call form1.setmylasttop(Me.name, Me.Top)
  Call form1.setmylastleft(Me.name, Me.Left)
  Call form1.setmylastwidth(Me.name, Me.Width)
  Call form1.setmylastheight(Me.name, Me.Height)
End If
exuld:
On Error GoTo 0
End

End Sub


Private Sub hlp_help_Click()
Call Command18_Click
End Sub

Private Sub Image5_Click()
'd2infile = "Form1": d2insub = "Image5_Click"
Call Label6_DblClick
End Sub



Private Sub Image6_DblClick()
'd2infile = "Form1": d2insub = "Image6_DblClick"
Call Label6_DblClick
End Sub

Private Sub Label3_DblClick()
'd2infile = "Form1": d2insub = "Label3_DblClick"
Call Label6_DblClick
End Sub

Private Sub Label5_dblClick()

'd2infile = "Form1": d2insub = "Label5_dblClick"
Load pplan
On Error Resume Next
Call pplan.SetFocus
On Error GoTo 0
End Sub

Public Sub Label6_DblClick()
'd2infile = "Form1": d2insub = "Label6_DblClick"
Load adrselect
adrselect.SetFocus
End Sub

Private Sub Label7_DblClick()
'd2infile = "Form1": d2insub = "Label7_DblClick"
If form1.isfieldmissing("opt_prios", "id") Then Exit Sub
Load prios
On Error Resume Next
Call prios.SetFocus
On Error GoTo 0

End Sub

Private Sub Label8_Click()
'd2infile = "Form1": d2insub = "Label8_Click"
Call rlist3
End Sub

Private Sub Label8_DblClick()
'd2infile = "Form1": d2insub = "Label8_DblClick"
Call Command29_Click
End Sub

Private Sub List1_Click()
Dim sid$, c$, na$, i%, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "List1_Click"
List2.Clear
sid$ = List1.List(List1.ListIndex)
If InStr(sid$, "::") > 1 Then sid = Left$(sid$, InStr(sid$, "::") - 1)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer

c$ = "SELECT id,name,position FROM kontakt where vid ='" + sid$ + "'"
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not rtmp.EOF Then
rtmp.MoveFirst
While Not rtmp.EOF
  If Not IsNull(rtmp!name) Then
    na$ = rtmp!name
    If trm(rtmp!Position) <> "" Then na$ = na$ + " (" + rtmp!Position + ")"
    List2.AddItem form1.crlffake(na$) & Space$(80) & "ID:" & rtmp!id
  End If
  rtmp.MoveNext
Wend
End If
rtmp.Close
If sid$ = "--" & form1.inmylanguage("Aktuelle Adressen") & "--" Then
  List1.Clear
  Call rlist3
  For i% = 0 To List3.ListCount - 1
    If InStr(List3.List(i%), form1.inmylanguage("Projekt") & ": ") = 1 Then
      sid$ = "select * from tplan where id='" & Mid(List3.List(i%), 10) & "'"
      Set rtmp = New ADODB.Recordset
      rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, sid$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not rtmp.EOF
        sid$ = trm(rtmp!orchester): If sid$ <> "" Then Call l1a(sid$)
        sid$ = trm(rtmp!veranstalter): If sid$ <> "" Then Call l1a(sid$)
        sid$ = trm(rtmp!Solist): If sid$ <> "" Then Call l1a(sid$)
        sid$ = trm(rtmp!projektbetreuer): If sid$ <> "" Then Call l1a(sid$)
        sid$ = trm(rtmp!tourneeleitung): If sid$ <> "" Then Call l1a(sid$)
        sid$ = trm(rtmp!dirigent): If sid$ <> "" Then Call l1a(sid$)
        rtmp.MoveNext
      Wend
    End If
  Next i%
  Call showprios
End If
End Sub

Public Sub List1_DblClick()
Dim sid$
'd2infile = "Form1": d2insub = "List1_DblClick"
shwAdrDetail.srchit% = 0
If uId$ = "" Then Exit Sub
sid$ = trm(List1.List(List1.ListIndex))
If InStr(sid$, "::") > 1 Then sid$ = Left$(sid$, InStr(sid$, "::") - 1)
usealtsuch = True
Load shwAdrDetail
Call shwAdrDetail.savecheck
Call shwAdrDetail.refreshadrdetail(sid$, "")
Call c1add(sid$)
shwAdrDetail.Combo3.text = sid$
Call shwAdrDetail.SetFocus
shwAdrDetail.srchit% = 1
End Sub


Private Sub List2_DblClick()
Dim sid$, p%, vid$, cid$, c$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "List2_DblClick"
If uId$ = "" Then Exit Sub
sid$ = List2.List(List2.ListIndex)
p% = InStr(sid$, "ID:")
If p% > 0 Then
  cid$ = trm(Left$(sid$, p% - 1))
  sid$ = Mid$(sid$, p% + 3)
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  c$ = "SELECT vid FROM kontakt where id='" + sid$ + "'"
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rtmp.EOF Then Exit Sub
  rtmp.MoveFirst
  If Not IsNull(rtmp!vid) Then
    vid$ = rtmp!vid
  Else
    vid$ = "-1"
  End If
  rtmp.Close
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(vid$, cid$)
  Call c1add(vid$)
  Call shwAdrDetail.SetFocus
End If
End Sub

Private Sub List3_Click()
Dim id$, c$, r As ADODB.Recordset, tpid$, dtg$
Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "List3_Click"
id$ = List3.List(List3.ListIndex)
If List3.ToolTipText = transe("Geburtstagsliste") Then
  dtg$ = cut_d1(id$, ":")
  tpid$ = Mid$(id$, InStr(id$, "(ID:") + 4)
  c$ = cut_d2bis(id$, ":"): c$ = cut_d1(c$, ",")
  Combo1.text = c$
End If

End Sub

Public Sub List3_DblClick()
Dim id$, c$, r As ADODB.Recordset, tpid$, dtg$, rrr
Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "List3_DblClick"
id$ = List3.List(List3.ListIndex)
If List3.ToolTipText = transe("Geburtstagsliste") Then
  dtg$ = cut_d1(id$, ":")
  tpid$ = Mid$(id$, InStr(id$, "(ID:") + 4)
  c$ = "select * from auftritt where datum>'" + datum2sql(Date) + "' and auftrittstyp='Geburtstag' and instr(bezeichnung,'" & tpid$ & "')>0"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
      id$ = r!id
  Else
      id$ = form1.newid("auftritt", "id", 20)
      form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 id$ + "','-1'" + _
                 ",'Geburtstag','" & tpid$ & "','" + _
                 dtg$ + "')")
      DoEvents
  End If
  Unload auftritt
  DoEvents
  Load auftritt
  Call auftritt.SetFocus
  Call auftritt.showrec(id$, 0)
Else
  break% = 1
  If id$ = form1.inmylanguage("Heute") Then
    Load kc
    Call kc.settag0(Date)
    If form1.getusersetting("kalenderimmeramersten", "nein") = "ja" Then kc.Text1.text = 1
  Else
    If InStr(id$, form1.inmylanguage("Projekt") & ": ") = 0 Then
      id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
      Unload auftritt
      DoEvents
      Load auftritt
      Call auftritt.SetFocus
      Call auftritt.showrec(id$, 0)
    Else
      tpid$ = Mid$(id$, 10)
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
  End If
End If

End Sub

Private Sub suchen_Click()
'd2infile = "Form1": d2insub = "suchen_Click"
Call Command8_Click
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'd2infile = "Form1": d2insub = "List3_MouseDown"
If Button = 2 Then Call rlist3
End Sub

Private Sub mPopRestore_Click()
'called when the user clicks the popup menu Restore command
Dim Result As Long
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hWnd)
mPopupSys.Enabled = False
mPopExit.Enabled = False
mPopRestore.Enabled = False
ismin = False
Shell_NotifyIcon NIM_DELETE, nid
DoEvents
Me.Show
Load shwAdrDetail
End Sub

Private Sub mPopExit_Click()
       'called when user clicks the popup menu Exit command
       Call mPopRestore_Click
       Unload Me
End Sub


Private Sub outlk_Click()
Dim X, outl$


'd2infile = "Form1": d2insub = "outlk_Click"
outl$ = getmyoutlook()
If outl$ = "" Then Exit Sub
On Error Resume Next
X = Shell(outl$, 1)
On Error GoTo 0

End Sub


Private Sub Picture1_Click()

'd2infile = "Form1": d2insub = "Picture1_Click"
  Unload frmBrowser
  DoEvents
  frmBrowser.StartingAddress = "http://www.agencyprof.de/download/update/changelog.txt"
  Load frmBrowser

End Sub

Public Function ask_agencyprof_com(utf8$) As String
Dim c$

  ask_agencyprof_com = ""
  c$ = LCase(getusersetting("donotaskap.com", "no"))
  If c$ = "ja" Or c$ = "yes" Then Exit Function
  If trm(utf8$) = "" Then Exit Function
  Unload frmBrowser
  DoEvents
  brwhidden = True
  frmBrowser.StartingAddress = "http://ask.agencyprof.com/askutf828559-1.php?ut=" + utf8$
  Load frmBrowser
  DoEvents
  While frmBrowser.brwWebBrowser.Busy
    Call delay(200)
  Wend
  ask_agencyprof_com = cut_d1(cut_d2bis(frmBrowser.brwWebBrowser.Document.Body.innerhtml, "|"), "|")
  brwhidden = False
  Unload frmBrowser
End Function

Public Sub combo1_Change()
Dim l$, w$, c$, sgrp As String, swert As String
Dim r As ADODB.Recordset, rrr

If Not use_adrsuchtimer Then Exit Sub

break% = 1
If uId$ = "" Then Exit Sub
If Left(Combo1.text, 1) = ":" Then
  c$ = trm(Combo1.text): If Len(c$) < 2 Then Exit Sub
  c$ = Mid$(c$, 2)
  sgrp = cut_d1(c$, ":")
  swert = cut_d2bis(c$, ":")
  If sgrp <> "" Then
    c$ = "SELECT auftrittsfelder.FeldName FROM adresstypen INNER JOIN auftrittsfelder ON adresstypen.id = auftrittsfelder.typ "
    c$ = c$ + "Where auftrittsfelder.feldname Like '" + sgrp + "%' ORDER BY auftrittsfelder.FeldName"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
    l$ = ""
    While Not r.EOF
      c$ = trm(r!feldname)
      If InStr(LCase(l$), "|" + LCase(c$)) = 0 Then l$ = l$ + "|" + c$
      r.MoveNext
    Wend
    If swert <> "" Then
      l$ = Mid$(l$, 2)
      c$ = "select auftrittsid from auftritthigru where ("
      While l$ <> ""
        w$ = cut_d1(l$, "|")
        l$ = cut_d2bis(l$, "|")
        c$ = c$ + "FeldName='" + w$ + "' "
        If l$ <> "" Then c$ = c$ + " or "
      Wend
      c$ = c$ + ") and Felddaten like '%" + swert + "%'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
      rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
      List1.Clear
      While Not r.EOF
        w$ = trm(r!auftrittsid)
        List1.AddItem w$
        r.MoveNext
      Wend
    End If
    Exit Sub
  End If
End If
If Left(Combo1.text, 4) <> "typ:" Then
  Timer1.Enabled = False
  Timer1.Interval = 500
  Timer1.Enabled = True
  snotb4 = now() + usuchvz * msec
End If
End Sub
Public Sub setdbpsswd(p$)
'd2infile = "Form1": d2insub = "setdbpsswd"
dbpsswd$ = p$
End Sub
Public Sub setdbuid(p$)
'd2infile = "Form1": d2insub = "setdbuid"
dbuid$ = p$
End Sub

Public Sub sethomepath(p$)
hppth$ = p$
End Sub
Public Sub setloginname(u$)
Dim rtmp As ADODB.Recordset, c$, l$, M$
Dim conn As ADODB.Connection, ukalalways%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "setloginname"
uId$ = u$

If u$ = "_LOGOUT_" Then
  Call Command1_Click
  Exit Sub
End If
Load startup
Call startup.SetFocus
startup.List1.AddItem iml("Benutzer") + " " + u$ & " " + iml("wird angemeldet"): DoEvents
c$ = adopara$: M$ = ""
While c$ <> ""
  l$ = cut_d1(c$, ";"): c$ = cut_d2bis(c$, ";")
  If Left$(LCase(l$), 4) <> "pwd=" Then M$ = M$ + l$ + ";"
Wend
Call form1.startlog(uId$, "dbpara=" + M$)
Set conn = New ADODB.Connection
conn.ConnectionString = adopara$
conn.Open

startup.List1.AddItem iml("Benutzerdaten lesen"): DoEvents
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open "SELECT * FROM benutzerdaten where id='" & uId$ & "'", conn, adOpenDynamic, adLockReadOnly

If Not rtmp.EOF Then
  If Not IsNull(rtmp!fax) Then ufax$ = trm(rtmp!fax)
  Call form1.startlog(uId$, "fax=" + ufax$)
  If Not IsNull(rtmp!faxvorlage) Then ufaxrtf$ = rtmp!faxvorlage
  Call form1.startlog(uId$, "faxvorlage=" + ufaxrtf$)
  If Not IsNull(rtmp!name) Then uname$ = rtmp!name
  Call form1.startlog(uId$, "name=" + uname$)
  On Error Resume Next
  If Not IsNull(rtmp!editor) And trm(rtmp!editor) <> "" Then
    ueditor$ = rtmp!editor
  Else
    'ueditor$ = wordsuchen()
    ueditor = GetWordPath()
    If ueditor$ = "" Then ueditor$ = "write"
  End If
  On Error GoTo 0
  Call form1.startlog(uId$, "editor=" + ueditor$)
  If Not IsNull(rtmp!bildeditor) Then uphedit$ = rtmp!bildeditor
  Call form1.startlog(uId$, "bildeditor=" + uphedit$)
  If Not IsNull(rtmp!netscape47inbox) Then netscape47inbox$ = strrepl(rtmp!netscape47inbox, """", "")
  Call form1.startlog(uId$, "netscape47inbox=" + netscape47inbox$)
  If Not IsNull(rtmp!mailclient) Then mailclient$ = rtmp!mailclient
  Call form1.startlog(uId$, "mailclient=" + mailclient$)
  If Not IsNull(rtmp!immer_speichern) Then usavealways$ = rtmp!immer_speichern
  Call form1.startlog(uId$, "immer_speichern=" + usavealways$)
  If Not IsNull(rtmp!immer_kalender) Then ucalalways$ = rtmp!immer_kalender
  Call form1.startlog(uId$, "immer_kalender=" + ucalalways$)
  If Not IsNull(rtmp!erster_wochentag) Then ufdow$ = rtmp!erster_wochentag
  Call form1.startlog(uId$, "erster_wochentag=" + ufdow$)
  If Not IsNull(rtmp!mailserver) Then umsrvout$ = rtmp!mailserver
  Call form1.startlog(uId$, "mailserver=" + umsrvout$)

  If Not IsNull(rtmp!email) Then umailadr$ = rtmp!email
  Call form1.startlog(uId$, "email=" + umailadr$)

  If Not IsNull(rtmp!browser) Then ubrowse$ = rtmp!browser
  Call form1.startlog(uId$, "browser=" + ubrowse$)

  If Not IsNull(rtmp!MySQL) Then umysqld$ = rtmp!MySQL
  Call form1.startlog(uId$, "mysql=" + umysqld$)

  If Not IsNull(rtmp!mysqlhost) Then umysqlhost$ = rtmp!mysqlhost
  Call form1.startlog(uId$, "mysqlhost=" + umysqlhost$)

  If Not IsNull(rtmp!werkepool) Then uwpool$ = LCase(trm(rtmp!werkepool))
  Call form1.startlog(uId$, "werkepool=" + uwpool$)

  uwantstooltips% = 0
  If Not IsNull(rtmp!tooltips) Then
    If rtmp!tooltips = "ja" Then
      uwantstooltips% = 1
    End If
  End If
  Call form1.startlog(uId$, "tooltips=" + trm(uwantstooltips%))

  ukalalways% = 0
  If Not IsNull(rtmp!changelog) Then uclog$ = rtmp!changelog
  Call form1.startlog(uId$, "changelog=" & trm(uclog$))
  usuchvz = 600#:
  On Error Resume Next
  If Not IsNull(rtmp!SuchVerzoegerung) Then usuchvz = var2dbl(trm(rtmp!SuchVerzoegerung))
  If trm(usuchvz) = "" Then usuchvz = 600
  If usuchvz < 10 Then usuchvz = 10
  Call form1.startlog(uId$, "SuchVerzoegerung=" & trm(usuchvz))
  On Error GoTo 0
End If
login.login_exv% = 2
conn.Close
Unload login

End Sub

Public Function getuserid() As String

'd2infile = "Form1": d2insub = "getuserid"
getuserid = uId$

End Function
Public Function getmyhomepg() As String

'd2infile = "Form1": d2insub = "getmyhomepg"
getmyhomepg = getusersetting("agpkalurl", "")
If getmyhomepg <> "" Then Exit Function
If ubrowse$ <> "" Then
  getmyhomepg = ubrowse$
Else
  getmyhomepg = "http://www.Agencyprof.de"
End If

End Function
Public Function getmyeditor(ex$) As String
Dim xt$, edt$

'd2infile = "Form1": d2insub = "getmyeditor"
xt$ = LCase(ex$)
edt$ = getusersetting("take4" + xt$, "")
If edt$ <> "" Then
  If LCase(edt$) = "outlook" Then edt$ = GetOutlookPath()
  If edt$ <> "" Then
    getmyeditor = edt$
    Exit Function
  End If
End If
If exist(ueditor$) = 0 And InStr(ueditor$, "start ") = 0 Then
  ueditor$ = form1.getusersetting("editor2")
  If exist(ueditor$) = 0 Then ueditor$ = "write"
End If
getmyeditor = ueditor$

End Function
Public Function getmybildeditor() As String

'd2infile = "Form1": d2insub = "getmybildeditor"
getmybildeditor = uphedit$

End Function

Public Function getmymysqld() As String

'd2infile = "Form1": d2insub = "getmymysqld"
getmymysqld = umysqld$

End Function
Public Function getmymysqlhost() As String

'd2infile = "Form1": d2insub = "getmymysqlhost"
getmymysqlhost = dbserver$

End Function

Public Function getsuchvz() As Double

'd2infile = "Form1": d2insub = "getsuchvz"
getsuchvz = usuchvz

End Function

Public Sub faxan(adrid$, kid$, vorlage$, betreff$, volltext$, trgpth$, opt1$)
Dim pb%, cmd$, poaname$
Dim o%, p%, nam$, betr$, land$, plz$, ort$, plzort$, meinzeichen$, eml$, apostanrede$
Dim rtmp As ADODB.Recordset, udat As ADODB.Recordset, kabt$, pfadr As Boolean, pferg$, adressname$
Dim hdat As ADODB.Recordset, pout$
Dim ddtt As Long, terg As Long, trerg As Long, twerg As Long, brerg As Long, tz$, anredename$, postanrede$
Dim stra$, kstra$, postfach$, plzpostfach$, tel$, fax$, knam$, kpostanrede$, q%, t$, orgt$
Dim kplzort$, anred$, abred$, rrr, lll, opt$, fn$, c$, rechbez$, l$, t0$, marke$
Dim rev$, ttest$, rfeld$, rwert$, i%, uv$, erg$, ln$, vorlfn$, chgadr As Boolean, chgkadr As Boolean
Dim rtmppostfach As String
Dim rtmpplzpostfach As String, tmpl1$, tmpl2$
Dim kland$, kplz$, kort$, kpostfach$, kplzpostfach$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "faxan"
Call tm_start(0)
Call tm_start(1)
pout$ = ""
c$ = HinweisVonAuftritt(adrid$)
If InStr(c$, "!keine Briefe!") > 0 Then
  MsgBox "Es soll kein Brief/Faxerstellt werden (siehe Hinweise unter Kategorie Person der Adresse)."
  Exit Sub
End If
chgadr = False
If Not form1.isfieldmissing("opt_adresspool", "id") Then
  If InStr(LCase(opt1$), "elseadr") > 0 Then chgadr = True
End If
chgkadr = False
If kid$ <> "-1" Then
  If Not form1.isfieldmissing("opt_adresspool", "id") Then
    If InStr(LCase(opt1$), "elsekadr") > 0 Then chgkadr = True
  End If
  c$ = HinweisVonAuftritt(adrid$ + c$)
  If InStr(c$, "!keine Briefe!") > 0 Then
    MsgBox "Es soll kein Brief/Faxerstellt werden (siehe Hinweise unter Person dieses Kontaktes)."
    Exit Sub
  End If
End If
vorlfn$ = s0dir & "\" & dbname$ & ".rtf\" & vorlage$
If vorlagencache <> "" Then
  l$ = vorlfn$: rrr = 0
  vorlfn$ = vorlagencache + "\" + vorlage$
  If nexist(vorlfn$) Then
    On Error Resume Next
    Call FileCopy(l$, vorlfn$)
    rrr = Err
    On Error GoTo 0
  End If
  If rrr <> 0 Or nexist(vorlfn$) Then
    vorlagencache = ""
    vorlfn$ = s0dir & "\" & dbname$ & ".rtf\" & vorlage$
  End If
End If
If exist(vorlfn$) = 0 Then
  MsgBox "Vorlage unbekannt: " + vorlage$
  Exit Sub
End If
pfadr = False
pferg$ = getusersetting("postfachergänzen", "")
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adresse where id ='" + adrid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
land$ = "": plz$ = "": ort$ = "": plzort = "": nam$ = ""
kland$ = "": kplz$ = "": kort$ = "": kplzort = ""
If Not rtmp.EOF Then
  nam$ = trmx1(rtmp!postanrede)
  adressname$ = nam$
  tz$ = getusersetting("anredetrennung", " "): If tz$ = "|" Then tz$ = vbCrLf
  anredename$ = strrepl(trmx1(nam$ & tz$ & trmx1(rtmp!name)), "  ", " ")
  nam$ = trmx1(rtmp!name)
  apostanrede$ = trmx1(rtmp!postanrede)
  eml$ = trm(rtmp!email)
  rtmppostfach = trm(rtmp!postfach): If chgadr Then rtmppostfach = trm(shwAdrDetail.postf.text)
  rtmpplzpostfach = trm(rtmp!plzpostfach): If chgadr Then rtmpplzpostfach = trm(shwAdrDetail.plzp.text)
  postanrede$ = apostanrede$
  stra$ = trm(rtmp!strasse): If chgadr Then stra$ = trm(shwAdrDetail.datf(2).text)
  If shwAdrDetail.Check3.value = 1 And trm(rtmppostfach) <> "" And trm(rtmpplzpostfach) <> "" Then
    stra$ = trm(rtmppostfach)
    pfadr = True
    If pferg$ <> "" Then
      If InStr(LCase(stra$), pferg$) = 0 Then
        stra$ = pferg$ & " " & stra$
      End If
    End If
  End If
  land$ = trm(rtmp!land): If chgadr Then land$ = trm(shwAdrDetail.datf(14).text)
  If LCase(land$) = LCase(getusersetting("meinland")) Then land = ""
  If Not IsNull(rtmp!plz) Then plz$ = rtmp!plz
  If chgadr Then plz$ = trm(shwAdrDetail.datf(13).text)
  If pfadr Then plz$ = trm(rtmpplzpostfach)
  ort$ = trm(rtmp!ort): If chgadr Then ort$ = trm(shwAdrDetail.datf(3).text)
  postfach$ = trm(rtmppostfach)
  plzpostfach$ = trm(rtmpplzpostfach)
  If Not IsNull(rtmp!tel) Then tel$ = rtmp!tel
  If Not IsNull(rtmp!fax) Then fax$ = rtmp!fax
End If
plzort$ = form1.getplzort(land$, plz$, ort$)
meinzeichen$ = getusersetting("meinzeichen")
If meinzeichen$ = "" Then meinzeichen$ = initialen(uname$)
Set udat = New ADODB.Recordset
udat.CursorLocation = adUseServer
rrr = form1.adoopen(udat, "SELECT * FROM benutzerdaten where id ='" + uId$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
knam$ = ""
If kid$ <> "-1" Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    tz$ = getusersetting("anredetrennung", " "): If tz$ = "|" Then tz$ = vbCrLf
    kabt$ = getkontaktabteilungbyid(kid$)
    knam$ = trmx1(rtmp!name)
    If knam$ = nam$ Or knam$ = anredename$ Then
      knam$ = ""
    Else
      If kabt$ <> "" Then kabt$ = kabt$ & vbCrLf
      knam$ = kabt$ & trmx1(trmx1(rtmp!postanrede) & " " & knam$)
    End If
    If trm(rtmp!tel) <> "" Then tel$ = rtmp!tel
    If trm(rtmp!fax) <> "" Then fax$ = rtmp!fax
    If trm(rtmp!email) <> "" Then eml$ = trm(rtmp!email)
    kstra$ = trm(rtmp!strasse): If chgkadr Then kstra$ = trm(shwAdrDetail.kadat(0).text)
    If shwAdrDetail.Check3.value = 1 And trm(rtmp!postfach) <> "" And trm(rtmp!plzpostfach) <> "" Then
      kstra$ = trm(rtmp!postfach): If chgkadr Then kstra$ = trm(shwAdrDetail.kadat(5).text)
      pfadr = True
      If pferg$ <> "" Then
        If InStr(LCase(kstra$), LCase(pferg$)) = 0 Then
          kstra$ = pferg$ & " " & kstra$
        End If
      End If
    End If
    If kstra$ <> "" Then stra$ = kstra$
    kland$ = trm(rtmp!lkz): If chgkadr Then kland$ = trm(shwAdrDetail.kadat(5).text)
    If kland$ <> "" Then land$ = kland$
    kpostanrede$ = trmx1(rtmp!postanrede)
    'postanrede$ = kpostanrede$
    If LCase(land$) = LCase(getusersetting("meinland")) Then land = ""
    kplz$ = trm(rtmp!plz): If chgkadr Then kplz$ = trm(shwAdrDetail.kadat(2).text)
    If pfadr Then
      If trm(rtmp!plzpostfach) <> "" Then kplz$ = trm(rtmp!plzpostfach)
      If chgkadr Then kplz$ = trm(shwAdrDetail.kadat(4).text)
    End If
    If kplz$ <> "" Then plz$ = kplz$
    kort$ = trm(rtmp!ort): If chgkadr Then kort$ = trm(shwAdrDetail.kadat(3).text)
    If kort$ <> "" Then ort$ = kort$
    kplzort$ = plzort$
    kpostfach$ = trm(rtmp!postfach): If chgkadr Then kpostfach$ = trm(shwAdrDetail.kadat(5).text)
    If kpostfach$ <> "" Then postfach$ = kpostfach$
    kplzpostfach$ = trm(rtmp!plzpostfach): If chgkadr Then kplzpostfach$ = trm(shwAdrDetail.kadat(4).text)
    If kplzpostfach$ <> "" Then plzpostfach$ = kplzpostfach$
    plzort$ = form1.getplzort(land$, plz$, ort$)
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='" + kid$ + "' and user='" + anredeuser$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
    If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
  Else
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='" + kid$ + "' and user='system'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rtmp.EOF Then
      If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
      If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
    End If
  End If
Else
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='-1." + adrid$ + "' and user='" + anredeuser$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
    If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
  Else
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='-1." + adrid$ + "' and user='system'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rtmp.EOF Then
      If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
      If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
    End If
  End If
End If
If anred$ = "" Then anred$ = getusersetting("StandardAnrede", "")
If abred$ = "" Then abred$ = getusersetting("StandardAbrede", "")
terg = tm_stop(0)
Call dbg2f("Vorbereitung faxan: " + trm(terg) + " ms")
Call tm_start(0)
o% = FreeFile
On Error Resume Next
Call form1.dbg2f("vorlage=" + vorlfn$)
Open vorlfn$ For Input As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Vorlage " & lll & " kann nicht geöffnet werden."
  Exit Sub
End If
Call form1.dbg2f("Vorlage öffnen: " + trm(tm_stop(0)) + " ms")
opt$ = ""
If InStr(LCase(opt1$), "defaultname") > 0 Then opt$ = "noask"
tm_start (0)
If trgpth$ = "" Then
  fn$ = myuniquedocname(opt$)
Else
  fn$ = myuniquedocnameinpath(trgpth$, opt$)
End If
Call form1.dbg2f("myuniquedocname: " + trm(tm_stop(0)) + " ms")
If Len(trm(fn$)) = 0 Then Exit Sub
If trm(betreff$) = "" Then
  betr$ = InputBox(transe("Betreff:"), transe("Beschreibung"), betr$, 100, 100)
Else
  betr$ = betreff$
End If
If Len(trm(betr$)) = 0 Then Exit Sub
If trm(adrid$) <> "" And trm(betr$) <> "" And trm(vorlage$) <> "" And trm(fn$) <> "" Then
  Call tm_start(0)
  c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
              form1.newid("dochist", "id", 19) & "','" & _
              adrid$ & "','" & kid$ & "','" & _
              fn$ & "','" & _
              datum2sql(Date) & " " & Time & "','" & _
              uId$ & "','" & _
              Left$(strrepl(betr$, "'", "´"), 70) & "','" & _
              vorlage$ & "')"
  Call form1.sqlqry(c$)
  Call form1.dbg2f("Kontakthistorie: " + trm(tm_stop(0)) + " ms")
End If
rechbez$ = "Dokument an " + adrid$: If trm(knam$) <> "" Then rechbez$ = rechbez$ + " " + trm(knam$)
p% = FreeFile
trerg = 0: twerg = 0: brerg = 0
Call tm_start(0)
MousePointer = 11: DoEvents
Open fn$ For Output As #p%
While Not EOF(o%)
  Call tm_start(3)
  Line Input #o%, l$
  ddtt = tm_stop(3)
  trerg = trerg + ddtt
  brerg = brerg + Len(l$)
  While Len(l$) > 0
    q% = InStr(l$, bkmstart$)
    If q% > 0 Then
      If pout$ <> "" Then
        Print #p%, pout$;
        pout$ = ""
      End If
      t$ = Mid$(l$, q% + Len(bkmstart$))
      t$ = Left$(t$, InStr(t$, "}") - 1)
      orgt$ = t$
      t0$ = t$
      t$ = LCase(t0$)
      If Left$(t0$, 5) = "MARKE" Then
        If isdigit(Mid$(t0$, 6, 1)) <> 0 Then
          If Mid$(t0$, 7, 1) = "_" Then
            marke$ = Left$(t0$, 7)
            t$ = Mid$(t$, 8)
          End If
        End If
      End If
      If Left$(t0$, 1) = "M" And Mid$(t0$, 3, 1) = "_" Then
        If Mid$(t0$, 3, 1) = "_" Then
          marke$ = Left$(t0$, 3)
          t$ = Mid$(t$, 4)
        End If
      End If
      Call tm_start(3)
      Select Case LCase(t$)
        Case "name": Print #p%, Left$(l$, q% - 1);
                   If tz$ <> "nurdername" Then
                     tmpl1$ = strrepl(nam$, vbCrLf, "\par ")
                   End If
                   If knam$ <> "" Then
                     tmpl2$ = strrepl(knam$, vbCrLf, "\par ")
                   End If
                   If getusersetting("adressfolge", "") = "kontakt-name" Then
                     Print #p%, tmpl2$;
                     If tmpl1$ <> "" Then
                       Print #p%, "\par "; tmpl1$;
                     End If
                   Else
                     If tmpl1$ <> "" Then
                       Print #p%, tmpl1$;
                     End If
                     If tmpl2$ <> "" Then
                       If tmpl1$ <> "" Then Print #p%, "\par ";
                       Print #p%, tmpl2$;
                      End If
                   End If
        Case "anredename":
                    Print #p%, Left$(l$, q% - 1);
                    If tz$ <> "nurdername" Then
                      Print #p%, strrepl(strrepl(anredename$, vbCrLf, "\par "), "  ", " ");
                    End If
                   If knam$ <> "" Then
                     If tz$ = vbCrLf Then
                       Print #p%, strrepl(strrepl(tz$, vbCrLf, "\par "), "  ", " ");
                     End If
                     Print #p%, strrepl(strrepl(knam$, vbCrLf, "\par "), "  ", " ");
                   End If
        Case "adressname":
                    Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(adressname$, vbCrLf, "\par ");
        Case "strasse": Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(stra$, vbCrLf, "\par ");
        Case "ort": Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(strrepl(plzort$, vbCrLf, "\par "), "  ", " ");
        Case "nurort": Print #p%, Left$(l$, q% - 1): Print #p%, strrepl(ort$, vbCrLf, "\par ");
        Case "land": Print #p%, Left$(l$, q% - 1);: Print #p%, land$;
        Case "plz": Print #p%, Left$(l$, q% - 1);: Print #p%, plz$;
        Case "tel": Print #p%, Left$(l$, q% - 1);: Print #p%, userformatphone(tel$);
        Case "email": Print #p%, Left$(l$, q% - 1);: Print #p%, eml$;
        Case "fax": Print #p%, Left$(l$, q% - 1);: Print #p%, userformatphone(fax$);
        Case "postfach": Print #p%, Left$(l$, q% - 1);: Print #p%, postfach$;
        Case "plzpostfach": Print #p%, Left$(l$, q% - 1);: Print #p%, plzpostfach$;
        Case "anrede": Print #p%, Left$(l$, q% - 1);: Print #p%, anred$;
        Case "apostanrede": Print #p%, Left$(l$, q% - 1);: Print #p%, apostanrede$;
        Case "kpostanrede": Print #p%, Left$(l$, q% - 1);: Print #p%, kpostanrede$;
        Case "postanrede": Print #p%, Left$(l$, q% - 1);: Print #p%, postanrede$;
        Case "postanredename":
                           If trm(postanrede$) <> "" Then
                             If Right$(postanrede$, 1) <> " " Then postanrede$ = postanrede$ + " "
                           End If
                           Print #p%, Left$(l$, q% - 1);
                           poaname$ = strrepl(postanrede$ + " " + strrepl(nam$, vbCrLf, "\par "), "  ", " ")
                           Print #p%, poaname$;
                           If knam$ <> "" Then
                             Print #p%, "\par " + strrepl(strrepl(knam$, vbCrLf, "\par "), "  ", " ");
                           End If
        Case "abrede": Print #p%, Left$(l$, q% - 1);: Print #p%, abred$;
        Case "betreff": Print #p%, Left$(l$, q% - 1);: Print #p%, betr$;
        Case "meinzeichen": Print #p%, Left$(l$, q% - 1);: Print #p%, meinzeichen$;
        Case "text": Print #p%, Left$(l$, q% - 1);: Print #p%, repl1310rtf(volltext$);
        Case Else
      End Select
      twerg = twerg + tm_stop(3)
      If InStr(t$, "__") > 0 Then
        rev$ = Mid$(t$, InStr(t$, "__") + 2)
        ttest$ = Left$(t$, InStr(t$, "__") - 1)
      End If
      If ttest$ = "rel" Then
        If InStr(rev$, "__") > 0 Then
          rfeld$ = Mid$(rev$, InStr(rev$, "__") + 2)
          rev$ = Left$(rev$, InStr(rev$, "__") - 1)
        End If
        rwert$ = adressbeziehung(adrid$, rev$, rfeld$)
        Print #p%, Left$(l$, q% - 1);: Print #p%, rwert$;
      End If
      If ttest$ = "this" Then
        rfeld$ = Mid$(rev$, InStr(rev$, "__") + 2)
        rev$ = Left$(rev$, InStr(rev$, "__") - 1)
Debug.Print rfeld$; " - "; rev$
        ttest$ = adrid$: If kid$ <> "" And kid$ <> "-1" Then ttest$ = ttest$ + kid$
        cmd$ = "select FeldDaten as rc from auftritthigru where auftrittsid='" + ttest$ + "' and lcase(auftrittstyp)='" + LCase(rev$) + "' and lcase(FeldName)='" + LCase(rfeld$) + "'"
        rwert$ = ""
        Set hdat = New ADODB.Recordset
        hdat.CursorLocation = adUseServer
        rrr = form1.adoopen(hdat, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If rrr = 0 Then
'          rwert$ = cmd$ + " := " + trm(hdat!rc)
          rwert$ = trm(hdat!rc)
        End If
        Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(rwert$, vbCrLf, "\par ");
      End If
      If ttest$ = "user" Then
        If Not udat.EOF Then
          For i% = 0 To 21 ' see einstellungen.load
            If Len(udat.Fields(i%).name) = Len(rev$) - 1 Then   'aliase ermitteln
              If isdigit(Right$(rev$, 1)) <> 0 Then rev$ = Left$(rev$, Len(rev$) - 1)
            End If
            If LCase(udat.Fields(i%).name) = LCase(rev$) Then
              uv$ = ""
              If Not IsNull(udat.Fields(i%).value) Then uv$ = udat.Fields(i%).value
              Call tm_start(3)
              Print #p%, Left$(l$, q% - 1);:  Print #p%, strrepl(uv$, "\", "\\");
              twerg = twerg + tm_stop(3)
              i% = 33
            End If
          Next i%
          If i% < 30 Then
            erg$ = getusersetting(rev$)
            If erg$ <> "" Then
              Call tm_start(3)
              Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(erg$, "\", "\\");
              twerg = twerg + tm_stop(3)
            End If
          End If
        End If
      End If
      If ttest$ = "system" Then
        Call tm_start(3)
        Select Case LCase(rev$)
          Case "datum": Print #p%, Left$(l$, q% - 1);: Print #p%, Date;
          Case "wochentag": Print #p%, Left$(l$, q% - 1);: Print #p%, dayofweek(CDate(Date));
          Case "zeit": Print #p%, Left$(l$, q% - 1);: Print #p%, Left(Time, 5);
          Case "mwst": Print #p%, Left$(l$, q% - 1);: Print #p%, fixeurnozerotail(form1.sys_mwst / 100);
          Case "rechnr": Print #p%, new_rechnr(fn$, rechbez$);: Call dbgu("RechNr=" + getsystemsetting("RechNr", ""))
          Case Default:
        End Select
        twerg = twerg + tm_stop(3)
      End If

      ln$ = Mid$(l$, q% + 1)
      Do
        pb% = InStr(LCase(ln$), bkmend$ + LCase(orgt$))
        If pb% = 0 Then Line Input #o%, ln$
      Loop Until pb% > 0
        ln$ = Mid$(ln$, pb%)
      If InStr(ln$, "}") = 0 Then
        l$ = ""
      Else
        l$ = Mid$(ln$, InStr(ln$, "}") + 1)
      End If
    Else
      If Len(pout$) > 25000 Then
        Call tm_start(3)
        Print #p%, pout$;
        ddtt = tm_stop(3)
      twerg = twerg + ddtt
        pout$ = l$
      Else
        pout$ = pout$ + l$
      End If
      l$ = ""
    End If
    ttest$ = ""
    rev$ = ""

  Wend
Wend
If pout$ <> "" Then Print #p%, pout$;
Close #o%
Close #p%
MousePointer = 0
Call form1.dbg2f("Dokument erstellt: " + trm(tm_stop(0)) + " ms")
Call form1.dbg2f("Lesezeit: " + trm(trerg) + " ms")
Call form1.dbg2f("Schreibzeit: " + trm(twerg) + " ms")
Call form1.dbg2f("Bytes: " + trm(brerg))
If trerg > 0 Then Call form1.dbg2f("Bytes/ms lesen    : " + trm(brerg / trerg))
If twerg > 0 Then Call form1.dbg2f("Bytes/ms schreiben: " + trm(brerg / twerg))
If memono% = 1 And InStr(vorlage$, getusersetting("memovorlage", "notiz.rtf")) > 0 Then
  memono% = 0
  Exit Sub
End If

If InStr(LCase(opt1$), "noshow") = 0 Then Call openthisdoc(fn$, "")

End Sub
Public Function ExQDef(cmdstr$, whto$)

    Dim errLoop As Error, didr%, fbn$, o%, p%, xqdf$, rrr, c$, id$, fld$, w$

'd2infile = "Form1": d2insub = "ExQDef"
'  dbg cmdstr$
    ExQDef = True
    If trm(cmdstr$) = "" Then Exit Function
    If Not granted(cmdstr$) Then
      ExQDef = False
      Exit Function
    End If
    If whto$ = "adoc" And fallbackdir$ <> "" Then
      didr% = 1000
      Do
        fbn$ = fallbackdir$ & "\" & datum2sql(Date) & "-" & strrepl(Time$, ":", "-") & "-" & trm(didr%) & "-" & uId$
        didr% = didr% + 1
      Loop Until exist(fbn$ & ".sql") = 0 Or didr% > 9990
'      o% = FreeFile: Open fbn$ & ".sql.lock" For Output As #o%: Close #o%
      o% = FreeFile
      Open fbn$ & ".sql" For Output As #o%
      Print #o%, cmdstr$
      Close #o%
'      On Error Resume Next
'      Kill fbn$ & ".sql.lock"
'      On Error GoTo 0
    End If
    If whto$ = "adoc" And fallbackserverpath$ <> "" Then
        didr% = 1000
        Do
          fbn$ = fallbackserverpath$ & "\" & datum2sql(Date) & "-" & strrepl(Time$, ":", "-") & "-" & trm(didr%) & "-" & uId$
          didr% = didr% + 1
        Loop Until exist(fbn$ & ".sql") = 0 Or didr% > 9990
        o% = FreeFile
        Open fbn$ & ".sql" For Output As #o%
'?? besser??
'    cmdstr$ = strrepl(cmdstr$, "\", "|backslashbackslash|")
'    cmdstr$ = strrepl(cmdstr$, "|backslashbackslash|", "\\")
        xqdf$ = cmdstr$
        If Right$(xqdf$, 1) <> ";" Then xqdf$ = xqdf$ & ";"
        Print #o%, xqdf$
        Close #o%
    End If

  didr% = 0
  If backslashhandler = "an" Then
    cmdstr$ = strrepl(cmdstr$, "\", "\\")
  End If
  If whto$ = "adoc" Then
    On Error GoTo Err_Execute
    adoc.Execute cmdstr$
    On Error GoTo 0
  Else
      On Error GoTo Err_Execute
      alertdbo.Execute cmdstr$
      On Error GoTo 0
  End If
  If didr% = 0 And Not noalarms Then Call chkalarmlist(cmdstr$)
  Exit Function

Err_Execute:

  rrr = Err
  didr% = 1
  ExQDef = False
  Call errhdl("Fehlernummer: " & rrr & vbCr & Error$(rrr) & "statement=" & cmdstr$)
  Resume Next

End Function

Private Function errfilter(stmt$) As Boolean
errfilter = False
If InStr(LCase(stmt$), "insert into mailsafe ") > 0 And InStr(LCase(stmt$), "duplicate entry") > 0 Then
  errfilter = True
  Exit Function
End If
'Debug.Print stmt$
End Function

Public Function getfromtplan(tpid$, fld$) As String
Dim rtmp As ADODB.Recordset, rcd$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getfromtplan"
getfromtplan = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT " + fld$ + " as rc FROM tplan where id='" + tpid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
rcd$ = trm(rtmp!rc)
getfromtplan = rcd$

End Function

Public Function getkompnamebyid(kid$) As String
Dim n$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompnamebyid"
getkompnamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name,vornamen FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
'If IsNull(rtmp!Name) Or IsNull(rtmp!vornamen) Then Exit Function
getkompnamebyid = trm(rtmp!name & ", " + rtmp!vornamen)

End Function

Public Function getkompnachnamebywerkid(wid$) As String
Dim n$, kid$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompnachnamebywerkid"
getkompnachnamebywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT KomponistenNummer FROM w_loc where id='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!KomponistenNummer) Then Exit Function
kid$ = rtmp!KomponistenNummer
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name,vornamen FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!name) Or IsNull(rtmp!vornamen) Then Exit Function
getkompnachnamebywerkid = trm(rtmp!name)

End Function

Public Function getkompnrbyid(kid$) As String
Dim n$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompnrbyid"
getkompnrbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT kompnr FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
'If IsNull(rtmp!Name) Or IsNull(rtmp!vornamen) Then Exit Function
getkompnrbyid = trm(rtmp!kompnr)

End Function
Public Function getwerke4tt(id$) As String
Dim rtmp As ADODB.Recordset, wid$, sid$, g$, rrr, ll%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getwerke4tt"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

g$ = "  "
While Not rtmp.EOF
  wid$ = rtmp!werkid
  If Left$(wid$, 4) = "SBZ:" Then
    sid$ = Mid$(wid$, 5)
    wid$ = form1.getsatzidbywerkid(sid$)
  End If
  If sid$ = "" Then
    g$ = g$ + "  " + Left(getkompnamebywerkid(wid$), 6) + ":" + getwerknamebyid(wid$) + Chr$(13) + Chr$(10)
  Else
    g$ = g$ + "  " + Left(getkompnamebywerkid(wid$), 6) + ":" + getsatznamebyid(sid$) + " " + transe("aus") + " " + getwerknamebyid(wid$) + Chr$(13) + Chr$(10)
  End If
  rtmp.MoveNext
Wend
getwerke4tt = g$
End Function

Public Function getwerke(id$) As String
Dim rtmp As ADODB.Recordset, wid$, sid$, g$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getwerke"
getwerke = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Function
g$ = ""
While Not rtmp.EOF
  wid$ = rtmp!werkid
  If Left$(wid$, 4) = "SBZ:" Then
    sid$ = Mid$(wid$, 5)
    wid$ = form1.getsatzidbywerkid(sid$)
  End If
  If sid$ = "" Then
    g$ = g$ + getkompnamebywerkid(wid$) + ":" + getwerknamebyid(wid$) + Chr$(13) + Chr$(10)
  Else
    g$ = g$ + getkompnamebywerkid(wid$) + ":" + getsatznamebyid(sid$) + " " + transe("aus") + " " + getwerknamebyid(wid$) + Chr$(13) + Chr$(10)
  End If
  rtmp.MoveNext
Wend
getwerke = g$
End Function

Public Function getwerkids(id$) As String
Dim rtmp As ADODB.Recordset, wid$, sid$, g$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getwerke"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

g$ = ""
While Not rtmp.EOF
  wid$ = trm(rtmp!werkid)
  If Left$(wid$, 4) = "SBZ:" Then
    sid$ = Mid$(wid$, 5)
    wid$ = form1.getsatzidbywerkid(sid$)
  End If
  If g$ <> "" Then g$ = g$ + "|"
  g$ = g$ + wid$
  rtmp.MoveNext
Wend
getwerkids = g$
End Function

Public Function getkompnamebywerkid(wid$) As String
Dim n$, kid$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompnamebywerkid"
getkompnamebywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT KomponistenNummer FROM w_loc where id='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!KomponistenNummer) Then Exit Function
kid$ = rtmp!KomponistenNummer
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name,vornamen FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
'If IsNull(rtmp!Name) Or IsNull(rtmp!vornamen) Then Exit Function
getkompnamebywerkid = "" & rtmp!name & ", " & rtmp!vornamen

End Function

Public Function wrk_arrandedby(wid$) As String
Dim n$, kid$, rrr, rc$
Dim rtmp As ADODB.Recordset
Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "wrk_arrandedby"

wrk_arrandedby = ""
If form1.isfieldmissing("opt_arranged", "id") Then Exit Function

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT aid,wid from opt_arranged where wid='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rrr <> 0 Then
  wrk_arrandedby = "Error: (" + trm(rrr) + ")" + Error$(rrr)
  Exit Function
End If
rc$ = ""
While Not rtmp.EOF
  If rc$ <> "" Then rc$ = rc$ + "|"
  rc$ = rc$ + strrepl(form1.getnamebyid(rtmp!aid), vbCrLf, " - ")
  rtmp.MoveNext
Wend

wrk_arrandedby = rc$

End Function

Public Function wrk_publishedby(wid$) As String
Dim n$, kid$, rrr
Dim rtmp As ADODB.Recordset
Dim d2infile As String, d2insub As String, rc$
d2infile = "Form1": d2insub = "wrk_publishedby"

wrk_publishedby = ""
If form1.isfieldmissing("opt_arranged", "id") Then Exit Function

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT aid,wid from opt_published where wid='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rrr <> 0 Then
  wrk_publishedby = "Error: (" + trm(rrr) + ")" + Error$(rrr)
  Exit Function
End If
rc$ = ""
While Not rtmp.EOF
  If rc$ <> "" Then rc$ = rc$ + "|"
  rc$ = rc$ + strrepl(form1.getnamebyid(rtmp!aid), vbCrLf, " - ")
  rtmp.MoveNext
Wend
wrk_publishedby = rc$

End Function

Public Function getkompvornamenamebywerkid(wid$) As String
Dim n$, kid$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompvornamenamebywerkid"
getkompvornamenamebywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT KomponistenNummer FROM w_loc where id='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!KomponistenNummer) Then Exit Function
kid$ = rtmp!KomponistenNummer
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name,vornamen FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If Not IsNull(rtmp!name) Then getkompvornamenamebywerkid = rtmp!name
If IsNull(rtmp!name) Or IsNull(rtmp!vornamen) Then Exit Function
getkompvornamenamebywerkid = rtmp!vornamen + " " + rtmp!name

End Function
Public Function getkompdatesbywerkid(wid$) As String
Dim n$, kid$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompdatesbywerkid"
getkompdatesbywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT KomponistenNummer FROM w_loc where id='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!KomponistenNummer) Then Exit Function
kid$ = rtmp!KomponistenNummer
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT daten FROM k_loc where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!daten) Then Exit Function
getkompdatesbywerkid = rtmp!daten

End Function
Public Function getwerknamebyid(id$) As String
Dim n$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getwerknamebyid"
getwerknamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name FROM w_loc where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!name) Then Exit Function
getwerknamebyid = rtmp!name

End Function
Public Function getwerkopusnamebyid(id$) As String
Dim n$, rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getwerkopusnamebyid"
getwerkopusnamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name,opusname1 FROM w_loc where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
n$ = trm(rtmp!name)
If n$ = "" Then n$ = trm(rtmp!Opusname1)
getwerkopusnamebyid = n$

End Function
Public Function getdauerbywerkid(id$) As String
Dim rrr
Dim n$, p%
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getdauerbywerkid"
getdauerbywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT dauer FROM w_loc where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!Dauer) Then Exit Function
getdauerbywerkid = rtmp!Dauer

End Function

Public Function getgemanrbywerkid(id$) As String
Dim rrr
Dim n$, p%
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getdauerbywerkid"
getgemanrbywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT s14 FROM w_loc where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!s14) Then Exit Function
getgemanrbywerkid = rtmp!s14

End Function
Public Function getalbumbywerkid(id$) As String
Dim rrr
Dim n$, p%
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getalbumbywerkid"
getalbumbywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT s13 FROM w_loc where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!s13) Then Exit Function
getalbumbywerkid = trm(rtmp!s13)

End Function

Public Sub apmaillog(aid$, kid$, anadr$, sbj$)
Dim cmd$, l$, tel$
Dim o%, rrr, knam$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "apmaillog"

knam$ = "": tel$ = ""
If kid$ <> "-1" Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    knam$ = trm(rtmp!name)
    tel$ = trm(rtmp!tel)
  End If
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adresse where id ='" + aid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  If tel$ = "" Then tel$ = trm(rtmp!tel)
  l$ = """" + trm(rtmp!id) + """" + excelfeldtrenner()
  If knam$ <> "" Then l$ = l$ + """" + knam$ + """"
  l$ = l$ + excelfeldtrenner()
  l$ = l$ + """" + anadr$ + """" + excelfeldtrenner()
  l$ = l$ + """" + tel$ + """" + excelfeldtrenner()
  l$ = l$ + """" + sbj$ + """"
  o% = FreeFile
  On Error Resume Next
  Open s0d$ & "\sentmaillog_" & uId$ & ".csv" For Append As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    Exit Sub
  End If
  Print #o%, """" + datum2sql(Date) + " " + trm(Time) + """" + excelfeldtrenner() + l$
  Close #o%
End If

End Sub

Public Sub dbg2f(l$, Optional infile$, Optional insub$)
Dim o%, d0 As Double, diffd As Double, rrr

If dbg2file% = 0 Then Exit Sub
d0 = Date + Time: diffd = 0
o% = FreeFile
On Error Resume Next
Open s0d$ & "\debug2file_" & uId$ & ".txt" For Append As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Exit Sub
End If
If diffd > 0 Then Print #o%, Time & " "; "Anfrage hat" & str$(diffd); "s auf Freigabe gewartet"
Print #o%, Time & ":" + infile$ + ":" + insub$ + ":" & l$
Debug.Print Time & ":" + infile$ + ":" + insub$ + ":" & l$
Close #o%

End Sub
Public Sub log2f(l$, Optional infile$, Optional insub$)
Dim o%, rrr

o% = FreeFile
On Error Resume Next
MkDir s0d$ & "\" + "extralog"
rrr = Err
On Error GoTo 0
On Error Resume Next
Open s0d$ & "\extralog\logfile_" & uId$ & datum2sql(Date) & ".txt" For Append As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Exit Sub
End If
Print #o%, Time & ":" + infile$ + ":" + insub$ + ":" & l$
Close #o%

End Sub
Public Sub dbg(l$)
Dim o%, rrr

Debug.Print l$
If Len(uclog$) = 0 Then Exit Sub

o% = FreeFile
On Error Resume Next
Open uclog$ For Append As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Eine Log-Datei (" & uclog$ & ") konnte nicht geöffnet werden. Es wird nicht mehr protokolliert"
  uclog$ = ""
  Exit Sub
End If
'eins zu viel macht nix, eins zu wenig aber ...
If Right$(l$, 1) <> ";" Then l$ = l$ + ";"
Print #o%, l$
Close #o%

End Sub
Public Sub errhdl(l$)
Dim o%, i%, fn$, rrr

'd2infile = "Form1": d2insub = "errhdl"
If InStr(LCase(l$), "update finanzen ") > 0 Then Exit Sub
If errfilter(l$) Then
  Exit Sub
End If
Debug.Print l$

o% = FreeFile
i% = 0
If ehsc% > i% Then i% = ehsc%
Do
  fn$ = s0d$ & "\" + docs() + "\" & uId$ & trm(str$(i%)) & ".err"
  i% = i% + 1
Loop Until exist(fn$) = 0 Or i% > 1000
ehsc% = i%
Call dbg2f("Fehler: " & l$)
If i% < 1000 Then
  On Error Resume Next
  Open fn$ For Output As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    MsgBox transe("Fehler:") + Chr$(13) + Chr$(10) + l$
    Exit Sub
  End If
  Print #o%, trm(Date) + " " + trm(Time) + ": " + l$
  Close #o%
End If

End Sub

Public Function systemanrede(kid$) As String
Dim rrr
Dim rtmp As ADODB.Recordset, cmd$


Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "systemanrede"
systemanrede = ""

  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT an FROM anreden where kid='" + kid$ + "' AND user='system'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rtmp.EOF Then Exit Function
  If Not IsNull(rtmp!an) Then systemanrede = kommasettings(rtmp!an, "an")
  rtmp.Close

End Function

Public Function systemabrede(kid$) As String
Dim rrr
Dim rtmp As ADODB.Recordset, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "systemabrede"
systemabrede = ""

  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT ab FROM anreden where kid='" + kid$ + "' AND user='system'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rtmp.EOF Then Exit Function
  If trm(rtmp!Ab) <> "" Then systemabrede = kommasettings(rtmp!Ab, "ab")
  rtmp.Close

End Function

Public Function meineanrede(kid$) As String
Dim rrr
Dim rtmp As ADODB.Recordset, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "meineanrede"
meineanrede = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
cmd$ = "SELECT an FROM anreden where kid='" + kid$ + "' AND user='" + anredeuser$ + "'"
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then GoTo ab2an
If Not IsNull(rtmp!an) Then
  meineanrede = kommasettings(rtmp!an, "an")
Else
ab2an:
  meineanrede = systemanrede(kid$)
End If
rtmp.Close

End Function

Public Function meineabrede(kid$) As String
Dim rrr
Dim rtmp As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "meineabrede"
meineabrede = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT ab FROM anreden where kid='" + kid$ + "' AND user='" + anredeuser$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then GoTo ab2ab
If Not IsNull(rtmp!Ab) Then
  meineabrede = kommasettings(rtmp!Ab, "ab")
Else
ab2ab:
  meineabrede = systemabrede(kid$)
End If
rtmp.Close

End Function

Public Function meinefaxvorlage() As String
Dim rrr
Dim rtmp As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "meinefaxvorlage"
meinefaxvorlage = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT faxvorlage FROM benutzerdaten where id='" + uId$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If Not IsNull(rtmp!faxvorlage) Then meinefaxvorlage = rtmp!faxvorlage
rtmp.Close

End Function
Public Function meinetalistevorlage() As String
'd2infile = "Form1": d2insub = "meinetalistevorlage"
meinetalistevorlage = "taliste.rtf"
End Function
Public Function meineprgdruckvorlage() As String
'd2infile = "Form1": d2insub = "meineprgdruckvorlage"
meineprgdruckvorlage = getusersetting("programmvorlage", "prgdruck.rtf")
End Function
Public Function meinebriefvorlage() As String
Dim rrr
Dim rtmp As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "meinebriefvorlage"
meinebriefvorlage = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT briefvorlage FROM benutzerdaten where id='" + uId$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If Not IsNull(rtmp!briefvorlage) Then meinebriefvorlage = rtmp!briefvorlage
rtmp.Close

End Function
Public Function meinememovorlage() As String
'd2infile = "Form1": d2insub = "meinememovorlage"
meinememovorlage = getusersetting("memovorlage", "notiz.rtf")

End Function

Public Function immerspeichern() As String
'd2infile = "Form1": d2insub = "immerspeichern"
immerspeichern = usavealways$
End Function

Public Function immerkalender() As String
'd2infile = "Form1": d2insub = "immerkalender"
immerkalender = ucalalways$
End Function

Public Function myuniquedocname(para$, Optional extension$) As String
Dim o%, docexists As Boolean, doctest As Boolean, ext$, fn$, rrr, i%, pfn$, wert$, rcnt%, ask

'd2infile = "Form1": d2insub = "myuniquedocname"
myuniquedocname = ""
ext$ = ".rtf"
doctest = False
If getusersetting("convert2doc", "nein") = "ja" Then doctest = True
If trm(extension$) <> "" Then ext$ = extension$
If Left$(ext$, 1) <> "." Then ext$ = "." & ext$
o% = FreeFile
fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  On Error Resume Next
  MkDir s0d$ & "\" + docs()
  MkDir s0d$ & "\" + docs() + "\" & uId$
  On Error GoTo 0
Else
  Close #o%
  Kill fn$
End If
i% = 0
If udochiscore% > 0 Then i% = udochiscore%
If Left(para$, 3) = "FN:" Then
  pfn$ = Mid$(para$, 4)
  Do
    docexists = False
    i% = i% + 1
    If doctest Then
      fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & pfn$ & "-" & uId$ & "-" & trm(str$(i%)) & ".doc"
      docexists = Not nexist(fn$)
    End If
    fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & pfn$ & "-" & uId$ & "-" & trm(str$(i%)) & ext$
  Loop Until exist(fn$) = 0
Else
  Do
    docexists = False
    i% = i% + 1
    If doctest Then
      fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & uId$ & datum2sql(Date) & trm(i%) & ".doc"
      docexists = Not nexist(fn$)
    End If
    fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & uId$ & datum2sql(Date) & trm(i%) & ext$
  Loop Until (exist(fn$) = 0 And docexists = False) Or i > 20000
End If
udochiscore% = i%
If para$ <> "noask" Then
  wert$ = ""
  wert$ = form1.saveasBox(fn$)
  rcnt% = 0
  If trm(wert$) <> "" Then
    Call form1.dbg2f("myuniquedocname: got " + wert$)
    If LCase(Right$(wert$, 4)) <> ext$ Then wert$ = wert$ & ext$
    If DirName(wert$) = "" Then wert$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & wert$
    If exist(wert$) <> 0 And Right$(wert$, 8) = "temp.rtf" Then
kretry:
      Call form1.dbg2f("myuniquedocname: trying to delete " + wert$)
      On Error Resume Next
      Kill wert$
      rrr = Err
      On Error GoTo 0
      Call form1.dbg2f("myuniquedocname: delete result: " + trm(rrr) + " " + Error$(rrr))
      If rrr = 70 Then
        rcnt% = rcnt% + 1
        wert$ = DirName(wert$) & "\temp" & trm(rcnt%) & ".rtf"
        If exist(wert$) <> 0 Then
          If rcnt% < 10 Then GoTo kretry
          MsgBox "Es existieren zu viele nicht löschbare temporäre Dateien." & vbCrLf & "Bitte prüfen Sie das."
          End
        End If
      Else
        Call form1.dbg2f("myuniquedocname: Error must be 70 to be handled")
      End If
      If getusersetting("convert2doc", "nein") <> "nein" Then
        Call form1.dbg2f("myuniquedocname (convert2doc): trying to delete " + DirName(wert$) & basename(FileName(wert$), ".rtf") & ".doc")
        On Error Resume Next
        Kill DirName(wert$) & basename(FileName(wert$), ".rtf") & ".doc"
        On Error GoTo 0
      End If
    End If
    Call form1.dbg2f("myuniquedocname: current filename: " + wert$)
    If Right$(wert$, 8) = "temp.rtf" Then
      If getusersetting("convert2doc", "nein") <> "nein" Then
        On Error Resume Next
        Kill DirName(wert$) & basename(FileName(wert$), ".rtf") & ".doc"
        On Error GoTo 0
      End If
    End If
    If exist(wert$) <> 0 Then
      ask = MsgBox("Die Datei " + wert$ + " existiert bereits. Überschreiben?", vbYesNo + vbCritical + vbDefaultButton2, "Vorhandene Datei löschen?")
      If ask = vbNo Then
        wert$ = ""
      Else
        On Error Resume Next
        Kill wert$
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then wert$ = ""
      End If
    End If
  End If
  fn$ = ""
  If wert$ <> "" Then fn$ = wert$
End If
Call form1.dbg2f("myuniquedocname: returning filename: " + fn$)
myuniquedocname = fn$

End Function
Public Function myuniquedocnameinpath(para$, opt1$) As String
Dim o%, localpara$, pfn$, bfn$, fn$, rrr, i%, wert$, ask As Integer

'd2infile = "Form1": d2insub = "myuniquedocnameinpath"
localpara$ = para$
myuniquedocnameinpath = ""
pfn$ = ""
If Left(localpara$, 3) = "FN:" Then
  pfn$ = Mid$(localpara$, 4)
  localpara$ = DirName(pfn$)
  pfn$ = FileName(pfn$)
  bfn$ = basename(pfn$, ".rtf")
End If
o% = FreeFile
fn$ = localpara$ + "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Call makedir(localpara$)
Else
  Close #o%
  Kill fn$
End If
i% = 0
If udochiscore1% > 0 Then i% = udochiscore1%
If pfn$ = "" Then
  Do
    i% = i% + 1
    fn$ = localpara$ & "\" & uId$ & Right$(Date, 2) + Left$(Date, 2) & Mid$(Date, 4, 2) + "-" + trm(str$(i%)) & ".rtf"
  Loop Until exist(fn$) = 0
Else
  Do
    i% = i% + 1
    fn$ = localpara$ & "\" + bfn$ + "-" + trm(str$(i%)) + ".rtf"
  Loop Until exist(fn$) = 0
End If
udochiscore1% = i%
If LCase(opt1$) <> "noask" Then
  wert$ = ""
  wert$ = saveasBox(fn$)
  If trm(wert$) <> "" Then
    If LCase(Right$(wert$, 4)) <> ".rtf" Then wert$ = wert$ + ".rtf"
    If DirName(wert$) = "" Then wert$ = localpara$ & "\" & wert$
    If exist(wert$) <> 0 Then
      If InStr(wert, "\temp.rtf") = 0 Then
        ask = MsgBox("Die Datei " + wert$ + " existiert bereits. Überschreiben?", vbYesNo + vbCritical + vbDefaultButton2, "Vorhandene Datei löschen?")
      Else
        ask = vbYes
      End If
      If ask = vbNo Then
        wert$ = ""
      Else
        Kill wert$
      End If
    End If
  End If
  fn$ = ""
  If wert$ <> "" Then fn$ = wert$
End If
myuniquedocnameinpath = fn$

End Function
Public Function myuniquebmpnameinpath(para$) As String
Dim o%, fn$, rrr, i%, wert$, ask As Integer

'd2infile = "Form1": d2insub = "myuniquebmpnameinpath"
myuniquebmpnameinpath = ""
o% = FreeFile
fn$ = para$ + "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  On Error Resume Next
  MkDir para$
  On Error GoTo 0
Else
  Close #o%
  Kill fn$
End If
i% = 0
Do
  i% = i% + 1
  fn$ = para$ & "\" & uId$ & Left$(Date, 2) & Mid$(Date, 4, 2) & trm(str$(i%)) & ".bmp"
Loop Until exist(fn$) = 0


  wert$ = ""
  wert$ = saveasBox(fn$)
  If trm(wert$) <> "" Then
    If LCase(Right$(wert$, 4)) <> ".bmp" Then wert$ = wert$ + ".bmp"
    If DirName(wert$) = "" Then wert$ = para$ & "\" & wert$
    If exist(wert$) <> 0 Then
      ask = MsgBox("Die Datei " + wert$ + " existiert bereits. Überschreiben?", vbYesNo + vbCritical + vbDefaultButton2, "Vorhandene Datei löschen?")
      If ask = vbNo Then
        wert$ = ""
      Else
        Kill wert$
      End If
    End If
  End If
  fn$ = ""
  If wert$ <> "" Then fn$ = wert$

myuniquebmpnameinpath = fn$

End Function
Public Function mydatadir() As String

'd2infile = "Form1": d2insub = "mydatadir"
mydatadir = s0d$ & "\" + docs() + "\" & uId$

On Error Resume Next
MkDir s0d$ & "\" + docs()
MkDir mydatadir
On Error GoTo 0

End Function
Public Function mylocaldatadir() As String

'd2infile = "Form1": d2insub = "mydatadir"
mylocaldatadir = localdir & "\" + docs() + "\" & uId$

On Error Resume Next
MkDir localdir & "\" + docs()
MkDir mylocaldatadir
On Error GoTo 0

End Function
Public Function mytmpdir() As String

'd2infile = "Form1": d2insub = "mytmpdir"
mytmpdir = s0d$ & "\" + docs() + "\" & uId$

On Error Resume Next
MkDir mydatadir()
MkDir mytmpdir
On Error GoTo 0

End Function

Public Function docs() As String

'd2infile = "Form1": d2insub = "docs"
docs = doc0dir$

End Function
Public Function medien() As String

'd2infile = "Form1": d2insub = "medien"
medien = m0dir$

End Function
Public Function myfirstdayofweek() As String

'd2infile = "Form1": d2insub = "myfirstdayofweek"
myfirstdayofweek = ufdow$

End Function
Public Function mymailoutserver() As String

'd2infile = "Form1": d2insub = "mymailoutserver"
mymailoutserver = umsrvout$

End Function
Public Function mymailaddress() As String

'd2infile = "Form1": d2insub = "mymailaddress"
mymailaddress = umailadr$

End Function

Public Function sqlqry(stmt$)

If dbg2file% <> 0 Then Call dbg2f(stmt$)
'If InStr(stmt$, " finanzen ") > 0 Then
'  Debug.Print stmt$
'End If
sqlqry = form1.ExQDef(stmt$, "adoc")

End Function

Public Sub alrtdbsqlqry(stmt$)

'd2infile = "Form1": d2insub = "wawisqlqry"
If dbg2file% <> 0 Then Call dbg2f(stmt$)
Call form1.ExQDef(stmt$, "alrtdbc")

End Sub

Public Sub openthisdoc(fn$, Options As String)
Dim xx$, fup$, nfn$, sky$, c$, rrr, X, o%

'd2infile = "Form1": d2insub = "openthisdoc"
rrr = 0
xx$ = trm(form1.getmyeditor(FileExtension(fn$)))
If xx$ = "" Or xx$ = "write.exe" Then rrr = 1
If rrr = 0 Then
Call form1.dbg2f("openthisdoc (" + xx$ + "): " + fn$, "", "")
  fup$ = Chr$(34)
  If getusersetting("KurzeDateinamen", "nein") = "ja" Then fup$ = ""
  If InStr(xx$, "start ") = 1 Then
    o% = FreeFile
    Open "c:\Agencyprof\wstart.bat" For Output As #o%
    Print #o%, xx$ & " " & form1.fixfilename(fn$)
    Print #o%, "exit"
    Close #o%
    On Error Resume Next
    Call form1.dbg2f("Shell(cmd.exe): " + Chr$(34) + form1.fixfilename(xx$) + Chr$(34) & " " & fup$ & form1.fixfilename(fn$) & fup$)
    X = Shell("c:\Agencyprof\wstart.bat", vbMinimizedNoFocus)
    rrr = Err
    On Error GoTo 0
    DoEvents
  Else
    On Error Resume Next
    Call form1.dbg2f("Shell: " + Chr$(34) + form1.fixfilename(xx$) + Chr$(34) & " " & fup$ & form1.fixfilename(fn$) & fup$)
    X = Shell(Chr$(34) + form1.fixfilename(xx$) + Chr$(34) & " " & fup$ & form1.fixfilename(fn$) & fup$, 1)
    rrr = Err
    On Error GoTo 0
    DoEvents
    If rrr = 0 And LCase(FileExtension(fn$)) = "rtf" Then
      sky$ = getusersetting("convert2doc", "nein")
      If sky$ <> "nein" And InStr(Options, "noconvert") = 0 Then
        If LCase(Right$(fn$, 4)) = ".rtf" Then
          nfn$ = Left(fn$, Len(fn$) - 4) + ".doc"
          wait 1
          AppActivate "Microsoft Word"
          SendKys sky$, 1
          wait 2
          If Not nexist(nfn$) Then
            Me.BackColor = convertcolor
            c$ = "update dochist set docname='" + nfn$ + "' where docname='" + fn$ + "';"
            Call sqlqry(c$)
            deldoclist.AddItem fn$
          End If
        End If
      End If
    End If
  End If
End If
If rrr <> 0 Then
  Load rtfview
  rtfview.loadtext (fn$)
End If

End Sub
Public Function getAdrProperty(kid$, prop$) As String
Dim rrr
Dim rtmp As ADODB.Recordset
Dim sidp%, sida$, sidk$, sid$, t$, f$, e$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getAdrProperty"
getAdrProperty = ""
t$ = "": f$ = "": e$ = "": sid$ = kid$: sidk$ = "-1": sidp% = InStr(sid$, "{"): sida$ = sid$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
If sidp% > 0 Then
  sidk$ = trm(Left(sid$, sidp% - 1))
  sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  rrr = form1.adoopen(rtmp, "SELECT " + prop$ + " as prop FROM kontakt where id='" + sidk$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    t$ = trm(rtmp!prop)
  End If
Else
  rrr = form1.adoopen(rtmp, "SELECT " + prop$ + " as prop FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    t$ = trm(rtmp!prop)
  End If
End If
getAdrProperty = t$
End Function

Public Function gettelfaxmail(kid$) As String
Dim rrr
Dim rtmp As ADODB.Recordset
Dim sidp%, sida$, sidk$, sid$, t$, f$, e$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "gettelfaxmail"
gettelfaxmail = ""
t$ = "": f$ = "": e$ = ""
sid$ = kid$: sidk$ = "-1"
sidp% = InStr(sid$, "{")
sida$ = sid$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
If sidp% > 0 Then
  sidk$ = trm(Left(sid$, sidp% - 1))
  sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  rrr = form1.adoopen(rtmp, "SELECT tel,fax,email FROM kontakt where id='" + sidk$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    t$ = trm(rtmp!tel)
    f$ = trm(rtmp!fax)
    e$ = trm(rtmp!email)
  End If
Else
  rrr = form1.adoopen(rtmp, "SELECT tel,fax,email FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Function
  If Not rtmp.EOF Then
    If t$ = "" Then t$ = trm(rtmp!tel)
    If f$ = "" Then f$ = trm(rtmp!fax)
    If e$ = "" Then e$ = trm(rtmp!email)
  End If
End If
gettelfaxmail = sida$ + Chr$(13) & Chr$(10) & "Tel.: " & t$ & Chr$(13) & Chr$(10) & "Fax: " & f$ & Chr$(13) & Chr$(10) & "eMail: " & e$
End Function

Public Function get_kontaktname_by_id(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_kontaktname_by_id"
get_kontaktname_by_id = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
n$ = strrepl(kid$, "'", "")
rrr = form1.adoopen(rtmp, "SELECT name FROM kontakt where id='" + n$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!name) Then Exit Function
get_kontaktname_by_id = rtmp!name

End Function

Public Function get_defaultchecklisttext(chkid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_defaultchecklisttext"
get_defaultchecklisttext = ""
If chkid$ <> "" Then

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT checkpoint FROM opt_checklists where id='" + chkid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!checkpoint) Then Exit Function
get_defaultchecklisttext = rtmp!checkpoint


End If
End Function

Public Function getadridbykontaktid(kid As String) As String
Dim rrr, sidp%
Dim n$, sid$, sidk$, sida$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getadridbykontaktid"
getadridbykontaktid = ""

If InStr(kid, "{") > 0 Then
    sid$ = kid
    sidp% = InStr(sid$, "{")
    sida$ = sid$
    sidk$ = trm(Left(sid$, sidp% - 1))
    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    kid$ = form1.get_kontaktid_by_name(sida$, sidk$)
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT vid FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!vid) Then Exit Function
getadridbykontaktid = rtmp!vid

End Function
Public Function getsatzidbywerkid(sid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getsatzidbywerkid"
getsatzidbywerkid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT wid FROM sbz_loc where id='" + sid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!wid) Then Exit Function
getsatzidbywerkid = rtmp!wid

End Function
Public Function getsatznamebyid(sid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getsatznamebyid"
getsatznamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT satzbezeichnung FROM sbz_loc where id='" + sid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!satzbezeichnung) Then Exit Function
getsatznamebyid = rtmp!satzbezeichnung

End Function
Public Function getkontaktemailbyid(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkontaktemailbyid"
getkontaktemailbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT email FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!email) Then Exit Function
getkontaktemailbyid = rtmp!email

End Function

Public Function getkontaktpositionbyid(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkontaktpositionbyid"
getkontaktpositionbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT position FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!Position) Then Exit Function
getkontaktpositionbyid = trm(rtmp!Position)

End Function

Public Function getkontaktabteilungbyid(kid$) As String
Dim rrr
Dim n$, vid$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkontaktabteilungbyid"
getkontaktabteilungbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT vid FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Function
vid$ = trmx1(rtmp!vid)
rtmp.Close
If vid$ = "" Then Exit Function
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT felddaten FROM auftritthigru where auftrittsid='" & vid$ & kid$ & "' and feldname='Abteilung' and auftrittstyp='Person'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Function
getkontaktabteilungbyid = trmx1(rtmp!felddaten)
rtmp.Close

End Function

Public Function getemailbyid(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getemailbyid"
getemailbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT email FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!email) Then Exit Function
getemailbyid = rtmp!email

End Function
Public Function getnamebyid(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getnamebyid"
getnamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!name) Then Exit Function
getnamebyid = rtmp!name

End Function

Public Function get1erg(c$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

get1erg = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly)
If rrr <> 0 Then
  Exit Function
End If
If rtmp.EOF Then Exit Function
On Error Resume Next
If IsNull(rtmp!wert) Then Exit Function
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Function
get1erg = rtmp!wert

End Function

Public Function get1hordeerg(c$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

get1hordeerg = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, clddb, adOpenDynamic, adLockReadOnly)
If rrr <> 0 Then
  Exit Function
End If
If rtmp.EOF Then Exit Function
On Error Resume Next
If IsNull(rtmp!wert) Then Exit Function
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Function
get1hordeerg = rtmp!wert

End Function

Public Function getidbyid(kid$) As String
'used to check existence
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getidbyid"
getidbyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!id) Then Exit Function
getidbyid = rtmp!id

End Function


Public Function getidbyname(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getidbyname"
getidbyname = kid$
If kid$ = "" Then Exit Function
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresse where name='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!id) Then Exit Function
getidbyname = rtmp!id

End Function

Function typesof(vid) As String
Dim rrr, typ$, typwert$
Dim n$, sid$, sidk$, sida$, sidp%
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "": d2insub = ""
typesof = ""

If trm(vid) <> "" Then
  If InStr(vid, "{") > 0 Then
    sid$ = vid
    sidp% = InStr(sid$, "{")
    sida$ = sid$
    sidk$ = trm(Left(sid$, sidp% - 1))
    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  Else
    sida$ = vid
    sidk$ = "-1"
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT typ FROM adresstyp where lcase(vid)='" & sida$ & "' and kid='" & sidk$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    While Not rtmp.EOF
      typesof = typesof + "|" + trm(rtmp!typ) + "|"
      rtmp.MoveNext
    Wend
  End If
End If

End Function

Function isoftype(vid, typist$) As String
Dim rrr, typ$, typwert$
Dim n$, sid$, sidk$, sida$, sidp%
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "isoftype"
isoftype = "-1"
typ$ = cut_d1(typist$, "|")
typwert$ = cut_d2bis(typist$, "|")

If trm(vid) <> "" Then
  If InStr(vid, "{") > 0 Then
    sid$ = vid
    sidp% = InStr(sid$, "{")
    sida$ = sid$
    sidk$ = trm(Left(sid$, sidp% - 1))
    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  Else
    sida$ = vid
    sidk$ = "-1"
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT id,wert FROM adresstyp where lcase(vid)='" & sida$ & "' and typ='" & typ$ & "' and kid='" & sidk$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rtmp.EOF Then Exit Function
    If typwert$ <> "" Then
      If InStr(LCase(trm(rtmp!wert)), typwert$) > 0 Then isoftype = trm(rtmp!wert)
    Else
      isoftype = trm(rtmp!wert)
    End If
End If

End Function

Function AdrIDErmittlung(vid) As String
Dim sid$, sidp%, sida$, sidk$

'd2infile = "Form1": d2insub = "AdrIDErmittlung"
AdrIDErmittlung = ""

If Not IsNull(vid) Then
  If InStr(vid, "{") > 0 Then
    sid$ = vid
    sidp% = InStr(sid$, "{")
    sida$ = sid$
    sidk$ = trm(Left(sid$, sidp% - 1))
    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
  Else
    sida$ = vid
    sidk$ = ""
  End If
  AdrIDErmittlung = sida$ & sidk$
End If

End Function
Function kisoftype(vid, typist$) As String
Dim rrr, typ$, typwert$
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "kisoftype"
kisoftype = "-1"
typ$ = cut_d1(typist$, "|")
typwert$ = cut_d2bis(typist$, "|")
If Not IsNull(vid) Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
'woher kommt das instr('" + vid + "',vid)=1 and ?
'  n$ = "SELECT id,wert FROM adresstyp where instr('" + vid + "',vid)=1 and instr('" + vid + "',kid)>0 and typ='" + typ$ + "'"
  n$ = "SELECT id,wert FROM adresstyp where instr('" + vid + "',kid)>0 and typ='" + typ$ + "'"
  rrr = form1.adoopen(rtmp, n$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

  If rtmp.EOF Then Exit Function
  If typwert$ <> "" Then
    If InStr(LCase(trm(rtmp!wert)), typwert$) > 0 Then kisoftype = trm(rtmp!wert)
  Else
    kisoftype = trm(rtmp!wert)
  End If
  
End If

End Function

Public Function isofadr(t$, f$) As String
Dim i%

'd2infile = "Form1": d2insub = "isofadr"
isofadr = ""
i% = 0
While adrfeldcache(i%) <> ""
  If adrfeldcache(i%) = f$ Then
    isofadr = f$
    Exit Function
  End If
  i% = i% + 1
Wend
End Function

Private Sub combo1_LostFocus()
'd2infile = "Form1": d2insub = "combo1_LostFocus"
Timer1.Enabled = False

End Sub

Private Sub pin_KeyDown(KeyCode As Integer, Shift As Integer)

'd2infile = "Form1": d2insub = "pin_KeyDown"
If KeyCode = 13 Then
  Call Command8_Click
End If

End Sub

Private Sub prj_cal_Click()
Call Label5_dblClick
End Sub

Private Sub prj_chamb_Click()
Call Command21_Click
End Sub

Private Sub prj_cross_Click()
Call Command20_Click
End Sub

Private Sub prj_kstl_Click()
Call Command7_Click
End Sub

Private Sub prj_orch_Click()
Call Command4_Click
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub sqlmess_Click()
'd2infile = "Form1": d2insub = "sqlmess_Click"
Load agx
sqlmess.Visible = False
End Sub

Private Sub ta_chamb_Click()
Call Command21_Click
End Sub

Private Sub ta_cross_Click()
Call Command20_Click
End Sub

Private Sub ta_kstl_Click()
Call Command17_Click
End Sub

Private Sub ta_orch_Click()
Call Command4_Click
End Sub

Private Sub Timer1_Timer()
Dim s$

'd2infile = "Form1": d2insub = "Timer1_Timer"
If now() < snotb4 Then Exit Sub
Call dbg2f("Timer1 start")
Timer1.Enabled = False
s$ = Combo1.text
If s$ <> "" Then
  break% = 0
  Call rlist1(s$)
End If
Call dbg2f("Timer exit")

End Sub

Public Function repl1310htm(l$)
Dim r$, li$, i%, z$

'd2infile = "Form1": d2insub = "repl1310htm"
If InStr(l$, Chr$(13) + Chr$(10)) > 0 Then
  li$ = l$
  While InStr(li$, Chr$(13) + Chr$(10)) > 0
  For i% = 1 To Len(li$)
    z$ = Mid$(li$, i%, 1)
    If z$ = Chr$(13) Then
      i% = i% + 1
      r$ = r$ + "<br>"
    Else
      r$ = r$ + z$
    End If
  Next i%
  li$ = r$
  Wend
Else
  r$ = l$
End If
repl1310htm = r$

End Function

Public Function repl1310rtf(l$)
Dim r$, pt$, lin$, li$, i%, z$

'd2infile = "Form1": d2insub = "repl1310rtf"
lin$ = l$
If auftrittsdruck_currvorlage$ <> "" And auftrittsdruck_currfeld$ <> "" Then
  Call form1.dbg2f("repl1310rtf: " + auftrittsdruck_currvorlage$ + ", " + auftrittsdruck_currfeld$)
  pt$ = getusersetting(auftrittsdruck_currvorlage$ + "_" + auftrittsdruck_currfeld$, "N/A")
  If pt$ <> "N/A" Then
    lin$ = strrepl(pt$, "$wert$", l$)
  End If
  pt$ = getusersetting(auftrittsdruck_currvorlage$ + "_" + auftrittsdruck_currfeld$ + "_" + l$, "N/A")
  If pt$ <> "N/A" Then
    lin$ = pt$
  End If
End If
If InStr(lin$, Chr$(13) + Chr$(10)) > 0 Then
  li$ = lin$
  While InStr(li$, Chr$(13) + Chr$(10)) > 0
  For i% = 1 To Len(li$)
    z$ = Mid$(li$, i%, 1)
    If z$ = Chr$(13) Then
      i% = i% + 1
      r$ = r$ + "\par "
    Else
      r$ = r$ + z$
    End If
  Next i%
  li$ = r$
  Wend
Else
  r$ = lin$
End If
repl1310rtf = r$

End Function
Public Function fieldnameonly(fn)
Dim f$

'd2infile = "Form1": d2insub = "fieldnameonly"
fieldnameonly = fn
f$ = fieldnameonly
If InStr(fn, ".") Then
  f$ = Mid$(fn, InStr(fn, ".") + 1)
  If InStr(f$, ".") Then f$ = Left$(f$, InStr(f$, ".") - 1)
End If
fieldnameonly = f$
End Function

Public Function newlfdid(t$, key$) As Long
Dim stmp As ADODB.Recordset, cmd$, rrr, nnr As Long

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "newlfdid"
cmd$ = "SELECT max(" + key$ + ") as lnum FROM " + t$
Set stmp = New ADODB.Recordset
stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  newlfdid = -1
  Exit Function
End If
nnr = trm0(stmp!lnum) + 1
cmd$ = "insert into " + t$ + " (" + key$ + ") values(" + trm(nnr) + ")"
Call sqlqry(cmd$)
newlfdid = nnr

End Function

Public Function newid(t$, key$, l)
Dim id$, stmp As ADODB.Recordset, le%, cmd$, rrr, stw As Boolean

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "newid"
le% = l
Do
  id$ = mkkey(le%)
  cmd$ = "SELECT * FROM " + t$ + " where " + key$ + "='" + id$ + "'"
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Function
  stw = stmp.EOF
  stmp.Close
Loop Until stw
newid = id$

End Function

Public Function newidbase(t$, key$, l, b$, p$)
Dim id$, stmp As ADODB.Recordset, le%, cmd$, rrr, stw As Boolean

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "newid"
le% = l
Do
  id$ = b$ + mkkey(le%) + p$
  cmd$ = "SELECT * FROM " + t$ + " where " + key$ + "='" + id$ + "'"
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Function
  stw = stmp.EOF
  stmp.Close
Loop Until stw
newidbase = id$

End Function

Public Function ortausadr(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "ortausadr"
ortausadr = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT ort FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function

ortausadr = "" & rtmp!ort & ""

End Function

Public Function plzausadr(kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "plzausadr"
plzausadr = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT plz FROM adresse where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function

plzausadr = "" & rtmp!plz & ""

End Function

Public Sub prgdruck(id$, V$, mitbesetzung As Integer, auftrittsid$)
Dim rrr, lfdnr As Integer
Dim o%, p%, nam$, vorlage$, ort$, dat$, ueb$, rev$, wid$, sid$, diri$, sol$, dau, t$, ln$, pb%, l0$
Dim rtmp As ADODB.Recordset, stmp As ADODB.Recordset, orch$, fn$, l$, q%, k$, d$, bz$, ftm$, kl$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "prgdruck"

If V$ <> "" Then vorlage$ = V$
If vorlage$ = "" Then vorlage$ = meineprgdruckvorlage()
If nexist(s0d$ & "\" + dbname$ + ".rtf\" & vorlage$) Then
  MsgBox "Vorlage unbekannt: " + s0d$ & "\" + dbname$ + ".rtf\" & vorlage$
  Exit Sub
End If

ueb$ = "": orch$ = "": diri$ = "": sol$ = """"
If auftrittsid$ <> "" Then
  ueb$ = "select * from auftritt where id='" & auftrittsid$ & "'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, ueb$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    ueb$ = trm(rtmp!ort & ", den " & datfromsql(rtmp!datum))
  End If
  dat$ = "select felddaten from auftritthigru where auftrittsid='" & auftrittsid$ & "' and feldname='Solist'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, dat$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then sol$ = trm(rtmp!felddaten)
  dat$ = "select felddaten from auftritthigru where auftrittsid='" & auftrittsid$ & "' and (feldname='Orchester' or feldname='Orch_Ensemble')"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, dat$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then orch$ = trm(rtmp!felddaten)
  dat$ = "select felddaten from auftritthigru where auftrittsid='" & auftrittsid$ & "' and feldname='Dirigent'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, dat$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then diri$ = trm(rtmp!felddaten)
  ueb$ = trm(getnamebyid(orch$) & ", " & ueb$ & ", " & getnamebyid(diri$) & ", " & getnamebyid(sol$))
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM programm where programmid='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  ort$ = "" & rtmp!Veranstaltungsort
  dat$ = "" & rtmp!anfangsdatum
'  ueb$ = "" & rtmp!ueberschrift
End If
o% = FreeFile
Open s0d$ & "\" & dbname$ + ".rtf\" & vorlage$ For Input As #o%
p% = FreeFile
fn$ = trm(form1.myuniquedocname(""))
If fn$ = "" Then Exit Sub
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  
  
  q% = InStr(l$, "PUBLISHERSLISTE")
  If q% > 0 Then
    l$ = ""
    l0$ = "\pard {\trowd \trgaph70\trleft-70 \cellx400\cellx3450\cellx6000\cellx7500\cellx8800\cellx9990\cellx10990\cellx11990 \pard \intbl "
    Print #p%, l0$
    Print #p%, "# \cell "
    Print #p%, "Title\cell ";
    Print #p%, "Original Band\cell ";
    Print #p%, "Arranged by\cell ";
    Print #p%, "Publisher\cell ";
    Print #p%, "Album\cell ";
    Print #p%, "Time\cell ";
    Print #p%, "GEMA - Nummer\cell ";
    Print #p%, "\pard \intbl \row }\pard"
    
    lfdnr = 1
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid,besetztid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rtmp.EOF
      
      wid$ = trm(rtmp!werkid): sid$ = ""
      If Left$(wid$, 4) = "SBZ:" Then
        sid$ = Mid$(wid$, 5)
        wid$ = form1.getsatzidbywerkid(sid$)
      End If
      k$ = getkompvornamenamebywerkid(wid$): kl$ = LCase(k$)
      If kl$ <> "pause" And kl$ <> "oder" Then
        Print #p%, l0$
        Print #p%, "" + trm(lfdnr) + "\cell "
        If sid$ = "" Then
          Print #p%, form1.repl1310rtf(getwerknamebyid(wid$));
        Else
          Print #p%, form1.repl1310rtf(getsatznamebyid(sid$) + " " + transe("aus") + " " + getwerknamebyid(wid$));
        End If
        Print #p%, "\cell ";
        Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        k$ = wrk_arrandedby(wid$): Print #p%, "" + form1.repl1310rtf(strrepl(k$, "|", "\par ")) + "\cell ";
        k$ = wrk_publishedby(wid$): Print #p%, "" + form1.repl1310rtf(strrepl(k$, "|", "\par ")) + "\cell ";
        k$ = getalbumbywerkid(wid$): Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        k$ = getdauerbywerkid(wid$): Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        k$ = getgemanrbywerkid(wid$): Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        Print #p%, "\pard \intbl \row }\pard"
        lfdnr = lfdnr + 1
      End If
      rtmp.MoveNext
    Wend
  End If
  
  
  q% = InStr(l$, "GEMALISTE")
  If q% > 0 Then
    l$ = ""
     l0$ = "\pard {\trowd \trgaph70\trleft-70 \cellx400\cellx3450\cellx7500\cellx8800\cellx9990 \pard \intbl "
    Print #p%, l0$
    Print #p%, "# \cell "
    Print #p%, "Title\cell ";
    Print #p%, "Original Band\cell ";
    Print #p%, "Time\cell ";
    Print #p%, "GEMA #\cell ";
    Print #p%, "\pard \intbl \row }\pard"

    lfdnr = 1
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid,besetztid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rtmp.EOF
      
      wid$ = trm(rtmp!werkid): sid$ = ""
      If Left$(wid$, 4) = "SBZ:" Then
        sid$ = Mid$(wid$, 5)
        wid$ = form1.getsatzidbywerkid(sid$)
      End If
      k$ = getkompvornamenamebywerkid(wid$): kl$ = LCase(k$)
      If kl$ <> "pause" And kl$ <> "pause pause" And kl$ <> "oder" Then
        Print #p%, l0$
        Print #p%, "" + trm(lfdnr) + "\cell "
        If sid$ = "" Then
          Print #p%, form1.repl1310rtf(getwerknamebyid(wid$));
        Else
          Print #p%, form1.repl1310rtf(getsatznamebyid(sid$) + " " + transe("aus") + " " + getwerknamebyid(wid$));
        End If
        Print #p%, "\cell ";
        Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        k$ = getdauerbywerkid(wid$): Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        k$ = getgemanrbywerkid(wid$): Print #p%, "" + form1.repl1310rtf(k$) + "\cell ";
        Print #p%, "\pard \intbl \row }\pard"
        lfdnr = lfdnr + 1
      End If
      rtmp.MoveNext
    Wend
  End If
  
  q% = InStr(l$, "PRGDRUCK")
  If q% > 0 Then
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid,besetztid FROM programmliste where programmid='" + id$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rtmp.EOF
        wid$ = trm(rtmp!werkid): sid$ = ""
        If Left$(wid$, 4) = "SBZ:" Then
          sid$ = Mid$(wid$, 5)
          wid$ = form1.getsatzidbywerkid(sid$)
        End If
        Print #p%, "\trowd \trgaph70\trleft-70 \cellx2410\cellx9476 \pard \intbl "
        k$ = getkompvornamenamebywerkid(wid$)
        d$ = getkompdatesbywerkid(wid$)
        If d$ <> "" Then d$ = "(" + d$ + ")"
        If Left$(LCase$(k$), 7) = "pause p" Then
          k$ = "Pause"
          d$ = ""
        End If
        If Left$(LCase$(k$), 7) = "oder od" Then
          k$ = ""
          d$ = ""
        End If
        Print #p%, "{"; form1.repl1310rtf("" & k$ & ""); "\par "; d$; "\cell "
        If k$ <> "Pause" And k$ <> "Oder" Then
          If sid$ = "" Then
            Print #p%, form1.repl1310rtf(getwerknamebyid(wid$));
          Else
            Print #p%, form1.repl1310rtf(getsatznamebyid(sid$) + " " + transe("aus") + " " + getwerknamebyid(wid$));
          End If
        End If
        Print #p%, "\par \par "
        If sid$ = "" Then
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT satzbezeichnung FROM sbz_loc where wid='" + wid$ + "' order by satznummer", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not stmp.EOF
            If Left(LCase("" & stmp!satzbezeichnung), 7) <> "noten: " Then Print #p%, form1.repl1310rtf("" & stmp!satzbezeichnung & ""); "\par "
            stmp.MoveNext
          Wend
        End If
        If mitbesetzung > 0 Then
          bz$ = ""
          If IsNull(rtmp!besetztid) Then
             bz$ = form1.defaultbesetzt(wid)
          Else
             bz$ = form1.bestzstr(rtmp!besetztid)
          End If
          If Left$(bz$, 9) = "Standard:" Then bz$ = trm(Mid$(bz$, 10))
          If bz$ <> "" Then Print #p%, "\par Besetzung: " & form1.repl1310rtf("" & bz$ & ""); "\par "
        End If
        dau = getdauerbywerkid(wid$)
        If dau <> "" And dau <> "0" And sid$ = "" Then Print #p%, form1.repl1310rtf("" & dau & " Minuten"); "\par "
        Print #p%, "\cell \pard \intbl \row }\pard"
        rtmp.MoveNext
    Wend
  Else
    While Len(l$) > 0
      q% = InStr(l$, bkmstart$)
      If q% > 0 Then
        t$ = Mid$(l$, q% + Len(bkmstart$))
        Print #p%, Left$(l$, q% - 1)
        t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
        If InStr(t$, "user__") > 0 Then
          rev$ = trm(Mid$(t$, InStr(t$, "__") + 2))
          ftm$ = form1.getusersetting(rev$)
          Print #p%, repl1310rtf(ftm$);
        Else
          Select Case t$
            Case "system__datum": Print #p%, trm(Date)
            Case "ueberschrift": Print #p%, ueb$
            Case "id": Print #p%, id$
            Case "veranstaltungsort": Print #p%, ort$
            Case "datum": Print #p%, dat$
            Case Else
          End Select
        End If
        ln$ = Mid$(l$, q% + 1)
        Do
            pb% = InStr(LCase(ln$), bkmend$ + t$)
            If pb% = 0 Then Line Input #o%, ln$
        Loop Until pb% > 0
        ln$ = Mid$(ln$, pb%)
        If InStr(ln$, "}") = 0 Then
            l$ = ""
        Else
            l$ = Mid$(ln$, InStr(ln$, "}") + 1)
        End If
      Else
        Print #p%, l$
        l$ = ""
      End If
    Wend
  End If

Wend
Close #o%
Close #p%

Call form1.openthisdoc(fn$, "")

End Sub

Public Function dayofweek(dtg) As String
Dim d$, rrr

'd2infile = "Form1": d2insub = "dayofweek"
dayofweek = ""
d$ = ""
If trm(dtg) <> "" Then
  On Error Resume Next
  d$ = dayname(Weekday(dtg))
  rrr = Err
  On Error GoTo 0
End If
If rrr = 0 Then
  dayofweek = form1.inmylanguage(d$)
Else
  dayofweek = "n/a"
End If
End Function
Public Function longdayofweek(dtg) As String

'd2infile = "Form1": d2insub = "longdayofweek"
longdayofweek = ""
If trm(dtg) <> "" Then longdayofweek = longdayname(Weekday(dtg))

End Function

Function ersterauftritt(id$) As String
Dim c$
Dim rtmp As ADODB.Recordset

ersterauftritt = ""
c$ = "SELECT auftritt.ID as fdata FROM auftritt "
c$ = c$ + "Where auftritt.TourneeplanID = '" + id$ + "' And auftritt.auftrittstyp = 'Künstlerauftritt' ORDER BY auftritt.Datum, auftritt.Zeit"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rtmp.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If Not rtmp.EOF Then ersterauftritt = rtmp!fdata
End Function

Public Function auftrittsdruck(id$, vorlage$, wmode$, l4adr As String) As String
Dim d$, rvp%, lprg$, cmx$, sidkid$, fc%, rvf$, kstart$, ln$, pb%, al_v$, p_add%, nfno$, rfeld$, kid$
Dim cpm%, brkcpm%, lx$, px%, tr$, bk%, w$, pprov As String, prvk$, k$, dau$, ston$, pacont As Boolean
Dim xa_typ$, xa_kont$, xa_knam$, pax$, ava, al%, al_id$, al_vorlage$, rtf_vorlage$, rvl%, avl%, ctmp$
Dim fkt As String, pa$, cmd$, zw$, xa_adr$, xa_feld$, zn$, func$, felddaten$, xa_erg$, palstd As Boolean
Dim rrr, programmid$, datumfuerkurs$, fn$, q%, t$, ot$, ot0$, ttest$, prependlater$, adr_anred$, adr_abred$
Dim auftrittsaison$, langwochentag$, gfa$, amwst%, udn$, bruttohonorar As Double, bhfremd As String, rvfof%, rvol$
Dim mwsthonorar As Double, tarahonorar As Double, nettohonorar As Double, onlyif$, anyvatnumber As String
Dim adr_tel$, adr_fax$, adr_plzort$, nfn$, hfn$, sidp%, c$, vorl$, i%, docdtg$, j%, wochentag$
Dim adr_stra$, adr_land$, adr_pa$, adr_plz$, adr_ort$, adr_postfach$, adr_plzpostfach$, vorlfn$
Dim l$, wert$, bkms%, land$, plz$, ort$, plzort As String, adr_id$, adr_nam$, adr_adressnam$
Dim o%, o0%, p%, p0%, nam$, dat$, ueb$, nop As Integer, i0%, j0%, l0$, rrtmp As ADODB.Recordset
Dim rtmp As ADODB.Recordset, a As ADODB.Recordset, rv As ADODB.Recordset, rev$, orev$, strt As Boolean
Dim provisionaufbrutto_netto As Double, nfnv As Boolean, vorlagefn$, s As ADODB.Recordset
Dim provisionaufbrutto_mwst As Double, hK As Double, cutoff$, seriouswarning As Boolean
Dim provisionaufbrutto_brutto As Double, s1 As Double, awae$, xa_afld$, xanz As Double, xmwst As Double, xnet As Double
Dim bruttohonabzglprovbrutto As Double, prgo%, rechbez$, wk As ADODB.Recordset, rwert$
Dim rprog As ADODB.Recordset, stmp As ADODB.Recordset, xmplarc%, tanwaf As Boolean
Dim al_r As ADODB.Recordset, udat As ADODB.Recordset, auf As ADODB.Recordset, aufh As ADODB.Recordset
Dim hdat As ADODB.Recordset, sqludn, sqludw
Dim tpid$, todofl%, kurs As Double, betrag As Double, halle$, hinw$, la$
Dim marke$, t0$, meinzeichen$, sid$, sida$, sidk$, iwcurrent$, iwl As Long
Dim listenzaehler%, listenummer%, vnm$, sidkfeld$, usesidk As Boolean, adruckmerkslot As Integer
Dim sbzid$, werkid$, createdocindir As String, createdocinsubdir As String, kkurz1$, nameonly As Boolean, esz As String
Dim engwochentag$, englangwochentag$, pardepress As Boolean, dbguo$, adrnameonly As Boolean
Dim bruttohonorarwaehrung As String, meinewhrng As String, hkdat As String, xtake As Boolean
Dim ldblquotereplace As String, l4id$
Dim glcnt As Long, provisionsfeld$
Dim xprov As Double, xpmwst As Double, xzw As Double, xa2$, xwae$, xdat$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittsdruck"
listenzaehler% = 0
seriouswarning = False
pardepress = False
honvalid = 0
tanwaf = False
glcnt = 0

Call delalias
If form1.getusersetting("textalsnamewennadressefehlt", "nein") = "ja" Then tanwaf = True
sys_mwst = var2dbl(trm(getusersetting("MwSt", "1900")))
thismwst = sys_mwst
meinzeichen$ = getusersetting("meinzeichen")
If meinzeichen$ = "" Then meinzeichen$ = initialen(uname$)
todofl% = 1
palstd = False
If Left$(vorlage$, 1) = "_" Then
  vorlage$ = Mid$(vorlage$, 2)
  todofl% = 0
End If
If InStr(LCase(vorlage$), "adressen_cloud-export") > 0 Then Exit Function

If vorlagencache <> "" Then
  l$ = vorlage$: rrr = 0
  vorlfn$ = vorlagencache + "\" + FileName(vorlage$)
  If nexist(vorlfn$) Then
    On Error Resume Next
    Call FileCopy(l$, vorlfn$)
    rrr = Err
    On Error GoTo 0
  End If
  If rrr <> 0 Or nexist(vorlfn$) Then
    vorlagencache = ""
    vorlfn$ = s0dir & "\" & dbname$ & ".rtf\" & vorlage$
  End If
  vorlage$ = vorlfn$
End If

If exist(vorlage$) = 0 Then
  MsgBox "Vorlage unbekannt: " + vorlage$
  auftrittsdruck = ""
  Exit Function
End If
xmplarc% = 1
If Not nexist(vorlage$ & ".ini") Then
  o% = FreeFile: createdocindir = ""
  Open vorlage$ & ".ini" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    p% = InStr(l$, "=")
    If p% > 0 Then
      wert$ = Mid(l$, p% + 1)
      Select Case Left(LCase(l$), p% - 1)
        Case "exemplare"
          xmplarc% = Val(wert$)
        Case "speichern_in"
          createdocindir = getaliasfeld(wert$)
        Case "speichern_insubdir"
          createdocinsubdir = wert$
        Case Else
      End Select
    End If
  Wend
  Close #o%
End If
If todofl% = 1 Then
  For listenummer% = 0 To 14: todo(listenummer%).Clear: Next listenummer%
End If
bkms% = 0: bkmlcount%(bkms%) = 0
bkmlist$(bkms%, bkmlcount%(bkms%)) = "meinzeichen"
bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
rechbez$ = ""
If wmode$ = "adresse" Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adresse where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  land$ = "": plz$ = "": ort$ = "": plzort = ""
  If Not rtmp.EOF Then
    adr_id$ = id$
    adr_nam$ = trmx1(rtmp!postanrede & " " & trmx1(rtmp!name))
    adr_adressnam$ = trmx1(rtmp!postanrede & " " & trmx1(rtmp!name))
    rechbez$ = rechbez$ + " " + adr_adressnam$ + " {ID: " + id$ + "}"
    If Not IsNull(rtmp!strasse) Then adr_stra$ = rtmp!strasse
    If Not IsNull(rtmp!land) Then adr_land$ = rtmp!land
    If Not IsNull(rtmp!postanrede) Then adr_pa$ = rtmp!postanrede
    If LCase(adr_land$) = LCase(getusersetting("meinland")) Then land = ""
    If Not IsNull(rtmp!plz) Then adr_plz$ = rtmp!plz
    If Not IsNull(rtmp!ort) Then
      adr_ort$ = rtmp!ort
    End If
    adr_postfach$ = trm(rtmp!postfach)
    adr_plzpostfach$ = trm(rtmp!plzpostfach)
    If Not IsNull(rtmp!tel) Then adr_tel$ = rtmp!tel
    If Not IsNull(rtmp!fax) Then adr_fax$ = rtmp!fax
  End If
  adr_plzort$ = form1.getplzort(adr_land$, adr_plz$, adr_ort$)
  anyvatnumber = ""
  If AuftrittsdruckFuerAdresse$ <> "" Then
    c$ = "select FeldDaten from auftritthigru where auftrittsid='" + AuftrittsdruckFuerAdresse$ + "' and FeldName='VATNumber'"
    Set a = New ADODB.Recordset
    a.CursorLocation = adUseServer
    rrr = form1.adoopen(a, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not a.EOF Then
      anyvatnumber = trm(a!felddaten)
    End If
  End If
  nfn$ = form1.myuniquedocname("")
  Call dbg2f("auftrittsdruck: nfn=" + nfn$)
  GoTo nohavetplan
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftritt where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
tpid$ = ""
If Not IsNull(rtmp!TourneeplanID) Then tpid$ = rtmp!TourneeplanID
If Len(rechbez$) < 128 Then
  rechbez$ = rechbez$ + " Termin " + trm(rtmp!datum) + " " + trm(rtmp!zeit) + " / " + trm(tpid$)
  If trm(rtmp!bezeichnung) <> "" Then rechbez$ = rechbez$ + ", " + trm(rtmp!bezeichnung)
End If
If todofl% = 1 Then
  For listenummer% = 0 To 14: todo(listenummer%).Clear: Next listenummer%
  vnm$ = form1.getusersetting("TerminDokumentname", "Typ-Termindatum-User")
  hfn$ = datum2sql(rtmp!datum)
  hfn$ = Mid$(hfn$, 3, 2) & Mid$(hfn$, 6, 2) & Mid$(hfn$, 9, 2)
  nfn$ = ""
  nfnv = False
  sida$ = ""
  If Left(vnm$, 4) <> "Typ-" Then
    sid$ = auftritt.Text2(0).text
    sidp% = InStr(sid$, "{")
    sida$ = sid$
    If p% > 1 Then
      sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
    End If
    c$ = "select kdnr from adresse where id='" & sida$ & "'"
    Set a = New ADODB.Recordset
    a.CursorLocation = adUseServer
rrr = form1.adoopen(a, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not a.EOF Then
      sida$ = trm(a!kdnr)
    End If
  End If
  vorl$ = Mid$(vorlage$, InStr(vorlage$, "_") + 1)
  vorl$ = Left$(vorl$, InStr(vorl$, ".") - 1)
  For i% = 1 To Len(vorl$)
    If Mid$(vorl$, i%, 1) <> LCase(Mid$(vorl$, i%, 1)) Then
      nfn$ = nfn$ & Mid$(vorl$, i%, 1)
      If Not isdigit(Mid$(vorl$, i%, 1)) <> 0 Then nfnv = True
    End If
  Next i%
  If Not nfnv Then
    hfn$ = vorl$ & hfn$
  Else
    hfn$ = nfn$ & hfn$
  End If
  If sida$ <> "" Then hfn$ = sida$ & "-" & hfn$
  If tpid$ <> "-1" Then hfn$ = word1(tpid$) & "-" & FileName(hfn$)
  If createdocindir = "" Then
    If wmode$ <> "ical" Then
      nfn$ = form1.myuniquedocname("FN:" & hfn$)
    Else
      nfn$ = form1.s0dir() + "\" + form1.medien() + "\icaltxt.txt"
      On Error Resume Next: Kill nfn$: On Error GoTo 0
    End If
  Else
    createdocindir = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and feldname='" + createdocindir + "';"
    Set rrtmp = New ADODB.Recordset
    rrtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rrtmp, createdocindir, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrtmp.EOF Then
      createdocindir = ""
      If wmode$ = "ical" Then
        nfn$ = form1.s0dir() + "\" + form1.medien() + "\icaltxt.txt"
        On Error Resume Next: Kill nfn$: On Error GoTo 0
      Else
        nfn$ = form1.myuniquedocname("FN:" & hfn$)
      End If
    Else
      createdocindir = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(trm(rrtmp!felddaten))
      If trm(createdocinsubdir) <> "" Then createdocindir = createdocindir + "\" + createdocinsubdir
      nfn$ = trm(myuniquedocnameinpath("FN:" + createdocindir + "\" + hfn$, ""))
    End If
  End If
  If Len(trm(nfn$)) = 0 Then Exit Function
End If

If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.Caption = "erstelle Dokument"
If wmode$ <> "ical" Then tplan.MousePointer = 11
docdtg$ = datum2sql(Date) & " " & Time
If todofl% = 1 Then
  dochistlist.Clear
  Call clear_honorarliste
End If
'creating bookmarklist
wochentag$ = dayofweek(CDate(datfromsql(rtmp!datum)))
langwochentag$ = longdayofweek(CDate(datfromsql(rtmp!datum)))
auftrittsaison$ = saison(trm(datfromsql(rtmp!datum)))
engwochentag$ = dictionarylookup(wochentag$)
englangwochentag$ = dictionarylookup(langwochentag$)
Do
  bkmlist$(bkms%, bkmlcount%(bkms%)) = sqla.TableDefs("auftritt").Fields(bkmlcount%(bkms%)).name
  bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
Loop Until bkmlcount%(bkms%) >= sqla.TableDefs("auftritt").Fields.Count
bkmlist$(bkms%, bkmlcount%(bkms%)) = "saison": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "terminende": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "wochentag": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "langwochentag": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "engwochentag": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "englangwochentag": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "mwst0text": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
bkmlist$(bkms%, bkmlcount%(bkms%)) = "bruttohonorarxumrechnung": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
'bkmlcount%(bkms%) = bkmlcount%(bkms%) - 1   astatus nicht verschlucken
gfa$ = "honorar"
If LCase(auftrittstyp(id$)) = "orchesterauftritt" Then
  gfa$ = getaliasfeld(gfa$)
  If gfa$ <> "honorar" Then
    If AuftrittsdruckFuerAdresse$ <> "" Then
      gfa$ = auftrittshonorarfeldbyname(id$, AuftrittsdruckFuerAdresse$)
    End If
  End If
End If
c$ = "": provisionsfeld$ = "provision"
If LCase(auftrittstyp(id$)) = "künstlerauftritt" Then
  gfa$ = getaliasfeld(gfa$)
  If gfa$ = "honorar" Then
    If AuftrittsdruckFuerAdresse$ <> "" Then
      c$ = auftrittshonorarfeldbyadrid(id$, AuftrittsdruckFuerAdresse$)
      If InStr(LCase(c$), "honorar") = 1 And Len(c$) = 8 Then
        provisionsfeld$ = "provision" + onlynums(c$)
        Call addalias("honorar", c$)
        Call addalias("provision", provisionsfeld$)
        c$ = "select -1 as mwst,1 as anz," + c$ + " as netto from usr_künstlerauftritt where id='" & id$ & "'"
      Else
        c$ = ""
      End If
    End If
  End If
End If
If gfa = "honorar" Then
  If c$ = "" Then
    c$ = "select mwst,anz,netto,waehrung from finanzen where id='" & id$ & "'"
  End If
Else
  If c$ = "" Then
    c$ = "select mwst,anz,netto,waehrung from finanzen where id='" & gfa$ & "(ID:" & id$ & "'"
  End If
End If
Set udat = New ADODB.Recordset
udat.CursorLocation = adUseServer
rrr = form1.adoopen(udat, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not udat.EOF Then
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "programmid":  bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "mwst": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "provisionaufhonorarva_netto": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "provisionaufbrutto_netto": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "provisionaufbrutto_mwst": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "provisionaufbrutto_brutto": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "bruttohonabzglprovbrutto": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "mwsthonorar": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "tarahonorar": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "nettohonorar": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "bruttohonorarumrechnung": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "bruttohonorarumrechnungcrlf": bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  amwst% = trm0(udat!mwst)
  If amwst% = -1 Then
    sqludn = ohnewaehrung(trm0(udat!netto))
    sqludw = nurdiewaehrung(trm0(udat!netto))
  Else
    sqludn = trm(udat!netto)
    sqludw = trm(udat!waehrung)
  End If
  udn$ = trm(sqludn)
  If udn$ = "" Then udn$ = "0"
  bruttohonorar = var2dbl(trm(udat!anz)) * var2dbl(udn$)
  If l4adr$ = "" Then
    bruttohonorar = var2dbl(HonorarVonAuftrittByAdr(id$, ""))
  End If
  bruttohonorar = bruttohonorar + (bruttohonorar * MwStFuerAuftritt(id$) / 10000)
  If sqludw = "" Then sqludw = "€"
  bruttohonorarwaehrung = trm(sqludw)
  If form1.getusersetting("MeineWaehrung", transe("€")) <> sqludw Then
  awae$ = bruttohonorarwaehrung
  bhfremd = ""
  meinewhrng = form1.getusersetting("MeineWaehrung", transe("€"))
  hK = 1
  If awae$ <> "" And meinewhrng <> awae$ Then
    bhfremd = "(" & bruttohonorar & " " & awae$
    hK = var2dbl(strrepl(kursvom(awae$, rtmp!datum), ".", ","))
    hkdat = kursdatum(awae$, rtmp!datum)
    'hK = CCur(hK)
    If hK = 0 Then hK = 10000000
    On Error Resume Next
    s1 = CCur(ohnewaehrung(trm(bruttohonorar))) / hK
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      bruttohonorar = fixeur(s1)
      bruttohonorarwaehrung = meinewhrng
    End If
    bhfremd = bhfremd & ", " + transe("Kurs") + ": " & trm(hK) & " " & meinewhrng & "/" & awae$ & " " + transe("am") + " " & hkdat
    bhfremd = bhfremd & ")"
  End If
  mwsthonorar = bruttohonorar / ((100 + MwStFuerAuftritt(id$)) / 100)
  mwsthonorar = mwsthonorar * MwStFuerAuftritt(id$) / 100
  nettohonorar = bruttohonorar - mwsthonorar
  tarahonorar = bruttohonorar - nettohonorar
  End If
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "bruttohonorar"
  bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
End If
bkms% = 1: bkmlcount%(bkms%) = 0
Set a = New ADODB.Recordset
a.CursorLocation = adUseServer
rrr = form1.adoopen(a, "SELECT feldname FROM auftrittsfelder where typ='" + rtmp!auftrittstyp + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not a.EOF
  l$ = getaliasfeld(a!feldname)
  bkmrflag$(bkmlcount%(bkms%)) = ""
  If InStr(l$, ".") > 0 Then
    bkmrflag$(bkmlcount%(bkms%)) = Left$(l$, InStr(l$, ".") - 1)
    If bkmrflag$(bkmlcount%(bkms%)) = "adrselect" Then bkmrflag$(bkmlcount%(bkms%)) = "adresse"
    l$ = Mid$(l$, InStr(l$, ".") + 1)
    If InStr(l$, ".") > 0 Then
      l$ = Left$(l$, InStr(l$, ".") - 1)
    End If
  End If
  If LCase(l$) = provisionsfeld$ Then
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
    rrr = form1.adoopen(s, "SELECT felddaten FROM auftritthigru where feldname='" + l$ + "' and auftrittstyp='" & rtmp!auftrittstyp & "' and auftrittsid='" & id$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rrr = 0 Then
      If Not s.EOF Then
        felddaten$ = trm(s!felddaten)
        If InStr(felddaten$, "/") > 0 Then
          thismwst = var2dbl(trm(cut_d1(trm(cut_d2bis(felddaten$, "/")), "%"))) * 100
        End If
      End If
    End If
  End If
  If LCase(l$) = "programm" Then
    programmid$ = ""
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
rrr = form1.adoopen(s, "SELECT felddaten FROM auftritthigru where feldname='Programm' and auftrittstyp='" & rtmp!auftrittstyp & "' and auftrittsid='" & id$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not s.EOF Then
      programmid$ = trm(s!felddaten)
    End If
  End If
  bkmlist$(bkms%, bkmlcount%(bkms%)) = l$
  bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
  a.MoveNext
Wend
For i% = 0 To 9
  bkmlist$(bkms%, bkmlcount%(bkms%)) = "merkslot" + trm(i%)
  bkmrflag$(bkmlcount%(bkms%)) = "adresse"
  bkmlcount%(bkms%) = bkmlcount%(bkms%) + 1
Next i%
bkmlcount%(bkms%) = bkmlcount%(bkms%) - 1
datumfuerkurs$ = datum2sql(Date)

nohavetplan:
Set udat = New ADODB.Recordset
udat.CursorLocation = adUseServer
rrr = form1.adoopen(udat, "SELECT * FROM benutzerdaten where id ='" + uId$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While xmplarc% > 0
xmplarc% = xmplarc% - 1
o% = FreeFile
vorlagefn$ = FileName(vorlage$)
auftrittsdruck_currvorlage$ = vorlagefn$
Call dbg2f("auftrittsdruck: reading " + vorlagefn$)
Open vorlage$ For Input As #o%
ldblquotereplace = form1.getusersetting("lquotereplace", "ja")
If ldblquotereplace = "ja" And InStr(LCase(vorlage$), "engl") > 0 Then
  ldblquotereplace = "nein"
End If
p% = FreeFile
If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.List1.AddItem "Vorlage: " & vorlage$: dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
fn$ = form1.myuniquedocname("noask")
Call dbg2f("auftrittsdruck: unique name=" + fn$)
If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.List1.AddItem "erstelle Datei: " & fn$: dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
Open fn$ For Output As #p%
While Not EOF(o%)
  l$ = ""
  Do
If shwled Then
  cb1.BackColor = RGB(255, 128, 0)
  cb1.Cls
End If
    Line Input #o%, l0$
    glcnt = glcnt + 1
    If (glcnt Mod 1000) = 0 Then
      If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.List1.AddItem trm(glcnt) + " Zeilen gelesen": dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
      Debug.Print glcnt; " Zeilen gelesen"
    End If
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
    If ldblquotereplace = "ja" And InStr(l0$, "\ldblquote ") > 0 Then
      l0$ = strrepl(l0$, "\ldblquote ", "\'84")
    End If
    l$ = l$ & l0$
  Loop Until bkmktest1(l$) = True
'  Loop Until bkmktest1(l$) = True Or Len(l$) > 16000
  l$ = replacealiastext(l$)
  While Len(l$) > 0
    q% = InStr(l$, bkmstart$)
    auftrittsdruck_currfeld$ = ""
    If q% > 0 Then
      t$ = Mid$(l$, q% + Len(bkmstart$)): marke$ = ""
      Print #p%, Left$(l$, q% - 1);
      t0$ = Left$(t$, InStr(t$, "}") - 1)
      ot0$ = t0$
      t$ = LCase(t0$)
      If Left$(t0$, 5) = "MARKE" Then
        If isdigit(Mid$(t0$, 6, 1)) <> 0 Then
          If Mid$(t0$, 7, 1) = "_" Then
            marke$ = Left$(t0$, 7)
            t$ = Mid$(t$, 8)
            ot0$ = Mid$(ot$, 8)
          End If
        End If
      End If
      If LCase(Left$(t0$, 1)) = "m" And Mid$(t0$, 3, 1) = "_" Then
        If Mid$(t0$, 3, 1) = "_" Then
          marke$ = Left$(t0$, 3)
          t$ = Mid$(t$, 4)
        End If
      End If
      adruckmerkslot = -1
      If LCase(Left$(t0$, 3)) = "as_" And isnumber(Mid$(t0$, 4, 1)) And Mid$(t0$, 5, 1) = "_" Then
        adruckmerkslot = Val(Mid$(t0$, 4, 1))
        t$ = Mid$(t$, 6)
      End If
      t$ = strrepl(t$, "\'fc", "ü")
      t$ = strrepl(t$, "\'e4", "ä")
      t$ = strrepl(t$, "\'f6", "ö")
      ot0$ = strrepl(ot0$, "\'fc", "ü")
      ot0$ = strrepl(ot0$, "\'e4", "ä")
      ot0$ = strrepl(ot0$, "\'f6", "ö")
      orev$ = ""
      If InStr(ot0$, "__") > 0 Then
        orev$ = Mid$(ot0$, InStr(ot0$, "__") + 2)
      End If
Debug.Print t$; "-->";
      nameonly = False
      adrnameonly = False
      rev$ = "": ttest$ = getaliasfeld(t$)
      If InStr(ttest$, "__") > 0 Then
        rev$ = Mid$(ttest$, InStr(ttest$, "__") + 2)
        If LCase(rev$) = "nurdername" Then
          rev$ = "name"
          nameonly = True
        End If
        If LCase(rev$) = "nuradressname" Then
          rev$ = "name"
          adrnameonly = True
        End If
        ttest$ = Left$(ttest$, InStr(ttest$, "__") - 1)
      End If
Debug.Print ttest$; "(" & rev$ & ") orev=" & orev$
      auftrittsdruck_currfeld$ = t$
      If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.List1.AddItem "Textmarke: " & t$ & "-->" & ttest$ & "(" & rev$ & ")": dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
      prependlater$ = ""
      If wmode$ = "adresse" Then
        If ttest$ = "this" Then
          rfeld$ = Mid$(rev$, InStr(rev$, "__") + 2)
          rev$ = Left$(rev$, InStr(rev$, "__") - 1)
Debug.Print rfeld$; " - "; rev$
          ttest$ = adr_id$: If kid$ <> "" And kid$ <> "-1" Then ttest$ = ttest$ + kid$
          cmd$ = "select FeldDaten as rc from auftritthigru where auftrittsid='" + ttest$ + "' and lcase(auftrittstyp)='" + LCase(rev$) + "' and lcase(FeldName)='" + LCase(rfeld$) + "'"
          rwert$ = ""
          Set hdat = New ADODB.Recordset
          hdat.CursorLocation = adUseServer
          rrr = form1.adoopen(hdat, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If rrr = 0 Then
'          rwert$ = cmd$ + " := " + trm(hdat!rc)
            rwert$ = trm(hdat!rc)
          End If
          'Print #p%, Left$(l$, q% - 1);: Print #p%, strrepl(rwert$, vbCrLf, "\par ");
          Print #p%, strrepl(rwert$, vbCrLf, "\par ");
        End If

        Select Case LCase(ttest$)
          Case "name": Call adpprint(adruckmerkslot, p%, repl1310rtf(adr_nam$)): Call dbgu(adr_nam$)
          Case "anyvatnumber": Call adpprint(adruckmerkslot, p%, repl1310rtf(anyvatnumber)): Call dbgu(anyvatnumber)
          Case "adressname": Call adpprint(adruckmerkslot, p%, repl1310rtf(adr_adressnam$)): Call dbgu(adr_adressnam$)
          Case "strasse": Call adpprint(adruckmerkslot, p%, repl1310rtf(adr_stra$)): Call dbgu(adr_stra$)
          Case "ort": Call adpprint(adruckmerkslot, p%, strrepl(form1.repl1310rtf(adr_plzort$), "  ", " ")): Call dbgu(adr_plzort$)
          Case "nurort": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_ort$)): Call dbgu(adr_ort$)
          Case "land": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(land$)): Call dbgu(land$)
          Case "plz": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_plz$)): Call dbgu(adr_plz$)
          Case "tel": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(userformatphone(adr_tel$))): Call dbgu(adr_tel$)
          Case "fax": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(userformatphone(adr_fax$))): Call dbgu(adr_fax$)
          Case "postfach": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_postfach$)): Call dbgu(adr_postfach$)
          Case "plzort": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_plzort$)): Call dbgu(adr_plzort$)
          Case "plzpostfach": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_plzpostfach$)): Call dbgu(adr_plzpostfach$)
          Case "postanrede": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_pa$)): Call dbgu(adr_pa$)
          Case "anrede": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_anred$)): Call dbgu(adr_anred$)
          Case "abrede": Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(adr_abred$)): Call dbgu(adr_abred$)
          Case Else
        End Select
      End If
      pacont = False
      If ttest$ = "fx" Then
        If trm(rev$) <> "" Then
          fkt = Left$(rev$, InStr(rev$, "__") - 1)
          pa$ = Mid$(rev$, InStr(rev$, "__") + 2)
          Select Case LCase(fkt)
            Case "ersterauftritt":
              xa_afld$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_afld$ = getaliasfeld(xa_afld$)
              xa_feld$ = pa$
              If InStr(xa_feld$, "__") > 0 Then
                xa2$ = Mid$(xa_feld$, InStr(xa_feld$, "__") + 2)
                xa_feld$ = Left$(xa_feld$, InStr(xa_feld$, "__") - 1)
              End If
              zw$ = ersterauftritt(tpid$)
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & zw$ & "' and feldname='" & xa_afld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
              rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              zw$ = ""
              If Not rv.EOF Then
                zw$ = trm0(rv!felddaten)
              End If
              If zw$ <> "" Then
                zw$ = getAdrProperty(zw$, xa_feld$)
                Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(zw$)): Call dbgu(zw$)
              End If
            Case "reset":
              If pa$ = "honorarliste" Then Call clear_honorarliste
            Case "setl4adr":
              If InStr(pa$, "__") > 0 Then
                xa_afld$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
                xa_afld$ = getaliasfeld(xa_afld$)
              End If
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
              rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                l4adr = trm(rv!felddaten)
              End If
              zw$ = HonorarVonAuftrittByAdr(id$, l4adr)
              bruttohonorarwaehrung = nurdiewaehrung(zw$)
              nettohonorar = var2dbl(zw$)
              hK = 1
              If form1.getusersetting("MeineWaehrung", transe("€")) <> bruttohonorarwaehrung Then
                awae$ = bruttohonorarwaehrung
                bhfremd = ""
                meinewhrng = form1.getusersetting("MeineWaehrung", transe("€"))
                If awae$ <> "" And meinewhrng <> awae$ Then
                  bhfremd = "(" & nettohonorar & " " & awae$
                  hK = var2dbl(strrepl(kursvom(awae$, rtmp!datum), ".", ","))
                  hkdat = kursdatum(awae$, rtmp!datum)
                  'hK = CCur(hK)
                  If hK = 0 Then hK = 10000000
                  On Error Resume Next
                  s1 = CCur(ohnewaehrung(trm(nettohonorar))) / hK
                  rrr = Err
                  On Error GoTo 0
                  If rrr = 0 Then
                    nettohonorar = fixeur(s1)
                    bruttohonorarwaehrung = meinewhrng
                  End If
                End If
                bhfremd = bhfremd & ", " + transe("Kurs") + ": " & trm(hK) & " " & meinewhrng & "/" & awae$ & " " + transe("am") + " " & hkdat
                bhfremd = bhfremd & ")"
              End If
              bruttohonorar = nettohonorar + (bruttohonorar * MwStFuerAuftritt(id$) / 10000)
              mwsthonorar = bruttohonorar / ((100 + MwStFuerAuftritt(id$)) / 100)
              tarahonorar = bruttohonorar - nettohonorar
            Case "finanzen":
              If InStr(pa$, "__") > 0 Then
                xa_afld$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
                xa_afld$ = getaliasfeld(xa_afld$)
              End If
              xa_feld$ = pa$
              If InStr(xa_feld$, "__") > 0 Then
                xa2$ = Mid$(xa_feld$, InStr(xa_feld$, "__") + 2)
                xa_feld$ = Left$(xa_feld$, InStr(xa_feld$, "__") - 1)
              End If
              zw$ = ""
              If xa_feld$ = "provision1" Or xa_feld$ = "provision2" Or xa_feld$ = "provision3" Or xa_feld$ = "provision4" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" + xa_feld$ + "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                xpmwst = mwst
                zw$ = ""
                If Not rv.EOF Then
                  zw$ = trm0(cut_d1(trm0(rv!felddaten), "/"))
                End If
              End If
              If xa_feld$ = "provmp" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                xpmwst = mwst
                If Not rv.EOF Then
                  zw$ = strrepl(trm0(cut_d2bis(trm0(rv!felddaten), "/")), "%", "")
                  If zw$ <> "" Then
                    On Error Resume Next
                    xpmwst = 100 * CDbl(zw$)
                    rrr = Err
                    On Error GoTo 0
                    If rrr <> 0 And warnmeondata Then MsgBox ("Illegal field content: ..." + vbCrLf + trm0(rv!felddaten))
                  End If
                End If
                zw$ = trm(xpmwst / 100)
              End If
              If xa_feld$ = "provabn" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xmwst = form1.MwStFuerAuftritt(id$)
                  xnet = CDbl(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                End If
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa2$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xprov = CDbl(strrepl(trm0(cut_d1(trm0(rv!felddaten), "/")), "%", ""))
                  xpmwst = CDbl(strrepl(trm0(cut_d2bis(trm0(rv!felddaten), "/")), "%", ""))
                End If
                xzw = xnet + (xmwst * xnet / 100)
                If InStr(strrepl(cut_d1(trm0(rv!felddaten), "/"), " ", ""), "%") > 0 Then
                  zw$ = fixeur(xzw * xprov / 100) + " " + xwae$
                Else
                  zw$ = fixeur(xprov) + " " + xwae$
                End If
              End If
              If xa_feld$ = "provabm" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xmwst = form1.MwStFuerAuftritt(id$)
                  xnet = CDbl(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                End If
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa2$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xprov = CDbl(strrepl(trm0(cut_d1(trm0(rv!felddaten), "/")), "%", ""))
                  xpmwst = CDbl(strrepl(trm0(cut_d2bis(trm0(rv!felddaten), "/")), "%", ""))
                End If
                If InStr(strrepl(cut_d1(trm0(rv!felddaten), "/"), " ", ""), "%") > 0 Then
                  xzw = xnet + (xmwst * xnet / 100)
                  xzw = xzw * xprov / 100
                  xzw = xzw * xpmwst / 100
                Else
                  xzw = xprov * xpmwst / 100
                End If
                zw$ = fixeur(xzw) + " " + xwae$
              End If
              If xa_feld$ = "provabb" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xmwst = form1.MwStFuerAuftritt(id$)
                  xnet = CDbl(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                End If
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa2$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xprov = CDbl(strrepl(trm0(cut_d1(trm0(rv!felddaten), "/")), "%", ""))
                  xpmwst = CDbl(strrepl(trm0(cut_d2bis(trm0(rv!felddaten), "/")), "%", ""))
                End If
                If InStr(strrepl(cut_d1(trm0(rv!felddaten), "/"), " ", ""), "%") > 0 Then
                  xzw = xnet + (xmwst * xnet / 100)
                  xzw = xzw * xprov / 100
                  xzw = xzw + xzw * xpmwst / 100
                Else
                  xzw = xprov + xprov * xpmwst / 100
                End If
                zw$ = fixeur(xzw) + " " + xwae$
              End If
              If xa_feld$ = "waehrung" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                zw$ = "€"
                If Not rv.EOF Then
                  zw$ = nurdiewaehrung(trm0(rv!felddaten))
                End If
              End If
              If xa_feld$ = "netto2brutto" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xmwst = form1.MwStFuerAuftritt(id$)
                  xnet = CCur(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                  zw$ = fixeur(xnet + (xmwst * xnet / 10000))
                  zw$ = zw$ + " " + xwae$
                End If
              End If
              If xa_feld$ = "netto2netto" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xnet = CCur(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                  zw$ = fixeur(xnet)
                  zw$ = zw$ + " " + xwae$
                End If
              End If
              If xa_feld$ = "netto2mwst" Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_afld$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
                rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xmwst = form1.MwStFuerAuftritt(id$)
                  xnet = CDbl(ohnewaehrung(rv!felddaten))
                  xwae$ = nurdiewaehrung(trm0(rv!felddaten))
                  xdat$ = trm(rtmp!datum)
                  xnet = xkurs(xdat, xwae$, xnet)
                  xwae$ = form1.getusersetting("MeineWaehrung", transe("€"))
                  zw$ = fixeur(xmwst * xnet / 10000)
                  zw$ = zw$ + " " + xwae$
                End If
              End If
              If xa_feld$ = "mwstwert" Then
                cmd$ = "select anz,netto,mwst from finanzen where id='" + xa_afld$ + "(ID:" + id$ + "'"
                Set wk = New ADODB.Recordset
                wk.CursorLocation = adUseServer
                rrr = form1.adoopen(wk, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If rrr = 0 Then
                  If Not wk.EOF Then
                    xanz = CDbl(ohnewaehrung(wk!anz))
                    xmwst = CDbl(trm0(wk!mwst))
                    xnet = CDbl(trm0(wk!netto))
                    zw$ = fixeur(xanz * xmwst * xnet / 10000)
                  End If
                End If
              End If
              If xa_feld$ = "brutto" Then
                cmd$ = "select anz,netto,mwst from finanzen where id='" + xa_afld$ + "(ID:" + id$ + "'"
                Set wk = New ADODB.Recordset
                wk.CursorLocation = adUseServer
                rrr = form1.adoopen(wk, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If rrr = 0 Then
                If Not wk.EOF Then
                  xanz = CDbl(trm0(wk!anz))
                  xmwst = CDbl(trm0(wk!mwst))
                  xnet = CDbl(ohnewaehrung(wk!netto))
                  zw$ = fixeur((xanz * xnet) + (xanz * xmwst * xnet / 10000))
                End If
                End If
              End If
              If zw$ = "" Then
                cmd$ = "select " + pa$ + " as erg from finanzen where id='" + xa_afld$ + "(ID:" + id$ + "'"
                Set wk = New ADODB.Recordset
                wk.CursorLocation = adUseServer
                rrr = form1.adoopen(wk, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If rrr = 0 Then
                If Not wk.EOF Then
                  zw$ = trm(wk!erg)
                  If xa_feld$ = "netto" Then
                    zw$ = fixeur(CDbl(trm0(zw$)))
                  End If
                  If xa_feld$ = "mwst" Then
                    zw$ = strrepl(fixeur(CLng(trm0(zw$)) / 100), ".", "")
                    While Right(zw$, 1) = "0" And Len(zw$) > 1
                      zw$ = Left(zw$, Len(zw$) - 1)
                    Wend
                    If Right(zw$, 1) = "," Then zw$ = Left(zw$, Len(zw$) - 1)
                  End If
                End If
              End If
              End If
              If zw$ <> "" Then Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(zw$)): Call dbgu(zw$)
            Case "terminanzahl":
              cmd$ = "select id from auftritt where auftrittstyp='" + pa$ + "' and tourneeplanid='" + tpid$ + "'"
              Set wk = New ADODB.Recordset
              wk.CursorLocation = adUseServer
rrr = form1.adoopen(wk, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              p0% = 0
              While Not wk.EOF
                p0% = p0% + 1
                wk.MoveNext
              Wend
              p0% = p0% + form1.refeddates(tpid$)
              Print #p, form1.repl1310rtf(trm(p0%));: Call dbgu(zw$)
            Case "weitere":
              If InStr(pa$, "__") > 1 Then
                xa_adr$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
                xa_adr$ = getaliasfeld(xa_adr$)
              End If
              xa_feld$ = pa$
              cmd$ = "select * from auftritt where tourneeplanid='" + tpid$ + "' and auftrittstyp='" + xa_feld$ + "' order by datum,zeit;"
              Set wk = New ADODB.Recordset
              wk.CursorLocation = adUseServer
rrr = form1.adoopen(wk, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              While Not wk.EOF
                zw$ = ""
                If wk!id <> id$ Then
                  'Debug.Print wk!datum; " "; wk!zeit; " "; wk!ort; " "; wk!auftrittstyp; " "; wk!id
                  zw$ = printdatfromsql(trm(wk!datum)) + " " + trm(wk!zeit) + " Uhr in " + trm(wk!ort)
                End If
                wk.MoveNext
                If Not wk.EOF And zw$ <> "" Then zw$ = zw$ + vbCrLf
                If zw$ <> "" Then Print #p, form1.repl1310rtf(zw$);: Call dbgu(zw$)
              Wend
            Case "kalk":
              xa_adr$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_adr$ = getaliasfeld(xa_adr$)
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='kalku_" & rtmp!auftrittstyp & "' and feldname='" & xa_adr$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                l0$ = trm(rv!felddaten)
                i0% = linesof(l0$): j0% = 1
                While i0% > 0
                  i0% = i0% - 1
                  cmd$ = lineof(j0%, l0$): j0% = j0% + 1
                  p0% = InStr(cmd$, "|")
                  If p0% > 0 Then
                    zn$ = Left$(cmd$, p0% - 1)
                    zw$ = Mid$(cmd$, p0% + 1)
                    p0% = InStr(zw$, "|")
                    func$ = ""
                    If p0% > 0 Then
                      func$ = Mid$(zw, p0% + 1)
                      If p0% > 1 Then
                        zw$ = Left$(zw$, p0% - 1)
                      Else
                        zw$ = ""
                      End If
                    End If
                  End If
                  If LCase(zn$) = LCase(xa_feld$) Then
                    Print #p, form1.repl1310rtf(zw$);: Call dbgu(zw$)
                    iwcurrent$ = trm0(strrepl(zw$, ".", ""))
                    iwl = Val(iwcurrent$)
                    iwcurrent$ = inworten(iwl)
                    i0% = 0
                  End If
                Wend
              End If
            Case "mkan":
              xa_adr$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_adr$ = getaliasfeld(xa_adr$)
              xa_feld$ = pa$
              'adresse suchen
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_adr$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                'kontakt lesen
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
                Set stmp = New ADODB.Recordset
                stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not stmp.EOF Then
                  cmd$ = get_kontaktid_by_name(trm(rv!felddaten), trm(stmp!felddaten))
                  cmd$ = meineanrede(cmd$)
                  Print #p, form1.repl1310rtf(cmd$);: Call dbgu(cmd$)
                End If
              End If
            Case "man":
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
              rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                cmd$ = meineanrede("-1." & trm(rv!felddaten))
                Print #p, form1.repl1310rtf(cmd$);: Call dbgu(cmd$)
              End If
            Case "w2bis":
              xa_feld$ = pa$
              pacont = True
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
              rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                cmd$ = cut_d2bis(trm(rv!felddaten), " ")
                Print #p, form1.repl1310rtf(cmd$);: Call dbgu(cmd$)
              End If
            Case "langwochentag":
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
              rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                On Error Resume Next
                cmd$ = longdayofweek(CDate(cut_d1(trm(rv!felddaten), " ")))
                rrr = Err
                On Error GoTo 0
                If rrr <> 0 Then
                  Call dbgu("Fehler bei der Ermittlung des Wochentages für das Datum '" + cut_d1(trm(rv!felddaten), " ") + "' " + transe("aus") + " '" + trm(rv!felddaten) + "'")
                  seriouswarning = True
                  cmd$ = ""
                End If
                Print #p, form1.repl1310rtf(cmd$);: Call dbgu(cmd$)
              End If
            Case "exemplar"
              If pa$ = "alphanummer" Then
                Print #p, form1.repl1310rtf(trm(Chr$(Asc("A") + xmplarc%)));: Call dbgu(trm(Chr$(Asc("A") + xmplarc%)))
              End If
              If pa$ = "nummer" Then
                Print #p, form1.repl1310rtf(trm(str$(xmplarc%)));: Call dbgu(trm(str$(xmplarc%)))
              End If
            Case "mab":
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                cmd$ = meineabrede("-1." & trm(rv!felddaten))
                Print #p, form1.repl1310rtf(cmd$);: Call dbgu(cmd$)
              End If
            Case "plzort":
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                felddaten$ = trm(rv!felddaten)
                sida$ = felddaten$
                If InStr(felddaten$, "{") > 0 Then
                  sid$ = felddaten$
                  sidp% = InStr(sid$, "{")
                  sida$ = sid$
                  sidk$ = trm(Left(sid$, sidp% - 1))
                  sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
                  sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
                  cmd$ = "select * from kontakt where id='" & sidk$ & "'"
                  Set stmp = New ADODB.Recordset
                  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                  If Not stmp.EOF Then
                    If trm(stmp!lkz) & trm(stmp!plz) & trm(stmp!ort) <> "" Then
                      xa_erg$ = getplzort(trm(stmp!lkz), trm(stmp!plz), trm(stmp!ort))
                      If xa_erg$ <> "" Then
                        Print #p, form1.repl1310rtf(xa_erg$);: Call dbgu(xa_erg$)
                        sida$ = ""
                      End If
                    End If
                  End If
                Else
                  sida$ = trm(rv!felddaten)
                End If
                If sida$ <> "" Then
                  cmd$ = "select * from adresse where id='" & sida$ & "'"
                  Set stmp = New ADODB.Recordset
                  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                  If Not stmp.EOF Then
                    xa_erg$ = getplzort(trm(stmp!land), trm(stmp!plz), trm(stmp!ort))
                    Print #p, form1.repl1310rtf(xa_erg$);: Call dbgu(xa_erg$)
                  End If
                End If
              End If
            Case "xa":
              xa_adr$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_adr$ = getaliasfeld(xa_adr$)
              xa_typ$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
'kein alias beim typ
'              xa_typ$ = getaliasfeld(xa_typ$)
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_adr$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & trm(rv!felddaten) & "' and auftrittstyp='" & xa_typ$ & "' and instr(lcase(feldname),'" & xa_feld$ & "')=1"
                Set stmp = New ADODB.Recordset
                stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not stmp.EOF Then
                  Print #p, form1.repl1310rtf(trm(stmp!felddaten));: Call dbgu(trm(stmp!felddaten))
                End If
              End If
            Case "adr":
              xa_typ$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and feldname='" & xa_feld$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                Print #p, form1.repl1310rtf(trm(rv!felddaten));: Call dbgu(trm(rv!felddaten))
              End If
            Case "xk":
              xa_adr$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_kont$ = Left$(pa$, InStr(pa$, "__") - 1): pa$ = Mid$(pa$, InStr(pa$, "__") + 2)
              xa_feld$ = pa$
              cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_kont$ & "'"
              Set rv = New ADODB.Recordset
              rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not rv.EOF Then
                xa_knam$ = trm(rv!felddaten)
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_adr$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  xa_feld$ = ratefeldaustabelle("kontakt", xa_feld$)
                  cmd$ = "select " & xa_feld$ & " as felddaten from kontakt where vid='" & trm(rv!felddaten) & "' and name like '" & xa_knam$ & "%'"
                  Set stmp = New ADODB.Recordset
                  stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                  If rrr = 0 Then
                    If Not stmp.EOF Then
                      Print #p, form1.repl1310rtf(trm(stmp!felddaten));: Call dbgu(trm(stmp!felddaten))
                    End If
                  Else
                    GoTo versuchkontakt
                  End If
                End If
              Else
versuchkontakt:
                rv.Close
                cmd$ = "select felddaten from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" & rtmp!auftrittstyp & "' and feldname='" & xa_adr$ & "'"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not rv.EOF Then
                  felddaten$ = trm(rv!felddaten)
                  sidk$ = ""
                  If InStr(felddaten$, "{") > 0 Then
                    sid$ = felddaten$
                    sidp% = InStr(sid$, "{")
                    sida$ = sid$
                    sidk$ = trm(Left(sid$, sidp% - 1))
                    sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
                    sidk$ = form1.get_kontaktid_by_name(sida$, sidk$)
                  Else
                    sida$ = felddaten$
                  End If
                  rv.Close
                  cmd$ = "select felddaten from auftritthigru where auftrittsid='" & sida$ & sidk$ & "' and auftrittstyp='" & xa_kont$ & "' and instr(feldname,'" & xa_feld$ & "')=1"
                  Set rv = New ADODB.Recordset
                  rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                  If Not rv.EOF Then
                    Print #p, form1.repl1310rtf(trm(rv!felddaten));: Call dbgu(trm(rv!felddaten))
                  End If
                  rv.Close
                End If
              End If
            Case "inworten":
              pax$ = ""
              If LCase(pa$) = "bruttohonorar" Then
                pa$ = inworten(CLng(num1("0" & trm(bruttohonorar))))
              Else
                If LCase(pa$) = "current" Then
                  pa$ = iwcurrent$
                Else
                  If InStr(LCase(pa$), "merkslot") = 1 Then
                    On Error Resume Next
                    pa$ = adruckmerkwert(Val(Mid$(pa$, 9)))
                    rrr = Err
                    On Error GoTo 0
                    If rrr <> 0 Then
                      pa$ = "FEHLER"
                    Else
                      pa$ = inworten(CLng(num1("0" & trm(pa$))))
                    End If
                  Else
                    Set a = New ADODB.Recordset
                    a.CursorLocation = adUseServer
rrr = form1.adoopen(a, "SELECT " + pa$ + " FROM usr_" & utabn(rtmp!auftrittstyp) & " where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                    If Not a.EOF Then
                      If Not IsNull(a.Fields(0).value) Then
                        ava = a.Fields(0).value
                        pa$ = inworten(CLng(num1("0" & trm(ava))))
                      Else
                        pa$ = ""
                      End If
                    End If
                  End If
                End If
              End If
              Call adpprint(adruckmerkslot, p%, pa$): Call dbgu(pa$)
            Case "inwords":
              pax$ = ""
              If LCase(pa$) = "bruttohonorar" Then
                pa$ = inwords(CLng(num1("0" & trm(bruttohonorar))))
              Else
                Set a = New ADODB.Recordset
                a.CursorLocation = adUseServer
rrr = form1.adoopen(a, "SELECT " + pa$ + " FROM usr_" & utabn(rtmp!auftrittstyp) & " where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If Not a.EOF Then
                  If Not IsNull(a.Fields(0).value) Then
                    ava = a.Fields(0).value
                    pa$ = inwords(CLng(num1("0" & trm(ava))))
                  Else
                    pa$ = ""
                  End If
                End If
              End If
              Call adpprint(adruckmerkslot, p%, pa$): Call dbgu(pa$)
            Case Default:
          End Select
        End If
      End If
      If ttest$ = "user" Then
        If Not udat.EOF Then
          For i% = 0 To 21 ' see einstellungen.load
            If Len(udat.Fields(i%).name) = Len(rev$) - 1 Then   'aliase ermitteln
              If isdigit(Right$(rev$, 1)) <> 0 Then rev$ = Left$(rev$, Len(rev$) - 1)
            End If
            If LCase(udat.Fields(i%).name) = LCase(rev$) Then
              Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(strrepl(trm(udat.Fields(i%).value), "\", "\\"))): Call dbgu(strrepl(trm(udat.Fields(i%).value), "\", "\\"))
              i% = 999
            End If
          Next i%
          If i% < 999 Then
            Call adpprint(adruckmerkslot, p%, form1.getusersetting(rev$, "")): Call dbgu(form1.getusersetting(rev$, ""))
          End If
        End If
      End If
      If ttest$ = "system" Then
        Select Case LCase(rev$)
          Case "numdat": dat$ = datum2sql(Date): dat$ = Mid$(dat$, 3, 2) & Mid$(dat$, 6, 2) & Mid$(dat$, 9, 2)
                         Call adpprint(adruckmerkslot, p%, dat$): Call dbgu(dat$)
          Case "datummonatlang":
                        c$ = Left(trm(Date), 2) + ". " + transe(mnams$(Val(Mid(trm(Date), 4, 2)))) + " " + Right(trm(Date), 4)
                        Call adpprint(adruckmerkslot, p%, c$)
                        Call dbgu(c$)
          Case "datum": c$ = datum2sql(trm(Date))
                        c$ = Mid$(c$, 9, 2) & "." & Mid$(c$, 6, 2) & "." & Mid$(c$, 1, 4)
                        Call adpprint(adruckmerkslot, p%, c$)
                        Call dbgu(c$)
          Case "saison":
               Call adpprint(adruckmerkslot, p%, saison(trm(Date))): Call dbgu(Date)
          Case "datumjahr2": Call adpprint(adruckmerkslot, p%, Right(trm(Date), 2)): Call dbgu(Date)
          Case "datumjahr": Call adpprint(adruckmerkslot, p%, Right(trm(Date), 4)): Call dbgu(Date)
          Case "datummonat": Call adpprint(adruckmerkslot, p%, Mid(trm(Date), 3, 2)): Call dbgu(Date)
          Case "datumtag": Call adpprint(adruckmerkslot, p%, Left(trm(Date), 2)): Call dbgu(Date)
          Case "zeit": Call adpprint(adruckmerkslot, p%, Left(Time, 5)): Call dbgu(Left(Time, 5))
          Case "mwst": Call adpprint(adruckmerkslot, p%, fixeurnozerotail(thismwst / 100)): Call dbgu(fixeurnozerotail(thismwst / 100))
          Case "rechnr": Call adpprint(adruckmerkslot, p%, new_rechnr(nfn$, rechbez$)): Call dbgu("Rechnr=" + getsystemsetting("RechNr", ""))
          Case Else: Call adpprint(adruckmerkslot, p%, form1.getsystemsetting(rev$)): Call dbgu(form1.getsystemsetting(rev$))
        End Select
      End If
      If ttest$ = "auftritte" Or InStr(ttest$, "auftritteif") = 1 Then
        onlyif$ = ""
        If ttest$ <> "auftritte" Then
          onlyif$ = Mid$(ttest$, 12)
        End If
        Print #p%, "#includelist#" & trm(listenzaehler%): Call dbgu("#includelist#" & trm(listenzaehler%))
        If wmode$ = "auftritt" Then
          Load tplan
          Call tplan.gotorec(tpid$)
          For al% = 0 To tplan.List6.ListCount - 1
            tplan.List6.ListIndex = al%
            DoEvents
            al_id$ = tplan.List6.List(tplan.List6.ListIndex)
            al_id$ = Mid$(al_id$, InStr(al_id$, "(AID:") + 5)
            xtake = True
            If onlyif$ <> "" Then
              xtake = False
              c$ = "SELECT FeldDaten FROM auftritthigru where auftrittsid='" + al_id$ + "' and FeldName='" + onlyif$ + "'"
              Set al_r = New ADODB.Recordset
              al_r.CursorLocation = adUseServer
              rrr = form1.adoopen(al_r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not al_r.EOF Then
                If trm0(al_r!felddaten) <> "" Then
                  xtake = True
                End If
              End If
            End If
            Set al_r = New ADODB.Recordset
            al_r.CursorLocation = adUseServer
            c$ = "SELECT * FROM auftritt where id='" + al_id$ + "'"
            rrr = form1.adoopen(al_r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If xtake And Not al_r.EOF Then
              al_vorlage$ = vorlagendir & "\" & al_r!auftrittstyp & "_" & rev$
              rtf_vorlage$ = al_vorlage$ & ".rtf"
              al_vorlage$ = al_vorlage$ & ".txt"
              If exist(rtf_vorlage$) = 0 Then
                rtf_vorlage$ = vorlagendir & "\" & al_r!auftrittstyp & "_zsys" & rev$ & ".rtf"
              End If
              If needrecomptxt(rtf_vorlage$, al_vorlage$) Or form1.getusersetting("deletetxtcache", "nein") = "ja" Then
                If Not nexist(al_vorlage$) Then
                  On Error Resume Next
                  Kill al_vorlage$
                  On Error GoTo 0
                End If
              End If
              If exist(rtf_vorlage$) > 0 And exist(al_vorlage$) = 0 Then
              rvl% = FreeFile
              Open rtf_vorlage$ For Input As #rvl%
              avl% = FreeFile
              Open al_vorlage$ For Output As #avl%
              cpm% = 0: brkcpm% = 0
              While Not EOF(rvl%) And brkcpm% = 0
                Line Input #rvl%, lx$
                px% = InStr(lx$, "Liste__startet__hier")
                If px% > 0 Then
                  cpm% = 1
                  pardepress = True
                  lx$ = Mid$(lx$, px% + Len("Liste__startet__hier") + 1)
                End If
                px% = InStr(lx$, "Liste__endet__hier")
                If px% > 0 Then
                  cpm% = 0
                  brkcpm% = 1
                  If px% > 1 Then
                    lx$ = Left$(lx$, px% - 1)
                    Print #avl%, lx$
                  End If
                End If
                If cpm% = 1 And Len(lx$) > 0 Then
                  If pardepress Then
                    pardepress = False
                    If Left(lx$, 5) = "\par " Then lx$ = Mid(lx$, 6)
                  End If
                  Print #avl%, lx$
                End If
              Wend
              Close #avl%
              Close #rvl%
              If brkcpm% = 0 Then
                dbupgrade.List1.AddItem transe("Keine Listenbegrenzung in") + " "
                dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
                seriouswarning = True
              End If
              End If
              tr$ = Dir(al_vorlage$)
              If tr$ <> "" Then
                lx$ = "": If trm(l4adr$) <> "" Then lx$ = "|" + l4adr$
                todo(listenzaehler%).AddItem al_id$ & ":" & al_vorlage$ + lx$
                ctmp$ = ProvisionFuerAuftritt(al_id$)
                Honorarliste$(3, honorarlcount%) = ctmp$
                Honorarliste$(4, honorarlcount%) = MwStFuerAuftritt(al_id$)
                Honorarliste$(5, honorarlcount%) = al_r!datum
                Honorarliste$(6, honorarlcount%) = trm(thismwst)
                If honorarlcount% > 0 And Left$(ctmp$, 1) = "0" Then
                  Honorarliste$(6, honorarlcount%) = Honorarliste$(6, honorarlcount% - 1)
                  thismwst = Honorarliste$(6, honorarlcount%)
                End If
                If AuftrittsdruckFuerAdresse$ <> "" And LCase(auftrittstyp(id$)) = "orchesterauftritt" Then
                  Honorarliste$(0, honorarlcount%) = auftrittshonorarfeldbyname(id$, AuftrittsdruckFuerAdresse$)
                Else
                  Honorarliste$(0, honorarlcount%) = HonorarVonAuftritt(al_id$)
                End If
                Honorarliste$(1, honorarlcount%) = ""
'lne$ = honorarlcount% & " <" & Honorarliste$(0, honorarlcount%) & "> " & al_id$
                Call inchonorarlcount
              End If
            End If
          Next al%

        Else   ' adresse includet termine

          For al% = 0 To shwAdrDetail.List2.ListCount - 1
            shwAdrDetail.List2.ListIndex = al%
            DoEvents
            al_id$ = shwAdrDetail.List2.List(shwAdrDetail.List2.ListIndex)
            al_id$ = Mid$(al_id$, InStr(al_id$, "(AID:") + 5)
            Set al_r = New ADODB.Recordset
            al_r.CursorLocation = adUseServer
rrr = form1.adoopen(al_r, "SELECT * FROM auftritt where id='" + al_id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not al_r.EOF Then
              al_vorlage$ = vorlagendir & "\" & al_r!auftrittstyp & "_" & rev$
              If nexist(al_vorlage$ + ".rtf") Then al_vorlage$ = vorlagendir & "\" & al_r!auftrittstyp & "_zsys" & rev$
              rtf_vorlage$ = al_vorlage$ & ".rtf"
              al_vorlage$ = al_vorlage$ & ".txt"
              If needrecomptxt(rtf_vorlage$, al_vorlage$) Or form1.getusersetting("deletetxtcache", "nein") = "ja" Then
                If Not nexist(al_vorlage$) Then
                  On Error Resume Next
                  Kill al_vorlage$
                  On Error GoTo 0
                End If
              End If
              If exist(rtf_vorlage$) > 0 And exist(al_vorlage$) = 0 Then
              rvl% = FreeFile
              Open rtf_vorlage$ For Input As #rvl%
              avl% = FreeFile
              Open al_vorlage$ For Output As #avl%
              cpm% = 0: brkcpm% = 0
              While Not EOF(rvl%) And brkcpm% = 0
                Line Input #rvl%, lx$
                px% = InStr(lx$, "Liste__startet__hier")
                If px% > 0 Then
                  cpm% = 1
                  lx$ = Mid$(lx$, px% + Len("Liste__startet__hier") + 1)
                End If
                px% = InStr(lx$, "Liste__endet__hier")
                If px% > 0 Then
                  cpm% = 0
                  brkcpm% = 1
                  If px% > 1 Then
                    lx$ = Left$(lx$, px% - 1)
                    Print #avl%, lx$
                  End If
                End If
                If cpm% = 1 And Len(lx$) > 0 Then Print #avl%, lx$
              Wend
              Close #avl%
              Close #rvl%
            End If
            tr$ = Dir(al_vorlage$)
            If tr$ <> "" Then
                lx$ = "": If trm(l4adr$) <> "" Then lx$ = "|" + l4adr$
                todo(listenzaehler%).AddItem al_id$ & ":" & al_vorlage$ + lx$
                ctmp$ = ProvisionFuerAuftrittByAdr(al_id$, adr_id$)
                Honorarliste$(3, honorarlcount%) = ctmp$
                Honorarliste$(4, honorarlcount%) = MwStFuerAuftritt(al_id$)
                Honorarliste$(5, honorarlcount%) = al_r!datum
                Honorarliste$(6, honorarlcount%) = trm(thismwst)
                If honorarlcount% > 0 And Left$(ctmp$, 1) = "0" Then
                  Honorarliste$(6, honorarlcount%) = Honorarliste$(6, honorarlcount% - 1)
                  thismwst = Honorarliste$(6, honorarlcount%)
                End If
                If InStr(Honorarliste$(3, honorarlcount%), "/") > 0 Then
                  Honorarliste$(6, honorarlcount%) = cut_d2bis(Honorarliste$(3, honorarlcount%), "/")
                  Honorarliste$(3, honorarlcount%) = cut_d1(Honorarliste$(3, honorarlcount%), "/")
                End If
                If AuftrittsdruckFuerAdresse$ <> "" And LCase(auftrittstyp(al_id$)) = "orchesterauftritt" Then
                  Honorarliste$(0, honorarlcount%) = auftrittshonorarbyname(al_id$, AuftrittsdruckFuerAdresse$)
                Else
                  Honorarliste$(0, honorarlcount%) = HonorarVonAuftrittByAdr(al_id$, adr_id$)
                End If
                Honorarliste$(1, honorarlcount%) = ""
                Call inchonorarlcount
              End If
            End If
          Next al%

        End If
      listenzaehler% = listenzaehler% + 1
      End If
      DoEvents
      If pacont Then ttest$ = xa_feld$
      For bkms% = 0 To 1
        For bk% = 0 To bkmlcount%(bkms%)
          If LCase(bkmlist(bkms%, bk%)) = LCase(ttest$) Then
            If bkms% = 0 Then
              If ttest$ = "datum" Then
                If ttest$ = "datum" Then
                  Call adpprint(adruckmerkslot, p%, printdatfromsql(rtmp.Fields(bk%).value)): Call dbgu(datfromsql(rtmp.Fields(bk%).value))
                End If
                datumfuerkurs$ = rtmp.Fields(bk%).value
              Else
                Select Case ttest$
                  Case "terminende"
                     Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(form1.auftrittsende(id$, ""))): Call dbgu(form1.auftrittsende(id$, ""))
                  Case "programmid"
                     Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(programmid$)): Call dbgu(programmid$)
                  Case "auftrittstyp"
                     Call adpprint(adruckmerkslot, p%, trm(repl1310rtf(transe(rtmp.Fields(bk%).value)))): Call dbgu(transe(rtmp.Fields(bk%).value))
                  Case "mwst"
                     Call adpprint(adruckmerkslot, p%, trm(repl1310rtf(fixeurnozerotail("0" & trm(Abs(amwst%) / 100))))): Call dbgu(trm(fixeurnozerotail("0" & trm(Abs(amwst%) / 100))))
                  Case "mwst0text"
                     If amwst% = 0 Then
                       c$ = getusersetting("mwst0text", "")
                       If c$ <> "" Then
                         Call adpprint(adruckmerkslot, p%, trm(repl1310rtf(c$))): Call dbgu("mwst0text=" + c$)
                       End If
                     End If
                  Case "bruttohonorarwaehrung"
                    Call adpprint(adruckmerkslot, p%, bruttohonorarwaehrung): Call dbgu(bruttohonorarwaehrung)
                  Case "provisionaufbrutto_netto"
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(provisionaufbrutto_netto)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(provisionaufbrutto_netto)) & " " & bruttohonorarwaehrung)
                  Case "provisionaufhonorarva_netto"
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(HonorarVA(id$) * provision)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(provisionaufbrutto_netto)) & " " & bruttohonorarwaehrung)
                  Case "provisionaufbrutto_mwst"
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(provisionaufbrutto_mwst)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(provisionaufbrutto_mwst)) & " " & bruttohonorarwaehrung)
                  Case "provisionaufbrutto_brutto"
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(provisionaufbrutto_brutto)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(provisionaufbrutto_brutto)) & " " & bruttohonorarwaehrung)
                  Case "bruttohonabzglprovbrutto"
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(bruttohonabzglprovbrutto)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(bruttohonabzglprovbrutto)) & " " & bruttohonorarwaehrung)
                  Case ""
                    Call adpprint(adruckmerkslot, p%, trm(fixeur(provisionaufbrutto_brutto)) & " " & bruttohonorarwaehrung): Call dbgu(trm(fixeur(provisionaufbrutto_brutto)) & " " & bruttohonorarwaehrung)
                  Case "meinzeichen"
                     Call adpprint(adruckmerkslot, p%, meinzeichen$): Call dbgu(meinzeichen$)
                  Case "wochentag"
                     Call adpprint(adruckmerkslot, p%, wochentag$): Call dbgu(wochentag$)
                  Case "langwochentag"
                     Call adpprint(adruckmerkslot, p%, langwochentag$): Call dbgu(langwochentag$)
                  Case "saison"
                     Call adpprint(adruckmerkslot, p%, auftrittsaison$): Call dbgu(auftrittsaison$)
                  Case "engwochentag"
                     Call adpprint(adruckmerkslot, p%, engwochentag$): Call dbgu(engwochentag$)
                  Case "englangwochentag"
                     Call adpprint(adruckmerkslot, p%, englangwochentag$): Call dbgu(englangwochentag$)
                  Case "nettohonorar"
                     Call adpprint(adruckmerkslot, p%, fixeur(nettohonorar) & " " & bruttohonorarwaehrung): Call dbgu(nettohonorar)
                  Case "mwsthonorar"
                     Call adpprint(adruckmerkslot, p%, fixeur(mwsthonorar) & " " & bruttohonorarwaehrung): Call dbgu(mwsthonorar)
                  Case "tarahonorar"
                     Call adpprint(adruckmerkslot, p%, fixeur(tarahonorar) & " " & bruttohonorarwaehrung): Call dbgu(mwsthonorar)
                  Case "bruttohonorar"
                     Call adpprint(adruckmerkslot, p%, fixeur(bruttohonorar) & " " & bruttohonorarwaehrung): Call dbgu(bruttohonorar)
                  Case "bruttohonorarumrechnung"
                     Call adpprint(adruckmerkslot, p%, repl1310rtf(bhfremd)): Call dbgu(bhfremd)
                  Case "bruttohonorarxumrechnung"
                     Call adpprint(adruckmerkslot, p%, repl1310rtf(trm(xkurs_publicratestring))): Call dbgu(trm(xkurs_publicratestring))
                  Case "bruttohonorarumrechnungcrlf"
                     Call adpprint(adruckmerkslot, p%, repl1310rtf(vbCrLf & bhfremd)): Call dbgu(bhfremd)
                  Case "astatus"
                    If Not IsNull(rtmp.Fields(bk%).value) Then
                      Call adpprint(adruckmerkslot, p%, get_eventstatusname(rtmp.Fields(bk%).value)): Call dbgu(get_eventstatusname(rtmp.Fields(bk%).value))
                    Else
                      Call adpprint(adruckmerkslot, p%, get_eventstatusname(-1)): Call dbgu(get_eventstatusname(-1))
                    End If
                  Case Else
                    On Error Resume Next
                    If ttest$ = "zeit" Then
                      c$ = trm(rtmp.Fields(bk%).value)
                      Call adpprint(adruckmerkslot, p%, Left(c$, 5)): Call dbgu(datfromsql(rtmp.Fields(bk%).value))
                    Else
                      If Not IsNull(rtmp.Fields(bk%).value) Then
                        Call adpprint(adruckmerkslot, p%, rtmp.Fields(bk%).value): Call dbgu(rtmp.Fields(bk%).value)
                      End If
                    End If
                    On Error GoTo 0
                End Select
              End If
              bk% = bkmlcount%(bkms%)
              bkms% = 1
            Else
              If Not islistfeld(LCase(rtmp!auftrittstyp), bkmlist(bkms%, bk%)) Or InStr(bkmlist(bkms%, bk%), "merkslot") = 1 Then
              Set a = New ADODB.Recordset
              a.CursorLocation = adUseServer
rrr = form1.adoopen(a, "SELECT * FROM usr_" & utabn(rtmp!auftrittstyp) & " where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not a.EOF Then
                felddaten$ = ""
                If InStr(bkmlist(bkms%, bk%), "merkslot") = 1 Then
                  felddaten = adruckmerkwert(Val(Mid$(bkmlist(bkms%, bk%), 9)))
                Else
                  On Error Resume Next
                  felddaten$ = trm(a.Fields(bkmlist(bkms%, bk%)).value)
                  rrr = Err
                  On Error GoTo 0
                  If rrr <> 0 Then felddaten$ = ""
                  If LCase(bkmlist(bkms%, bk%)) = "provision" Then
                    felddaten$ = trm(cut_d1(felddaten$, "/"))
                  End If
                  If InStr(LCase(bkmlist(bkms%, bk%)), "datum") > 0 Then
                    If Len(felddaten$) < 11 Then
                      felddaten$ = datum2sql(felddaten$)
                      felddaten$ = Mid$(felddaten$, 9, 2) & "." & Mid$(felddaten$, 6, 2) & "." & Mid$(felddaten$, 1, 4)
                    End If
                  End If
                End If
                w$ = ""
                If bkmrflag$(bk%) = "" Or bkmrflag$(bk%) = "finanzen" Then
                  ' möglichkeit f. kurs hinter dem Betrag
                  w$ = felddaten$: w$ = cut_d1(w$, ":::")
                  If InStr(LCase(bkmlist(bkms%, bk%)), "provision") = 1 Then
                    If l4adr$ = "" Then w$ = ProvisionFuerAuftrittByAdr(id$, l4adr$)
                    w$ = trm0(cut_d1(w$, "/"))
                  End If
                  If InStr(LCase(bkmlist(bkms%, bk%)), "honorar") = 1 Or InStr(LCase(bkmlist(bkms%, bk%)), "gesamtpreis") = 1 Then
                    If AuftrittsdruckFuerAdresse$ <> "" And LCase(auftrittstyp(id$)) = "orchesterauftritt" Then
                      Honorarliste$(0, honorarlcount%) = auftrittshonorarbyname(id$, AuftrittsdruckFuerAdresse$)
                      w$ = Honorarliste$(0, honorarlcount%)
                    Else
                      If l4adr$ <> "" Then
                        Honorarliste$(0, honorarlcount%) = felddaten$
                      Else
                        Honorarliste$(0, honorarlcount%) = HonorarVonAuftrittByAdr(id$, "")
                        w$ = Honorarliste$(0, honorarlcount%)
                      End If
                    End If
                    Honorarliste$(1, honorarlcount%) = ""
                    Call inchonorarlcount
                  End If
                  Print #p, form1.repl1310rtf("" & w$);: Call dbgu(w$)
                  If InStr(LCase(bkmlist(bkms%, bk%)), "provision") > 0 Then
                    honvalid = 0
                    provtyp = 0
                    If l4adr$ = "" Then felddaten$ = ProvisionFuerAuftrittByAdr(id$, l4adr$)
                    felddaten$ = strrepl(felddaten$, Chr$(10), " ")
                    felddaten$ = strrepl(felddaten$, Chr$(13), " ")
                    If trm(felddaten$) <> "" Then
                      If InStr(felddaten$, "/") > 0 Then
                        thismwst = var2dbl(trm(cut_d1(trm(cut_d2bis(felddaten$, "/")), "%"))) * 100
                        felddaten$ = trm(cut_d1(felddaten$, "/"))
                      End If
                      'If l4adr$ = "" Then felddaten$ = ProvisionFuerAuftrittByAdr(id$, l4adr$)
                      If InStr(felddaten$, "%") > 0 Then
                        provtyp = 1
                        pprov = cut_d1(felddaten$, "%")
                      Else
                        pprov = ohnewaehrung(felddaten$)
                      End If
                      On Error Resume Next
                      provision = CCur(pprov)
                      rrr = Err
                      On Error GoTo 0
                      If rrr <> 0 Then provision = 0
                      If provtyp = 1 Then
                        provision = provision / 100
                        provisionaufbrutto_netto = var2dbl(CCur(word1("" & bruttohonorar))) * provision
                      Else
                        provisionaufbrutto_netto = provision / hK
                      End If
                      provisionaufbrutto_mwst = thismwst * provisionaufbrutto_netto / 10000
                      provisionaufbrutto_brutto = provisionaufbrutto_netto + provisionaufbrutto_mwst
                      On Error Resume Next
                      bruttohonabzglprovbrutto = var2dbl(CCur(word1("" & bruttohonorar))) - provisionaufbrutto_brutto
                      rrr = Err
                      On Error GoTo 0
                      If rrr <> 0 Then bruttohonabzglprovbrutto = 0
                    End If
                  End If
                Else
                  If LCase(bkmrflag$(bk%)) = "vertragsnummer" And Not palstd Then
                    palstd = True
                    c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" & _
                          form1.newid("dochist", "id", 18) & "','system','-1','" & nfn$ & "','" & _
                          datum2sql(Date) & " " & Time & "','" & form1.getuserid() + "','" & Right(FileName(vorlage$), 39) + "','Vertragsnummer: " + trm(felddaten$) + "','Vertragsnummer " + trm(felddaten$) + "')"
                    Call form1.sqlqry(c$)
                  End If
                  If bkmrflag$(bk%) = "programm" Or bkmrflag$(bk%) = "programmsbz" Then
    Set rprog = New ADODB.Recordset
    rprog.CursorLocation = adUseServer
rrr = form1.adoopen(rprog, "SELECT werkid FROM programmliste where programmid='" + felddaten$ + "' order by position", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rprog.EOF Then
      If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then Print #p%, "\par"
      prvk$ = ""
      strt = True
    While Not rprog.EOF
      werkid$ = trm(rprog!werkid): sbzid$ = ""
      If Left$(werkid$, 4) = "SBZ:" Then
        sbzid$ = Mid$(werkid$, 5)
        werkid$ = form1.getsatzidbywerkid(sbzid$)
      End If
      If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
        k$ = "\trowd \trgaph70\trleft-70 \cellx2410\cellx9476 \pard \intbl "
        If Not nexist(vorlage$ + "." + rev$) Then
          o0% = FreeFile
          Open vorlage$ + "." + rev$ For Input As #o0%
          Line Input #o0%, k$
          Close #o0%
        End If
        Print #p%, k$
      End If
      k$ = getkompvornamenamebywerkid(werkid$)
      kkurz1$ = getkompnachnamebywerkid(werkid$)
      dau$ = getdauerbywerkid(werkid$): If sbzid$ <> "" Then dau$ = ""
      ston$ = getstimmtonbywerkid(werkid$)
      d$ = getkompdatesbywerkid(werkid$)
      If d$ <> "" Then d$ = "(" + d$ + ")"
      If (rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1") Or (Left$(LCase$(k$), 7) <> "pause p" And Left$(LCase$(k$), 7) <> "oder od") Then
        If Left$(LCase$(k$), 7) = "pause p" Or Left$(LCase$(k$), 7) = "oder od" Then
          k$ = ""
          d$ = ""
        End If
        If (k$ & d$) <> prvk$ Or (rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1") Then
          If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
            Print #p%, "{";
          Else
            If strt Then
              strt = False
            Else
              Print #p%, " \par ";
            End If
          End If
          Select Case rev$
            Case "kurz1":
              Call adpprint(adruckmerkslot, p%, form1.repl1310rtf("" & kkurz1$ & "")): Call dbgu(kkurz1$)
            Case Else:
              Call adpprint(adruckmerkslot, p%, form1.repl1310rtf("" & k$ & "")): Call dbgu(k$)
          End Select
          If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
            Print #p%, "\par ";
          Else
            Print #p%, " ";
          End If
          If rev$ <> "kurz1" Then Call adpprint(adruckmerkslot, p%, d$): Call dbgu(d$)
        End If
        prvk$ = k$ & d$
        If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
          Print #p, "\cell ";
        Else
          If LCase(k$) <> "pause" And LCase(k$) <> "oder" Then
            Print #p%, "\par   ";
          End If
        End If
        If sbzid$ = "" Then
          Select Case rev$
            Case "kurz1":
              dbguo$ = getwerkopusnamebyid(werkid$): Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(dbguo$)): Call dbgu(dbguo$)
            Case Else:
              If LCase(k$) <> "pause" And LCase(k$) <> "oder" Then
                dbguo$ = getwerknamebyid(werkid$): Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(dbguo$)): Call dbgu(dbguo$)
              End If
          End Select
        Else
          Call adpprint(adruckmerkslot, p%, form1.repl1310rtf(getsatznamebyid(sbzid$) + " " + transe("aus") + " " + getwerknamebyid(werkid$))): Call dbgu(getwerknamebyid(werkid$))
        End If
        If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
          If trm(dau$) <> "" Then Call adpprint(adruckmerkslot, p%, " (" + form1.repl1310rtf(dau$ & "Min.") + ") ")
          If trm(ston$) <> "" Then Call adpprint(adruckmerkslot, p%, " (" + form1.repl1310rtf("Stimmton: " + ston$) + ") ")
        End If
        If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then Print #p%, "\par "
        If rev$ = "sbz" And sbzid$ = "" Then
'mit Satzbezeichnungen?
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT satzbezeichnung FROM sbz_loc where wid='" + rprog!werkid + "' order by satznummer", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not stmp.EOF
            Call adpprint(adruckmerkslot, p%, form1.repl1310rtf("\tab " + stmp!satzbezeichnung + "\par ")): Call dbgu("" & stmp!satzbezeichnung & "")
            stmp.MoveNext
          Wend
        End If
        If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then Print #p%, "\cell \pard \intbl \row }";
        If rev$ <> "text" And rev$ <> "textohnesbz" And rev$ <> "kurz1" Then
          Print #p%, "\pard";
        End If
      End If
      rprog.MoveNext
    Wend
    Else
      If LCase(Left$(felddaten$, 6)) <> "datei:" And LCase(Left$(felddaten$, 3)) <> "fn:" Then
        Call adpprint(adruckmerkslot, p%, felddaten$): Call dbgu(felddaten$)
      Else
        rvp% = InStr(felddaten$, ":") + 1
        lprg$ = trm(Mid$(felddaten$, rvp%))
        If Not nexist(lprg$) Then
          If LCase(FileExtension(lprg$)) = "txt" Then
            prgo% = FreeFile
            Open trm(Mid$(felddaten$, rvp%)) For Input As prgo%
            While Not EOF(prgo%)
              Line Input #prgo%, lprg$
              Call adpprint(adruckmerkslot, p%, lprg$): Print #p%, "\par "
            Wend
            Close #prgo%
          Else
          End If
        Else
        End If
      End If
    End If
                  Else
                    sidk$ = ""
                    sidkfeld$ = ""
                    usesidk = False
                    'kontakt statt adresse
                    If InStr(felddaten$, "{") > 0 Then
                      sid$ = felddaten$
                      sidp% = InStr(sid$, "{")
                      sida$ = sid$
                      sidk$ = trm(Left(sid$, sidp% - 1))
                      sida$ = trm(Mid(sid$, sidp% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
                      sidkfeld$ = strrepl(strrepl(rev$, "__", ","), "land", "lkz")
                      If sidkfeld$ = "" Then sidkfeld$ = "name"
                      cmd$ = "select id," & sidkfeld$ & " from kontakt where instr(name,'" & sidk$ & "')=1 and vid='" & sida$ & "'"
                      'Adressdaten der Kontaktes ermitteln
                      If LCase(sidkfeld$) = "name" Then
                        'der name etwas anders als die adresse:
                        sidkfeld$ = "select name from " & bkmrflag$(bk%) & " where id='" & sida$ & "'"
                        Set rv = New ADODB.Recordset
                        rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, sidkfeld$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                        If Not rv.EOF Then
                          sidkfeld$ = trm(rv!name)
                          usesidk = True
                        Else
                          sidkfeld$ = ""
                        End If
                        rv.Close
                      Else
                        'andere daten der adresse
                        If InStr(sidkfeld$, ",") = 0 Then
                          If sidkfeld$ = "lkz" Then sidkfeld$ = "land"
                          sidkfeld$ = "select id," & strrepl(rev$, "__", ",") & " as feweadr from " & bkmrflag$(bk%) & " where id='" & sida$ & "'"
                          Set rv = New ADODB.Recordset
                          rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, sidkfeld$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                          If Not rv.EOF Then
                            sidkfeld$ = trm(rv!feweadr)
                          Else
                            sidkfeld$ = ""
                          End If
                          rv.Close
                        End If
                      End If
                    Else
                      cmd$ = "select id"
                      If rev$ <> "" Then
                        cmd$ = cmd$ & ","
                        cmd$ = cmd$ & strrepl(rev$, "__", ",")
                      End If
                      cmd$ = cmd$ & " from " & bkmrflag$(bk%) & " where id='" & felddaten$ & "'"
                    End If
                    'Print #p, cmd$; "\par"
                    Set rv = New ADODB.Recordset
                    rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                    If rrr = 0 Then
                    If Not rv.EOF Then
                      If Not dochistlock Then
                        If sidk$ = "" Then
                          cmx$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" + form1.newid("dochist", "id", 19) & "','" & felddaten$ & "','" & "-1" & "','" & nfn$ & "','" & docdtg$ & "','" & uId$ & "','..." + Right$(vorlage$, 18) + "','" + Right(FileName(vorlage$), 39) + "')"
                        Else
                          sidkid$ = trm(rv!id)
                          cmx$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" + form1.newid("dochist", "id", 19) & "','" & sida$ & "','" & sidkid$ & "','" & nfn$ & "','" & docdtg$ & "','" & uId$ & "','..." + Right$(vorlage$, 18) + "','" + Right(FileName(vorlage$), 39) + "')"
                        End If
                      End If
                      dochistlist.AddItem cmx$
                      For fc% = 1 To rv.Fields.Count - 1
'tel und fax nicht mehr automatsch ausschreiben
'                        Select Case LCase(rv.Fields(fc%).Name)
'                          Case "tel": Print #p%, UcaseFirstLetter(rv.Fields(fc%).Name); ": ";
'                          Case "fax": Print #p%, UcaseFirstLetter(rv.Fields(fc%).Name); ": ";
'                          Case Else:
'                        End Select
                        If Not IsNull(rv.Fields(fc%).value) Then
                          rvf$ = rv.Fields(fc%).value
                          If sidkfeld$ <> "" Then
                            If usesidk Then
                              If UCase(orev$) <> orev$ Then
                                If adrnameonly Then
                                  If InStr(felddaten$, "{") > 0 Then
'                                      rvf$ = trm(cut_d2bis(felddaten$, "{"))
'                                      rvf$ = trm(cut_d1(rvf$, "}"))
                                      rvf$ = sidkfeld$
                                  End If
                                Else
                                  kstart$ = strrepl(form1.getusersetting("AuftrittsdruckKontaktStart", ", "), "|", vbCrLf)
                                  If kstart$ = "nurdername" Then
                                    nameonly = True
                                    kstart$ = ""
                                  End If
                                  If nameonly Then
                                    If InStr(felddaten$, "{") > 0 Then
                                      rvf$ = trm(cut_d1(felddaten$, "{"))
                                    End If
                                  Else
                                    rvf$ = sidkfeld$ & kstart$ & rvf$ & form1.getusersetting("AuftrittsdruckKontaktEnde", "")
                                  End If
                                End If
                              End If
                            Else
                              If trm(rvf$) = "" Then rvf$ = sidkfeld$
                            End If
                          End If
                          rvp% = InStr(rvf$, bkmstart$)
                          If rvp% > 0 Then
                            prependlater$ = Mid$(rvf$, rvp%)
                            rvf$ = Left$(rvf$, rvp% - 1)
                          End If
                          If InStr(rvf$, "fromfile:") = 1 Then
                            rvf$ = vorlagendir & "\" + trm(cut_d2bis(rvf$, ":"))
                            If nexist(rvf$) Then
                              rvf$ = "Datei nicht gefunden."
                            Else
                              rvfof% = FreeFile
                              Open rvf$ For Input As #rvfof%
                              rvf$ = ""
                              While Not EOF(rvfof%)
                                Line Input #rvfof%, rvol$
'                                If rvf$ <> "" Then rvf$ = rvf$ + vbCrLf
                                rvf$ = rvf$ + rvol$
                              Wend
                              Close #rvfof%
                            End If
                          End If
                          Call adpprint(adruckmerkslot, p%, repl1310rtf(rvf$)): Call dbgu(rvf$)
                          If fc% < rv.Fields.Count - 1 Then
                            If Not IsNull(rv.Fields(fc% + 1).value) Then
                              Print #p%, " - ";
                            End If
                          End If
                        End If
                      Next fc%
                    Else
                      If tanwaf Then     'textalsnamewennadressefehlt
                        If LCase(orev$) = "name" Then
                          Call adpprint(adruckmerkslot, p%, repl1310rtf(felddaten$)): Call dbgu(felddaten$)
                        End If
                      End If
                    End If
                    Else
                      If Not pacont Then Call adpprint(adruckmerkslot, p%, repl1310rtf(felddaten$))
                    End If
                  End If
                End If
              End If


              Else
'Listenfeld
                cmd$ = "select * from auftritthigru where FeldName='" + bkmlist(bkms%, bk%) + "' and auftrittsid='" + rtmp!id + "';"
                Set rv = New ADODB.Recordset
                rv.CursorLocation = adUseServer
rrr = form1.adoopen(rv, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
                If rrr = 0 Then
                  If Not rv.EOF Then
                    cmd$ = strrepl(trm(rv!felddaten), "|", " - ")
                    cmd$ = strrepl(cmd$, " - " + vbCrLf, vbCrLf)
                    cmd$ = ausklammern(cmd$, "{}")
                    Call adpprint(adruckmerkslot, p%, repl1310rtf(cmd$))
                  End If
                End If
              End If
              bk% = bkmlcount%(bkms%)
            End If
          End If
        Next bk%
      Next bkms%
      If InStr(LCase(ttest$), "summe_") > 0 Then
        Call recalc_honorarliste(Int(sys_mwst))
        If LCase("summe_honorar_brutto") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, honorarsumme1_brutto$ + " " + transe("€")): Call dbgu(honorarsumme1_brutto$)
        End If
        If LCase("summe_honorar_netto") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, honorarsumme1_netto$ + " " + transe("€")): Call dbgu(honorarsumme1_netto$)
        End If
        If LCase("summe_honorar_mwst") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, mwstsumme$ + " " + transe("€")): Call dbgu(mwstsumme$)
        End If
        If LCase("summe_provision_brutto") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, provisionssumme_brutto$ + " " + transe("€")): Call dbgu(provisionssumme_brutto$)
        End If
        If LCase("summe_provision_netto") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, provisionssumme_netto$ + " " + transe("€")): Call dbgu(provisionssumme_netto$)
        End If
        If LCase("summe_provision_mwst") = LCase(ttest$) Then
          Call adpprint(adruckmerkslot, p%, provisionssumme_mwst$ + " " + transe("€")): Call dbgu(provisionssumme_mwst$)
        End If
      End If
      If ttest$ = "tdat" Then
        cmd$ = "select id,datum,zeit,ort from auftritt where (tourneeplanid='" & tpid$ & "' and ((auftrittstyp='künstlerauftritt') or (auftrittstyp='orchesterauftritt'))) order by datum;"
        Set auf = New ADODB.Recordset
        auf.CursorLocation = adUseServer
rrr = form1.adoopen(auf, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not auf.EOF
          cmd$ = "select felddaten from auftritthigru where (auftrittsid='" & auf!id & "' and feldname='halle');"
          Set aufh = New ADODB.Recordset
          aufh.CursorLocation = adUseServer
rrr = form1.adoopen(aufh, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          halle$ = "": If Not aufh.EOF Then halle$ = getnamebyid(aufh!felddaten)
          Print #p%, "{\f53\fs24 {" & datfromsql(auf!datum) & "} Beginn: {" & auf!zeit & "} Uhr {" & auf!ort & "}, {" & halle$ & "} \par }";
          auf.MoveNext
        Wend
      End If
      If ttest$ = "tdat2" Then
        cmd$ = "select id,ort from auftritt where (tourneeplanid='" & tpid$ & "' and ((auftrittstyp='künstlerauftritt') or (auftrittstyp='orchesterauftritt'))) order by datum;"
        Set auf = New ADODB.Recordset
        auf.CursorLocation = adUseServer
rrr = form1.adoopen(auf, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not auf.EOF
          cmd$ = "select felddaten from auftritthigru where (auftrittsid='" & auf!id & "' and feldname='honorar');"
          Set aufh = New ADODB.Recordset
          aufh.CursorLocation = adUseServer
rrr = form1.adoopen(aufh, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          halle$ = "": If Not aufh.EOF Then halle$ = aufh!felddaten
          cmd$ = "select felddaten from auftritthigru where (auftrittsid='" & auf!id & "' and feldname='hinweise');"
          Set aufh = New ADODB.Recordset
          aufh.CursorLocation = adUseServer
rrr = form1.adoopen(aufh, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          hinw$ = "": If Not aufh.EOF Then hinw$ = aufh!felddaten
          Print #p%, "{\f53\fs24 {" & auf!ort & "}, {" & halle$ & "}  {" & hinw$ & "} \par }";
          auf.MoveNext
        Wend
      End If
      ln$ = Mid$(l$, q% + 1)
      Do
        pb% = InStr(LCase(ln$), bkmend$ + LCase(t0$))
        If pb% = 0 Then
          On Error Resume Next
          Line Input #o%, ln
          rrr = Err
          On Error GoTo 0
          If rrr <> 0 Then ln$ = ""
        End If
      Loop Until pb% > 0
      cutoff$ = Left$(ln$, pb% - 1)
      If InStr(cutoff$, "}") > 0 Then cutoff$ = Mid$(cutoff$, InStr(cutoff$, "}") + 1)
      If InStr(cutoff$, "}") > 0 Or InStr(cutoff$, "{") > 0 Then
        dbupgrade.List1.AddItem cutoff$ & ": " + transe("Verdächtige Zeichen in ausgeschnittenem Text.")
        dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
        seriouswarning = False
      End If
      ln$ = Mid$(ln$, pb%)
      If InStr(ln$, "}") = 0 Then
            l$ = ""
      Else
            l$ = Mid$(ln$, InStr(ln$, "}") + 1)
      End If
      If prependlater$ <> "" Then
        l$ = prependlater$ & l$
      End If
    Else
If shwled Then
  cb1.BackColor = RGB(255, 128, 0)
  cb1.Cls
End If
      Print #p%, l$
      l$ = ""
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
    End If
  Wend
Wend
Close #o%
Close #p%
If todofl% = 1 Then
  If wmode$ <> "ical" Then
    i% = 0
    While dochistlist.ListCount > 0
      nop = 0
      dochistlist.ListIndex = 0
      If dochistlist.ListCount > 1 Then
        If dochistlist.List(0) = dochistlist.List(1) Then nop = 1
      End If
      If nop = 0 Then Call sqlqry(dochistlist.List(0))
      dochistlist.RemoveItem 0
    Wend
    dochistlist.Clear
    For listenummer% = 0 To 14
      todo2(listenummer%).Clear
      While todo(listenummer%).ListCount > 0
        todo(listenummer%).ListIndex = 0
        DoEvents
        al_id$ = Left$(todo(listenummer%).List(0), InStr(todo(listenummer%).List(0), ":") - 1)
        al_v$ = Mid$(todo(listenummer%).List(0), InStr(todo(listenummer%).List(0), ":") + 1)
        l4id$ = cut_d2bis(al_v$, "|")
        al_v$ = cut_d1(al_v$, "|")
        dochistlock = True
        todo2(listenummer%).AddItem auftrittsdruck(al_id$, "_" & al_v$, "auftritt", l4id$)
        dochistlock = False
        todo(listenummer%).RemoveItem 0
      Wend
    Next listenummer%
  End If
  Name fn$ As nfn$
  o% = FreeFile
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
  Open nfn$ For Input As #o%
  p% = FreeFile
  Open fn$ For Output As #p%
  While Not EOF(o%)
If shwled Then
  cb1.BackColor = RGB(255, 128, 0)
  cb1.Cls
End If
    Line Input #o%, l$
    glcnt = glcnt + 1
    If (glcnt Mod 1000) = 0 Then
      If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then dbupgrade.List1.AddItem trm(glcnt) + " Zeilen gelesen": dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
      Debug.Print glcnt; " Zeilen gelesen"
    End If
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
    i% = InStr(l$, "#includelist#")
    If i% > 0 Then
      Print #p%, Left(l$, i% - 1);
      listenummer% = Val(Mid$(l$, i% + 13))
If shwled Then
  cb1.BackColor = RGB(255, 128, 0)
  cb1.Cls
End If
      For i% = 0 To todo2(listenummer%).ListCount - 1
        p_add% = FreeFile
        If form1.getusersetting("Textmarkenverfolgen", "nein") = "ja" Then
          dbupgrade.List1.AddItem "liste #" + trm(i%) + "/" + trm(todo2(listenummer%).ListCount - 1) + ": " + todo2(listenummer%).List(i%): dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1: DoEvents
        End If
        Open todo2(listenummer%).List(i%) For Input As #p_add%
        While Not EOF(p_add%)
          Line Input #p_add%, la$
          Print #p%, la$
        Wend
        Close #p_add%
        Kill todo2(listenummer%).List(i%)
      Next i%
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
    Else
If shwled Then
  cb1.BackColor = RGB(255, 128, 0)
  cb1.Cls
End If
      Print #p%, l$
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
    End If
  Wend
  Close #o%
  Close #p%
  If wmode$ <> "ical" Then tplan.MousePointer = 0
  Kill nfn$
  Name fn$ As nfn$
  If xmplarc% > 0 Then
    nfno$ = nfn$ & "Kopie-" & trm(str(xmplarc%)) & ".rtf"
    On Error Resume Next
    Kill nfno$
    Name nfn$ As nfno$
    On Error GoTo 0
  Else
    nfno$ = nfn$
  End If
  If wmode$ <> "ical" Then
    Call form1.openthisdoc(nfno$, "")
  Else
    auftrittsdruck = nfn$
    If form1.getusersetting("Textmarkendebug", "nein") <> "ja" Then Unload dbupgrade
    MousePointer = 0
    Exit Function
  End If
End If
Wend   'xmplarc%
auftrittsdruck = fn$
If Not seriouswarning Then
  If form1.getusersetting("Textmarkendebug", "nein") <> "ja" Then Unload dbupgrade
End If
If tpid$ = "" Then
  Unload tplan
End If
MousePointer = 0
End Function


Private Sub Timer2_Timer()
Dim cmd$, r As ADODB.Recordset, rtmp As ADODB.Recordset, x4$, o%, tme$, rrr, tr, f$, t3a As Boolean
Dim X, t$, i%, isc As Boolean, c$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "Timer2_Timer"
Call dbg2f("Timer2")
t2tick% = t2tick% + 1
If t2tick% > 1000 Then t2tick% = t2tick% - 1000
uexpi$ = datum2sql(Date)

If form1.getusersetting("extralogtlnk", "no") = "ja" Then
  c$ = form1.get1erg("select count(*) as wert from sysvars where owner like'sysvar_system_tlnk_%'")
  If c$ <> dbg_prvtlnkcount Then
    If dbg_prvtlnkcount <> "" Then
      Call form1.log2f("topiccount changed from " + trm(dbg_prvtlnkcount) + " to " + c$, "form1", "delusersetting")
    End If
    dbg_prvtlnkcount = c$
  End If
End If
If autocheckmail Then
  isc = IsConnected()
  If isc Then
    'cmd$ = "Verbindung: " & ConnName
    cmd$ = "Internet: " & ConnName
  Else
    cmd$ = iml("Keine Internetverbindung")
  End If
End If
If getusersetting("pingdbserver", "nein") = "ja" Then
  Call dbg2f("dbping (inet:" + cmd$ + ") " + dbserver$)
  X = Ping(dbserver$, "Agencyprof Servertest", True)
  If X <> 0 Then
    Call dbg2f("dbping1 failed trying again result will not be checked (inet:" + cmd$ + ") " + dbserver$)
    X = Ping(dbserver, "Agencyprof Ping " + dbserver$, True)
  End If
End If
If Label4.Caption <> cmd$ Then
  Label4.Caption = cmd$
  If cmd$ = iml("Keine Internetverbindung") Then
    cbi.BackColor = RGB(255, 0, 0)
  Else
    cbi.BackColor = RGB(0, 255, 0)
  End If
End If
If uId$ = "" Then Exit Sub
If isstarting Then Call form1.startlog(uId$, "1st access to " + s0dir() & "\" & uId$ & ".run")
o% = FreeFile
On Error Resume Next
Open s0dir() & "\" & uId$ & ".run" For Output As #o%
rrr = Err
Close #o%
On Error GoTo 0
'If rrr <> 0 Then
'  MsgBox "Warnung: Kein Zugriff auf " + s0dir() & "\" & uId$ & ".run" + vbCrLf + "Fehler #" + trm(rrr) + ": " + Error$(rrr)
'End If
tme$ = Time
t3a = False
If Not isfieldmissing("opt_checks", "id") Then
  cmd$ = "select ownr,id from opt_checks where (isnull(ownr) or ownr='||' or ownr='' or ownr like '%|" + uId$ + "|%') and dtg<='" + datum2sql(Date) + "' and (isnull(confirmed) or confirmed not like 'ok%') limit 0,1"
  Call dbg2f("Timer2: " + cmd$)
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    If Not rtmp.EOF Then t3a = True
  End If
End If
cmd$ = "select id from todolist where an='" + uId$ + "' and status='neu' and ( datum<'" + datum2sql(Date) + "' or (datum='" + datum2sql(Date) + "' and zeit<='" + tme$ + "'))"
Call dbg2f("Timer2: " + cmd$)
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If Not rtmp.EOF Then t3a = True
End If
If t3a Then
  Timer3.Interval = 1000
  If getusersetting("todo", "normal") = "dringend" Then
    Timer3.Enabled = False
    On Error Resume Next
    Load todolist
    Call todolist.SetFocus
    On Error GoTo 0
    DoEvents
  Else
    Timer3.Enabled = True
  End If
Else
  Command11.Picture = Picture2.Picture
  Timer3.Enabled = False
End If
Call dbg2f("Timer2: Fehlerliste")
errmess.Visible = False
On Error Resume Next
tr = Dir(s0d$ & "\" + docs() + "\" & uId$ & "*.err")
rrr = Err
On Error GoTo 0
If tr <> "" Or rrr <> 0 Then
  errmess.Caption = transe("&Fehler")
  errmess.Visible = True
End If
Call dbg2f("Timer2: SQL- und VCF-Dateien")
On Error Resume Next
tr = Dir(s0d$ & "\*.sql")
rrr = Err
On Error GoTo 0
If tr <> "" Or rrr <> 0 Then
  errmess.Visible = False
  sqlmess.Caption = form1.inmylanguage("SQL")
  sqlmess.Visible = True
  DoEvents
Else
  sqlmess.Visible = False
End If
If Not sqlmess.Visible Then
  On Error Resume Next
  tr = Dir(s0d$ & "\*.vcf")
  rrr = Err
  On Error GoTo 0
  If tr <> "" Or rrr <> 0 Then
    errmess.Visible = False
    sqlmess.Caption = form1.inmylanguage("VCF")
    sqlmess.Visible = True
  Else
    sqlmess.Visible = False
  End If
End If
If (t2tick% Mod 10) = 3 And (autocheckmail Or upop$ = "dir:Inbox") Then
  Call dbg2f("Timer2: Mailtest")
  Call checkmail
End If
If ihavemail Then
  Call dbg2f("Timer2: Mailnachricht")
  If form1.darf_ich_sprechen() = True Then
    f$ = form1.wavdir() & "\..\alarm_mailin.wav"
    If exist(f$) <> 0 Then
      Call sndPlaySound(f$, SND_SYNC)
    End If
  End If
End If

Call dbg2f("Timer2 exit")
End Sub

Private Sub Timer3_Timer()

'd2infile = "Form1": d2insub = "Timer3_Timer"
Call dbg2f("Timer3 start")
t3tick% = t3tick% + 1
If t3tick% >= 1000 Then t3tick% = t3tick% - 1000
If (t3tick% Mod 2) = 1 Then
  Command11.Picture = Picture3.Picture
Else
  Command11.Picture = Picture2.Picture
End If
Call dbg2f("Timer3 exit")

End Sub

Private Sub Timer4_Timer()
Dim o%, fn$, X, i%, rrr, tr As String, adj As Integer

'd2infile = "Form1": d2insub = "Timer4_Timer"
Call dbg2f("Timer4 start")
fn$ = trm(Date & " " & Left(Time, 5))
If Label1.Caption <> fn$ Then
  Label1.Caption = fn$
  If dayvopen Then dayvw.gtoheute.Caption = fn$
End If
'Call dbg2f("Timer4 Teil1 - deldoclist")
If deldoclist.ListCount > 0 Then
  i% = 0
  While i% < deldoclist.ListCount
    On Error Resume Next
    Kill deldoclist.List(i%)
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      deldoclist.RemoveItem i%
    Else
      i% = i% + 1
    End If
  Wend
End If
If Me.BackColor = convertcolor And deldoclist.ListCount = 0 Then Me.BackColor = cleancolor()
tr = Dir(dir_mailoutbox + "\*.amf")
X = Me.BackColor
If tr <> "" Then
  Me.BackColor = RGB(255, 255, 0)
  Me.Caption = "... Emails im Postausgang ..."
  Me.Caption = Me.Caption + " " + transe("Datenbank") + ": " & dbname$ & " auf " + dbserver$ & ", " & s0d$
Else
  If X = RGB(255, 255, 0) Then Me.BackColor = cleancolor()
  Me.Caption = transe("Haupt-Formular - AgencyProf - ")
  Me.Caption = Me.Caption + " " + transe("Datenbank") + ": " & dbname$ & " auf " + dbserver$ & ", " & s0d$
End If
'Call dbg2f("Timer4 Teil3 - teste " & uId$ & ".run")
fn$ = s0d$ & "\" & uId$ & ".run"
If exist(fn$) = 0 Then
  On Error Resume Next
  Open fn$ For Output As #o%
  Close #o%
  On Error GoTo 0
End If
adj = 60 - Val(Right(Time, 2)) + 1
'Debug.Print Right(Time, 2); " "; adj
If adj >= 10 Then
  Timer4.Interval = 10000
Else
  Timer4.Interval = adj * 1000
End If
Call dbg2f("Timer4 exit")
End Sub

Private Sub tmrcld_Timer()
Dim nq As String, nqd As String, nqn As String, i As Integer, o%, l$, rrr, nqnum As Integer

tmrcld.Enabled = False
If Not cloud Then Exit Sub
Call dbg2f("Timer trmcld start")

DoEvents

nq = newcloudqfile()
If nq = "" Then
  cloud = False
  btncld.Caption = "error: no queue access"
  Call dbg2f("Timer trmcld exit")
  Exit Sub
End If
nqd = DirName(nq)
nqnum = cloudupds.ListCount
nqn = "Q: " + trm(nqnum)
If nqnum = 0 And hordex.ListCount = 0 Then
  nqn = nqn + vbCrLf + "in sync"
End If

If nqn <> btncld.Caption Then
  btncld.Caption = nqn
End If
tmrcld.Interval = 17000
tmrcld.Enabled = True
Call dbg2f("Timer trmcld exit")

End Sub

Private Sub trm_todo_Click()
Dim rrr

On Error Resume Next
Load create2do
rrr = Err
On Error GoTo 0

If rrr <> 0 Then Exit Sub
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), "" _
             , "", Date, Left(Time, 5))
create2do.Text1(1).Enabled = False
Call create2do.SetFocus

End Sub

Private Sub trm_todolist_Click()
Call Command11_Click
End Sub

Private Sub trmn_akt_Click()
Call Label8_Click
End Sub

Private Sub trmn_brthd_Click()
Call Command28_Click
End Sub

Private Sub trmn_cal_Click()
Call Command15_Click
End Sub

Private Sub trmn_dayvw_Click()
Call Command29_Click
End Sub

Private Sub trmn_list_Click()
Call Command6_Click
End Sub

Private Sub trmn_new_Click()
Call Command30_Click
End Sub

Private Sub uuid_DblClick()
'd2infile = "Form1": d2insub = "uuid_DblClick"
Load einstellungen

End Sub
Sub rlist3()
Dim rrr
Dim rtmp As ADODB.Recordset, whr$, c$, sdt$, edt$, dkz As Boolean, nosel, rl$, dkzs$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "rlist3"
On Error GoTo errhdl

List3.Clear
List3.AddItem form1.inmylanguage("Heute")
List3.ToolTipText = form1.inmylanguage("Heutige Termine und laufende Projekte")
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM auftritt where datum='" + datum2sql(Date) + "' order by zeit;", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF And break% = 0
  nosel = 1
  dkz = False
  If getusersetting("Dekadenzeigen", "nein") = "ja" Then dkz = True
  If dkz Or projekttyp(trm(rtmp!TourneeplanID)) <> "Dekade" Then
    rl$ = form1.dayofweek(rtmp!datum) + ", " & rtmp!datum & " " & iml(rtmp!auftrittstyp) & "(" & rtmp!bezeichnung & ")"
    If Not IsNull(rtmp!ort) Then rl$ = rl$ & " in " & rtmp!ort
    rl$ = rl$ & Space$(80) + "(AID:" & rtmp!id
    List3.AddItem rl$
  End If
  rtmp.MoveNext
Wend
sdt$ = datum2sql(Date): edt$ = sdt$
dkzs$ = ""
If Not dkz Then dkzs$ = " and (hauptperson<>'Dekade') "
whr$ = "where ( " & _
     "(von<='" & sdt$ & "') and (bis>='" & edt$ & "') " & dkzs$ + _
     ") order by hauptperson,id"
c$ = "select * from tplan " & whr$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  rl$ = form1.inmylanguage("Projekt") & ": " & rtmp!id
  List3.AddItem rl$
  rtmp.MoveNext
Wend
Exit Sub

errhdl:
  On Error GoTo 0
  break% = 1
  Exit Sub

End Sub


Public Sub upd_colorcache(t$, r%, g%, b%)
Dim i%, cid$

'd2infile = "Form1": d2insub = "upd_colorcache"
For i% = 0 To colorcachepointer% - 1
  If colorcacheid$(i%) = cid$ Then
    colorcache%(i%, 0) = r%
    colorcache%(i%, 1) = g%
    colorcache%(i%, 2) = b%
    i% = colorcachepointer% + 111
   End If
Next i%
If i% > colorcachepointer% Then
  colorcache%(colorcachepointer%, 0) = r%
  colorcache%(colorcachepointer%, 1) = g%
  colorcache%(colorcachepointer%, 2) = b%
  colorcachepointer% = colorcachepointer% + 1
  If colorcachepointer% > 99 Then colorcachepointer% = 99
End If

End Sub

Public Function get_eventstatuscolor(evid) As Long

'd2infile = "Form1": d2insub = "get_eventstatuscolor"
If IsNull(evid) Or evid < 0 Or evid > 9 Then
  get_eventstatuscolor = RGB(255, 255, 255)
Else
  get_eventstatuscolor = statusfarbe(evid)
End If

End Function
Public Function get_eventstatusname(evid) As String

'd2infile = "Form1": d2insub = "get_eventstatusname"
If IsNull(evid) Or evid < 0 Or evid > 9 Then
  get_eventstatusname = transe("kein Status")
Else
  get_eventstatusname = transe(statusname$(evid))
End If

End Function

Public Function get_eventcolor(evid$) As Long
Dim rrr
Dim at As ADODB.Recordset, i%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_eventcolor"
get_eventcolor = RGB(255, 255, 255)
For i% = 0 To colorcachepointer% - 1
  If colorcacheid$(i%) = evid$ Then
    get_eventcolor = RGB(colorcache%(i%, 0), colorcache%(i%, 1), colorcache%(i%, 2))
    Exit Function
  End If
Next i%
Set at = New ADODB.Recordset
at.CursorLocation = adUseServer
rrr = form1.adoopen(at, "SELECT * FROM auftrittstypen where id='" + evid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not at.EOF Then
  colorcacheid$(colorcachepointer%) = evid$
  For i% = 0 To 2
    colorcache%(colorcachepointer%, 2) = at.Fields(2 + i%).value
  Next i%
  get_eventcolor = RGB(colorcache%(colorcachepointer%, 0), colorcache%(colorcachepointer%, 1), colorcache%(colorcachepointer%, 2))
  colorcachepointer% = colorcachepointer% + 1
  If colorcachepointer% > 99 Then colorcachepointer% = 99
End If

End Function

Public Sub calcall(dtg$)

'd2infile = "Form1": d2insub = "calcall"
'  Load k2
'  Call Kalender.SetFocus
'  Call k2.gotodate(dtg$)

End Sub
Public Function nurdiewaehrung(l$) As String
Dim i%, n$, w$

'd2infile = "Form1": d2insub = "nurdiewaehrung"
For i% = 0 To waehrungen.ListCount - 1
  n$ = waehrungen.List(i%): n$ = Mid$(n$, InStr(n$, ":") + 1): n$ = Mid$(n$, InStr(n$, ":") + 1)
  w$ = waehrungen.List(i%): w$ = cut_d1(w$, ":")
  If InStr(l$, w$) Then
    nurdiewaehrung = w$
    Exit Function
  End If
  If n$ <> "" And InStr(l$, n$) Then
    nurdiewaehrung = n$
    Exit Function
  End If
Next i%
nurdiewaehrung = transe("€")

End Function
Public Function ohnewaehrung(w1a$) As String
Dim wa$, w$, t$

'd2infile = "Form1": d2insub = "ohnewaehrung"
wa$ = w1a$
w$ = nurdiewaehrung(wa$)
If InStr(wa$, w$) = "0" Then wa$ = wa$ + " " & w$
t$ = ""
If wa$ <> "" And w$ <> "" Then t$ = trm(Left$(wa$, InStr(wa$, w$) - 1))
If t$ = "" Then
   w$ = "0"
Else
   w$ = strrepl(t$, "--", "00")
End If

ohnewaehrung = w$

End Function
Public Function kursvom(wae$, wrdbl$) As String
Dim rrr
Dim r As ADODB.Recordset, wann$, i%, n$, w$, fx$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "kursvom"
If wae$ = "" Then
  kursvom = 1
  Exit Function
End If
wann$ = datum2sql(wrdbl$)
For i% = 0 To waehrungen.ListCount - 1
  n$ = waehrungen.List(i%): n$ = Mid$(n$, InStr(n$, ":") + 1): n$ = Mid$(n$, InStr(n$, ":") + 1)
  w$ = waehrungen.List(i%): w$ = cut_d1(w$, ":")
  If wae$ = n$ Then wae$ = w$
  If wae$ = w$ Then
    fx$ = Mid$(waehrungen.List(i%), InStr(waehrungen.List(i%), ":") + 1)
    If Left(fx$, 2) = "ja" Then
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT kurs FROM waehrung where id='" + wae$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then
         kursvom = r!kurs
         Exit Function
      End If
    Else
      If wann$ >= datum2sql(Date) Then
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM kurse where wid='" + wae$ + "' order by id", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          r.MoveLast
          kursvom = r!kurs
          Exit Function
        End If
      Else
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM kurse where wid='" + wae$ + "' and id<='" + wann$ + "ZZZ' order by id desc", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          kursvom = r!kurs
          Exit Function
        End If
      End If
    End If
    i% = waehrungen.ListCount - 1
  End If
Next i%

kursvom = "0"
End Function
Public Function kursdatum(wae$, wrdbl$) As String
Dim rrr
Dim r As ADODB.Recordset, wann$, w$, n$, i%, fx$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "kursdatum"
wann$ = datum2sql(wrdbl$)
For i% = 0 To waehrungen.ListCount - 1
  n$ = waehrungen.List(i%): n$ = Mid$(n$, InStr(n$, ":") + 1): n$ = Mid$(n$, InStr(n$, ":") + 1)
  w$ = waehrungen.List(i%): w$ = cut_d1(w$, ":")
  If wae$ = n$ Then wae$ = w$
  If wae$ = w$ Then
    fx$ = Mid$(waehrungen.List(i%), InStr(waehrungen.List(i%), ":") + 1)
    If Left(fx$, 4) = "nein" Then
      If wann$ >= datum2sql(Date) Then
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM kurse where wid='" + wae$ + "' order by id", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          r.MoveLast
          kursdatum = datfromsql(Left$(r!id, Len(r!id) - Len(wae$)))
          Exit Function
        End If
      Else
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM kurse where wid='" + wae$ + "' and id<='" + wann$ + "ZZZ' order by id desc", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          kursdatum = wrdbl$
          If Not r.EOF Then
          kursdatum = datfromsql(Left(r!id, 10))
        Else
          kursdatum = "1.1.1970"
        End If
        Exit Function
      End If
    End If
    i% = waehrungen.ListCount - 1
  End If
Next i%

kursdatum = datfromsql(wann$)
End Function

Public Function iwanttooltips()

'd2infile = "Form1": d2insub = "iwanttooltips"
iwanttooltips = uwantstooltips%
End Function

Function inworten(l As Long) As String
Dim erg1$
Dim i%, j%

Static p$(0 To 2)
ReDim z%(0 To 2), erg$(0 To 2)


If InStr(l, ",") > 1 Then
  l = Left$(l, InStr(l, ",") - 1)
End If
p$(2) = "million": p$(1) = "tausend": p$(0) = ""
l = Abs(l) ' negatve Zahlen schenken wir uns
erg1$ = Mid$(str$(l), 2)
While Len(erg1$) < 9
  erg1$ = " " + erg1$
Wend
For i% = 7 To 1 Step -3
  z%(j%) = Val(Mid$(erg1$, i%, 3))
  j% = j% + 1
Next i%
For j% = 2 To 0 Step -1
  erg$(j%) = mmkw(z%(j%))
  If Len(erg$(j%)) > 0 Then erg$(j%) = erg$(j%) + p$(j%)
  If j% > 0 And InStr(erg$(j%), "eins") = Len(erg$(j%)) - 3 Then
    erg$(j%) = Left$(erg$(j%), Len(erg$(j%)) - 1)
  End If
Next j%
inworten = erg$(2) + erg$(1) + erg$(0)

End Function

Function mmkw(w) As String

Static es$(0 To 9), ss$(0 To 3), zs$(0 To 9)

es$(0) = ""
es$(1) = "ein"
es$(2) = "zwei"
es$(3) = "drei"
es$(4) = "vier"
es$(5) = "fünf"
es$(6) = "sechs"
es$(7) = "sieben"
es$(8) = "acht"
es$(9) = "neun"
ss$(0) = "zehn"
ss$(1) = "elf"
ss$(2) = "zwölf"
zs$(0) = ""
zs$(1) = "xxx"
zs$(2) = "zwanzig"
zs$(3) = "dreißig"
zs$(4) = "vierzig"
zs$(5) = "fünfzig"
zs$(6) = "sechzig"
zs$(7) = "siebzig"
zs$(8) = "achtzig"
zs$(9) = "neunzig"

Dim e%, z%, h%
Dim hstr$, zstr$

e% = w Mod 10
z% = Int(((w Mod 100) - e) / 10)
h% = Int(w / 100)
If h% > 0 Then
  If z > 0 Or e > 0 Then
    hstr$ = es$(h%) + "hundertund"
  Else
    hstr$ = es$(h%) + "hundert"
  End If
End If
If z% > 0 Then
  If z% = 1 Then
    If e% < 3 Then
      zstr$ = ss$(e%)
    Else
      zstr$ = es$(e%) + "zehn"
    End If
  Else
    If e% > 0 Then
      zstr$ = es$(e%) + "und" + zs$(z%)
    Else
      zstr$ = zs$(z%)
    End If
  End If
Else
  If e% = 1 Then
    zstr$ = "eins"
  Else
    zstr$ = es$(e%)
  End If
End If
mmkw = hstr$ + zstr$

End Function

Function grantadrtyp() As String
Dim i As Integer

'd2infile = "Form1": d2insub = "grantadrtyp"
grantadrtyp = ""
If grantptr < 0 Then Exit Function
For i = 0 To grantptr
  If Left(granttab(i), 9) = "adresse__" Then
    grantadrtyp = Mid(granttab(i), 10)
    Exit For
  End If
Next i

End Function
Function granted(s$) As Boolean
Dim t$, ta$, t0$, i As Integer, p As Integer, j As Integer, rest$, ta0$

'd2infile = "Form1": d2insub = "granted"
granted = False
If grantptr < 0 Then
  granted = True
  Exit Function
End If
t0$ = s$
t$ = LCase(s$)
ta$ = ""
If Left$(t$, 12) = "delete from " Then
  ta$ = Mid$(t$, 13): rest$ = Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
Else
  If Left$(t$, 7) = "update " Then
    ta$ = Mid$(t$, 8): rest$ = Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
  Else
    If Left$(t$, 12) = "insert into " Then
      ta$ = Mid$(t$, 13): rest$ = Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
    End If
  End If
End If
'If ta$ = "adresstyp" Then ta$ = "adresse"
For i = 0 To grantptr
  If granttab(i) = ta$ Then
    granted = True
    Exit For
  End If
  If ta$ = "adresse" Or ta$ = "auftritthigru" Or ta$ = "kontakt" Then
    p = InStr(granttab(i), "__")
    If p > 0 Then
      ta0$ = Mid(granttab(i), p + 2)
      For j = 0 To shwAdrDetail.List1.ListCount - 1
        If cut_d1(shwAdrDetail.List1.List(j), ":") = ta0$ Then
          granted = True
          Exit Function
        End If
      Next j
    End If
  End If
  If ta$ = "kontakt" Then
    p = InStr(granttab(i), "__")
    If p > 0 Then
      ta0$ = Mid(granttab(i), p + 2)
      For j = 0 To shwAdrDetail.List1.ListCount - 1
        If LCase(cut_d1(shwAdrDetail.List1.List(j), ":")) = ta0$ Then
          granted = True
          Exit For
        End If
      Next j
    End If
  End If
Next i
End Function

Sub chkalarmlist(s$)
Dim t$, ta$, t0$, rest$

'd2infile = "Form1": d2insub = "chkalarmlist"
If noalarms Then Exit Sub
t0$ = s$
t$ = LCase(s$)
ta$ = ""
If Left$(t$, 12) = "delete from " Then
  ta$ = Mid$(t$, 13): rest$ = "lösche: " + Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
Else
  If Left$(t$, 7) = "update " Then
    ta$ = Mid$(t$, 8): rest$ = Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
  Else
    If Left$(t$, 12) = "insert into " Then
      ta$ = Mid$(t$, 13): rest$ = Mid$(ta$, InStr(ta$, " ") + 1): ta$ = cut_d1(ta$, " ")
    End If
  End If
End If

If ta$ <> "" Then
  Call runalarms(ta$, rest$)
End If

End Sub

Sub runalarms(t$, r$)
Dim rrr, g$, wh$, p%, ti$, qst$
Dim idnam$, re As ADODB.Recordset, trg%, todo As ADODB.Recordset, p1ing%, cmd$, fnd%, msg$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "runalarms"
On Error GoTo errhdl
If t$ = "alarmliste" Then Exit Sub
If t$ = "todolist" Then Exit Sub
If t$ = "dochist" Then Exit Sub

trg% = 0
idnam$ = LCase(sqla.TableDefs(t$).Fields(0).name)
If InStr(r$, idnam$) > 0 Then trg% = 1
cmd$ = "select * from alarmliste where tabelle='" + t$ + "'"
Set re = New ADODB.Recordset
re.CursorLocation = adUseServer
rrr = form1.adoopen(re, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
'Debug.Print cmd$
While Not re.EOF
  p1ing% = 0
'Debug.Print "g=" + trm(re!ontid)
  If LCase(trm(re!uId)) <> LCase(getuserid()) Then
  If trg% = 1 Then
    g$ = trm(re!ontid)
    If InStr(g$, "GRP:") = 1 Then
      If t$ = "adresse" Or t$ = "kontakt" Then
        g$ = Mid(g$, 5)
        p% = InStr(r$, "where id='")
        If p% > 0 Then
          ti$ = Mid(r$, p% + 10): ti$ = Left(ti$, Len(ti$) - 1)
          If t$ = "adresse" Then
            If isoftype(ti$, g$) <> "-1" Then p1ing% = 1
          Else
            If kisoftype(ti$, g$) <> "-1" Then p1ing% = 1
          End If
        End If
        If p% = 0 Then
          If Left(r$, 4) = "(id," Then
            p% = InStr(r$, "values('")
            If p% > 0 Then
              ti$ = cut_d1(Mid(r$, p% + 8), "'")
              If t$ = "adresse" Then
                If isoftype(ti$, g$) <> "-1" Then p1ing% = 1
              Else
                If kisoftype(ti$, g$) <> "-1" Then p1ing% = 1
              End If
            End If
          End If
        End If
      End If
    Else
      ti$ = cut_d1(g$, "|")
      If InStr(LCase(r$), LCase(ti$)) > 0 Then
        If t$ = "adresstyp" Then
          ti$ = cut_d2bis(g$, "|")
          p% = InStr(r$, "wert='")
          If p% > 0 Then
            qst$ = Mid$(r$, p% + 6)
            p% = InStr(qst$, "'")
            If p% > 1 Then qst$ = Left(qst$, p% - 1)
            If InStr(LCase(qst$), ti$) > 0 Then p1ing% = 1
          End If
        Else
          p1ing% = 1
        End If
      End If
    End If
  End If
  If p1ing% = 1 Or Len(re!ontid) = 0 Then
    cmd$ = "select * from todolist where status='neu' and an='" + re!uId + "' and von='" + t$ + "'"
    Set todo = New ADODB.Recordset
    todo.CursorLocation = adUseServer
    rrr = form1.adoopen(todo, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    fnd% = 0
    While Not todo.EOF And fnd% = 0
      If Len(re!ontid) = 0 Or InStr(trm(re!ontid), "GRP:") > 0 Then
        If p1ing% = 1 Then
          If InStr(todo!nachricht, strrepl(r$, "'", "-")) > 0 Then
            fnd% = 1
          End If
        Else
          fnd% = 1
        End If
      Else
        If InStr(todo!betreff, re!ontid) > 0 Then
          fnd% = 1
        End If
      End If
      todo.MoveNext
    Wend
    If fnd% = 0 Then
      msg$ = "Änderung an " + t$
      If Len(re!ontid) > 0 And InStr(trm(re!ontid), "GRP:") = 0 Then msg$ = msg$ + " id=" + re!ontid
      Call new2do(t$, re!uId, msg$, strrepl(r$, "'", "-"), datum2sql(Date), 0, 0, "", 0)
    End If
    todo.Close
  End If
  End If
  re.MoveNext
Wend
re.Close
Exit Sub
errhdl:
On Error GoTo 0

End Sub

Public Function getdbname()
'd2infile = "Form1": d2insub = "getdbname"
getdbname = dbname$
End Function
Public Function getdbserver()
'd2infile = "Form1": d2insub = "getdbserver"
getdbserver = dbserver$
End Function
Public Function getdbuid()
'd2infile = "Form1": d2insub = "getdbuid"
getdbuid = dbuid$
End Function
Public Function getdbpsswd()
'd2infile = "Form1": d2insub = "getdbpsswd"
getdbpsswd = dbpsswd$
End Function
Public Function getconnstr()
'd2infile = "Form1": d2insub = "getconnstr"
getconnstr = dbpara$
End Function
Public Function getadoconnstr()
'd2infile = "Form1": d2insub = "getadoconnstr"
getadoconnstr = adopara$
End Function

Function read_sysvar(varnam$, usr$) As String
Dim rrr
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "read_sysvar"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT wert FROM sysvars where id='" + varnam$ + "' and owner='" + usr$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  sqlqry ("insert into sysvars (id,owner,wert) values('" + varnam$ + "','" + usr$ + "','0')")
  read_sysvar = "0"
  r.Close
  Exit Function
End If
read_sysvar = r!wert
r.Close

End Function

Public Function medienname(d$)
Dim r$

'd2infile = "Form1": d2insub = "medienname"
r$ = strrepl(d$, ",", "")
If InStr(r$, "{") > 0 Then
  r$ = cut_d2bis(r$, "{")
  r$ = cut_d1(r$, "}")
End If
r$ = strrepl(r$, " ", "_")
r$ = strrepl(r$, vbCrLf, "_")
medienname = r$

End Function

Public Function cleancolor() As Long

'd2infile = "Form1": d2insub = "cleancolor"
cleancolor = &HE4E4E4

End Function

Public Function dirtycolor() As Long

'd2infile = "Form1": d2insub = "dirtycolor"
dirtycolor = dirtcol
'dirtycolor = RGB(211, 238, 249)


End Function

Function higruzeilen(typ$, fn$)
Dim rrr
Dim at As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "higruzeilen"
higruzeilen = 0
Set at = New ADODB.Recordset
at.CursorLocation = adUseServer
rrr = form1.adoopen(at, "SELECT zeilen FROM auftrittsfelder where typ='" + typ$ + "' and feldname='" + fn$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not at.EOF Then higruzeilen = at!zeilen

End Function

Public Function mylasttop(f$)
Dim inifile As String, l$, o%

'd2infile = "Form1": d2insub = "mylasttop"
l$ = "20"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".top"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylasttop = CInt(l$)
'If getusersetting("limitforms2screen", "ja") = "ja" Then
'  If mylasttop + Me.Top > Screen.Height Then
'    mylasttop = Screen.Height - Me.Height
'  End If
'End If

End Function
Public Function mylastleft(f$)
Dim inifile As String, l$, o%

'd2infile = "Form1": d2insub = "mylastleft"
l$ = "20"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".lft"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylastleft = CInt(l$)
'If getusersetting("limitforms2screen", "ja") = "ja" Then
'  If mylastleft + Me.Width > Screen.Width Then
'    mylastleft = Screen.Width - Me.Width
'  End If
'End If
End Function
Public Sub setmylastleft(f$, wert%)
Dim inifile As String, o%, rrr

'd2infile = "Form1": d2insub = "setmylastleft"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".lft"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

Print #o%, wert%
Close #o%

End Sub

Public Sub setmylasttop(f$, wert%)
Dim inifile As String, o%, rrr

'd2infile = "Form1": d2insub = "setmylasttop"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".top"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, wert%
  Close #o%
End If
End Sub


Public Sub setcolorselected(wert As Long)
'd2infile = "Form1": d2insub = "setcolorselected"
selectedcolor = wert

End Sub

Public Function getdateselected()
'd2infile = "Form1": d2insub = "getdateselected"
getdateselected = SelectedDate$

End Function

Public Function getcolorselected() As Long
'd2infile = "Form1": d2insub = "getcolorselected"
getcolorselected = selectedcolor

End Function

Public Function mylastwidth(f$, mode%)
Dim inifile As String, l$, o%

'd2infile = "Form1": d2insub = "mylastwidth"
If mode% = 0 Then
  l$ = "0"
Else
  l$ = str$(Width)
End If
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".wdt"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylastwidth = CInt(l$)

End Function

Public Sub setmyautoopen(f$, wert%)
Dim inifile As String, o%, rrr

'd2infile = "Form1": d2insub = "setmyautoopen"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".aut"
If wert% = 0 Then
  On Error Resume Next
  Kill inifile
  On Error GoTo 0
Else
  o% = FreeFile
  On Error Resume Next
  Open inifile For Output As #o%
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  Print #o%, wert%
  Close #o%
End If

End Sub

Public Sub setmylastwidth(f$, wert%)
Dim inifile As String, o%, rrr

'd2infile = "Form1": d2insub = "setmylastwidth"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".wdt"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
Print #o%, wert%
Close #o%

End Sub
Public Function mylastheight(f$, mode%)
Dim inifile As String, l$, o%

'd2infile = "Form1": d2insub = "mylastheight"
If mode% = 0 Then
  l$ = "0"
Else
  l$ = str$(Width)
End If
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".hgt"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
mylastheight = CInt(l$)

End Function

Public Sub setmylastheight(f$, wert%)
Dim inifile As String, o%, rrr


'd2infile = "Form1": d2insub = "setmylastheight"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + ".hgt"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, wert%
  Close #o%
End If
End Sub

Public Sub memonoshow()
'd2infile = "Form1": d2insub = "memonoshow"
memono% = 1
End Sub

Public Function myuniqueemlname() As String
Dim o%, fn$, rrr, i%

'd2infile = "Form1": d2insub = "myuniqueemlname"
myuniqueemlname = ""
o% = FreeFile
fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  On Error Resume Next
  MkDir s0d$ & "\" + docs() + ""
  MkDir s0d$ & "\" + docs() + "\" & uId$
  On Error GoTo 0
Else
  Close #o%
  Kill fn$
End If
i% = 0
If uemlhiscore% > 0 Then i% = uemlhiscore%
Do
  i% = i% + 1
  fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\" & uId$ & Left$(Date, 2) & Mid$(Date, 4, 2) & trm(str$(i%)) & ".apm"
Loop Until exist(fn$) = 0
uemlhiscore% = i%
myuniqueemlname = fn$

End Function

Sub dbupgrd()
Dim fn$, o%, p%, q%, l$, lvl%, amj, amn, mde%, dbx$, brkupd%

'd2infile = "Form1": d2insub = "dbupgrd"
brkupd% = 0
lvl% = Val(read_sysvar("dblevel", "_"))
If exist("rtfs.ini") <> 0 And InStr(LCase(dbname$), "apdemo") > 0 And InStr(LCase(App.EXEName), "apadmin") = 0 Then
  Call dbg2f("loading dbupgrade")
  Load dbupgrade
  dbupgrade.List1.Clear
  MousePointer = 11
  DoEvents
  On Error Resume Next
  MkDir s0d$ + "\apdemo.mdb.rtf"
  MkDir s0d$ + "\apdemo.rtf"
  On Error GoTo 0
  mde% = 0
  Call dbg2f("entpacke rtfs.ini")
  o% = FreeFile
  Open "rtfs.ini" For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    If mde% = 0 Then
      mde% = 1
      dbupgrade.List1.AddItem l$
      dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1
      p% = FreeFile
      Open form1.s0dir() + "\apdemo.mdb.rtf\" & l$ For Output As #p%
      q% = FreeFile
      Open form1.s0dir() + "\apdemo.rtf\" + l$ For Output As #q%
      DoEvents
    Else
      If l$ = "***EOF***AGENCYPROF***" Then
        mde% = 0
        Close #q%
        Close #p%
      Else
        Print #p%, l$
        Print #q%, l$
      End If
    End If
  Wend
  Close #o%
  On Error Resume Next
  Kill "rtfs.bak"
  Name "rtfs.ini" As "rtfs.bak"
  On Error GoTo 0
  Unload dbupgrade
  Call dbg2f("dbupgrade entlagden, rtfs.ini entpackt")
End If
MousePointer = 0

'On Error Resume Next
'Unload dbupgrade
'On Error GoTo 0
Call dbg2f("beendet")

End Sub

Sub crShell(fn$, addpause As Boolean)
Dim o%, c$, X

'd2infile = "Form1": d2insub = "crShell"
o% = FreeFile
Open fn$ + ".bat" For Output As #o%
c$ = form1.getmymysqld() + " -h " + form1.getmymysqlhost() + " -u " & form1.getdbuid() & " "
If form1.getdbpsswd() <> "" Then c$ = c$ & "-p" & form1.getdbpsswd() & " "
c$ = c$ & "-D " + form1.getdbname() + " <" + fn$ + ".txt"
Print #o%, "type " + fn$ + ".txt"
Print #o%, c$
If addpause Then Print #o%, "pause"
Close #o%
'X = Shell("notepad.exe " & fn$ + ".txt", 1)
X = Shell(fn$ + ".bat", 1)

End Sub


Function dval(a$)
'd2infile = "Form1": d2insub = "dval"
dval = 0
a$ = LCase(a$)
If a$ >= "0" And a$ <= "9" Then
  dval = Val(a$)
Else
  If a$ >= "a" And a$ <= "f" Then
    dval = (Asc(a$) - Asc("a")) + 10
  Else
    dval = 0
  End If
End If

End Function

Public Function getusersetting(fldn$, Optional vifnull As String) As String
Dim r As ADODB.Recordset, vin As String, c1md$, rrr, c$, i%, j%, sptr%, tooktime As Long

Dim d2infile As String, d2insub As String
Call tm_start(0)
d2infile = "Form1": d2insub = "getusersetting"
vin = "": sptr% = 0
If vifnull <> "" Then vin = vifnull
getusersetting = vin
If useusrcache = "ja" Then
  For i% = 0 To 199
    If usr_setting(0, i%) = LCase(fldn$) Then
      getusersetting = usr_setting(1, i%)
      If usr_set_hits(i%) < 10000 Then usr_set_hits(i%) = usr_set_hits(i%) + 1
'Debug.Print "getusrsetting f. cache:"; i%; "("; trm(usr_set_hits(i%)); "): "; fldn$; "="; usr_setting(1, i%)
      If i% > 0 Then
        If usr_set_hits(i%) > usr_set_hits(i% - 1) Then
          j% = usr_set_hits(i%): usr_set_hits(i%) = usr_set_hits(i% - 1): usr_set_hits(i% - 1) = j%
          For j% = 0 To 1
            c$ = usr_setting(j%, i%): usr_setting(j%, i%) = usr_setting(j%, i% - 1): usr_setting(j%, i% - 1) = c$
          Next j%
        End If
      End If
      tooktime = tm_stop(0)
'      Debug.Print "tooktime=" + trm(tooktime) + " ms (cached), cache #"; trm(sptr%); ": "; usr_setting(0, sptr%); "="; usr_setting(1, sptr%)
      Exit Function
    Else
      If usr_setting(0, i%) = "" Then
        sptr% = i%
        Exit For
      End If
    End If
  Next i%
End If
If sptr% > 199 Then sptr% = 0
On Error Resume Next
c1md$ = "SELECT " + fldn$ + " as rc FROM benutzerdaten where id='" + uId$ + "'"
If isstarting And fldn$ <> "localdir" Then Call startlog(getuserid(), "getusersetting:" + c1md$)
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c1md$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If Not r.EOF Then
    If Not IsNull(r!rc) Then getusersetting = r!rc
    r.Close
    If isstarting And fldn$ <> "localdir" And InStr(LCase(fldn$), "passw") = 0 Then Call startlog(getuserid(), "getusersetting returns:" + trm(r!rc))
    usr_setting(0, sptr%) = LCase(fldn$)
    usr_setting(1, sptr%) = getusersetting
    tooktime = tm_stop(0)
    Exit Function
  End If
End If

c$ = "SELECT wert as rc FROM sysvars where owner='sysvar_" & uId$ & "_" & fldn$ & "'"
If isstarting And fldn$ <> "localdir" And InStr(LCase(fldn$), "passw") = 0 Then Call startlog(getuserid(), "getusersetting:" + trm(c$))
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  tooktime = tm_stop(0)
  Exit Function
End If
If Not r.EOF Then
  If Not IsNull(r!rc) Then getusersetting = internaldecrypt(r!rc)
  If isstarting And fldn$ <> "localdir" And InStr(LCase(fldn$), "passw") = 0 Then Call startlog(getuserid(), "getusersetting:" + trm(r!rc))
  usr_setting(0, sptr%) = fldn$
  usr_setting(1, sptr%) = internaldecrypt(r!rc)
  r.Close
  tooktime = tm_stop(0)
'  Debug.Print "tooktime=" + trm(tooktime) + ", cache #"; trm(sptr%); ": "; usr_setting(0, sptr%); "="; usr_setting(1, sptr%)
  Exit Function
End If
getusersetting = getsystemsetting(fldn$, vin)
If isstarting And fldn$ <> "localdir" And InStr(LCase(fldn$), "passw") = 0 Then Call startlog(getuserid(), "getusersetting returns:" + trm(getusersetting))
usr_setting(0, sptr%) = LCase(fldn$)
usr_setting(1, sptr%) = getusersetting
tooktime = tm_stop(0)
'Debug.Print "tooktime=" + trm(tooktime) + " ms, cache #"; trm(sptr%); ": "; usr_setting(0, sptr%); "="; usr_setting(1, sptr%)
End Function

Public Sub setbenutzerdaten(fldn$, wert$)
Dim w$, c$

'd2infile = "Form1": d2insub = "setbenutzerdaten"
w$ = wert$
c$ = "update benutzerdaten set " & fldn$ & "='" & w$ & "' where id='" & uId$ & "'"
Call sqlqry(c$)
Select Case fldn$
  Case "immer_kalender": ucalalways$ = wert$
  Case Else:
End Select

End Sub

Public Sub pg_xp(t$, txt$)
Dim rrr
Dim o%, r As ADODB.Recordset, rtmp As ADODB.Recordset, i%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "pg_xp"
MousePointer = 11
o% = FreeFile
Open txt$ For Append As #o%

Print #o%,
Print #o%, "drop table " + t$ + ";"
Print #o%, "create table " + t$ + "("
For i% = 0 To sqla.TableDefs(t$).Fields.Count - 1
  Select Case sqla.TableDefs(t$).Fields(i%).Type
    Case 10: Print #o%, sqla.TableDefs(t$).Fields(i%).name; " char(" & trm(sqla.TableDefs(t$).Fields(i%).Size) & ")";
    Case 8: Print #o%, sqla.TableDefs(t$).Fields(i%).name; " char(20)";
    Case 12: Print #o%, sqla.TableDefs(t$).Fields(i%).name; " char(250)";
    Case 4: Print #o%, sqla.TableDefs(t$).Fields(i%).name; " int4";
    Case Else:
  End Select
  If i% < sqla.TableDefs(t$).Fields.Count - 1 Then Print #o%, ","
Next i%
Print #o%,
Print #o%, ");"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM " + t$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  Print #o%, "insert into " + t$ + " values(";
  For i% = 0 To sqla.TableDefs(t$).Fields.Count - 1
  Select Case sqla.TableDefs(t$).Fields(i%).Type
    Case 4: Print #o%, rtmp.Fields(i%).value;
    Case Else: Print #o%, "'"; rtmp.Fields(i%).value; "'";
  End Select
  If i% <> sqla.TableDefs(t$).Fields.Count - 1 Then Print #o, ",";
  Next i%
  Print #o%, ");"
  rtmp.MoveNext
Wend

Close #o%
MousePointer = 0
End Sub

Public Function ask2save()

'd2infile = "Form1": d2insub = "ask2save"
ask2save = MsgBox("Sie haben die Daten geändert; soll ich speichern?", vbYesNo + vbCritical + vbDefaultButton1, "Daten speichern?")

End Function

Public Function InboxMessageCount(dn As String) As Integer
Dim tr, n As Integer, rrr

n = 0
InboxMessageCount = POP_SOCKET_ERROR
On Error GoTo rout6767
tr = Dir(dn + "\*.amf")
rrr = Err
On Error GoTo 0
If rrr <> 0 Then GoTo rout6767
While tr <> ""
  n = n + 1
  tr = Dir
Wend
InboxMessageCount = n
rout6767:
On Error GoTo 0
End Function

Public Function myinboxname() As String
Dim o%, dn$, fn$, rrr

'd2infile = "Form1": d2insub = "myinboxname"
myinboxname = ""
o% = FreeFile
dn$ = s0d$ & "\" + docs() + "\" & uId$ & "\mail"
fn$ = dn$ & "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  On Error Resume Next
  MkDir s0d$ & "\" + docs() + ""
  MkDir s0d$ & "\" + docs() + "\" & uId$
  MkDir s0d$ & "\" + docs() + "\" & uId$ + "\mail"
  On Error GoTo 0
Else
  Close #o%
  Kill fn$
End If
dn$ = dn$ & "\" & Right$(Date, 2)
On Error Resume Next
MkDir dn$
On Error GoTo 0
dn$ = dn$ & "\" & Mid$(Date, 4, 2)
On Error Resume Next
MkDir dn$
On Error GoTo 0
myinboxname = dn$
End Function

Public Function myuniqueinboxname() As String
Dim o%, dn$, fn$, rrr, i%, fnext$

'd2infile = "Form1": d2insub = "myuniqueinboxname"
myuniqueinboxname = ""
o% = FreeFile
fnext$ = "." + getusersetting("mailfileextension", "eml")
dn$ = s0d$ & "\" + docs() + "\" & uId$ & "\mail"
fn$ = dn$ & "\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  On Error Resume Next
  MkDir s0d$ & "\" + docs() + ""
  MkDir s0d$ & "\" + docs() + "\" & uId$
  MkDir s0d$ & "\" + docs() + "\" & uId$ + "\mail"
  On Error GoTo 0
Else
  Close #o%
  Kill fn$
End If
dn$ = dn$ & "\" & Mid$(datum2sql(Date), 3, 2)
On Error Resume Next
MkDir dn$
On Error GoTo 0
dn$ = dn$ & "\" & Mid$(datum2sql(Date), 6, 2)
On Error Resume Next
MkDir dn$
On Error GoTo 0
i% = 0
If uinbhiscore% > 0 Then i% = uinbhiscore%
Do
  i% = i% + 1
  fn$ = dn & "\" & uId$ & datum2sql(Date) & "-" & trm(str$(i%)) & fnext$
Loop Until exist(fn$) = 0
uinbhiscore% = i%
myuniqueinboxname = fn$

End Function
Public Function myinboxdir() As String
Dim o%

'd2infile = "Form1": d2insub = "myinboxdir"
myinboxdir = s0d$ & "\" + docs() + "\" & uId$ & "\mail"

End Function
Sub noerrshow()

errsh% = 0

End Sub
Sub errshow()

errsh% = 1

End Sub


Sub chgcreate(l$)
Dim o%

'd2infile = "Form1": d2insub = "chgcreate"
o% = FreeFile
Open "sqlchg.txt" For Output As #o%
Print #o%, l$
Close #o%

End Sub
Sub chgappend(l$)
Dim o%

'd2infile = "Form1": d2insub = "chgappend"
o% = FreeFile
Open "sqlchg.txt" For Append As #o%
Print #o%, l$
Close #o%

End Sub

Public Sub dialme(nummer$)
Dim z$, i%, d$, n$

'd2infile = "Form1": d2insub = "dialme"
d$ = trm(nummer$)
Load AutoAnwahl
Call AutoAnwahl.SetFocus
i% = 1: n$ = ""
If Left(d$, 1) = "~" Then n$ = "~"
If Left(d$, 1) = "+" Then n$ = "00"
While i% <= Len(d$)
  z$ = Mid$(d$, i%, 1)
  If istnum(z$) > 0 Then
    n$ = n$ & z$
  End If
  i% = i% + 1
Wend
Call dbg2f("got " + nummer$ + "==>" + n$)
AutoAnwahl.nummer.text = n$
If form1.getusersetting("fb7050url", "") <> "" Then Call AutoAnwahl.cmdDial_Click

End Sub

Function ProvisionFuerAuftritt(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double
Dim typ$, stmp As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "ProvisionFuerAuftritt"
ProvisionFuerAuftritt = "0%"
thismwst = sys_mwst

If listenhauptperson <> "" Then
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
  rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aid$ & "' and feldname='" & listenhauptperson & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not stmp.EOF Then
    If trm(stmp!felddaten) <> "" Then
      ProvisionFuerAuftritt = ProvisionFuerAuftrittByAdr(aid$, trm(stmp!felddaten))
      Exit Function
    End If
  End If
End If
cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and feldname='Provision'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  hon$ = "0%"
  If Not IsNull(r!felddaten) Then
    hon$ = trm(r!felddaten)
    thismwst = var2dbl(trm(cut_d1(trm(cut_d2bis(hon$, "/")), "%"))) * 100
    hon$ = trm(cut_d1(hon$, "/"))
    ProvisionFuerAuftritt = hon$
  End If
End If
End Function

Function ProvisionFuerAuftrittByAdr(aufid$, adrid$) As String
Dim rrr, hon As Double
Dim stmp As ADODB.Recordset, ttmp As ADODB.Recordset, fnam$, j%, honfn$, honi%, atyp$, cmd$

Dim d2infile As String, d2insub As String, hsum As Double, psum As Double, c$, prov As Double
d2infile = "Form1": d2insub = "ProvisionFuerAuftrittByAdr"
ProvisionFuerAuftrittByAdr = ""

If adrid$ = "" Then
    
    cmd$ = "select FeldDaten from "
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      atyp$ = stmp!auftrittstyp
      If LCase(atyp$) = "künstlerauftritt" Or LCase(atyp$) = "promo" Or LCase(atyp$) = "orchesterauftritt" Or LCase(atyp$) = "dirigentenauftritt" Then
        hsum = 0
        For j% = 1 To sqla.TableDefs("usr_" & utabn(atyp$)).Fields.Count - 1
          fnam$ = LCase$(sqla.TableDefs("usr_" & utabn(atyp$)).Fields(j%).name)
          If InStr(LCase(fnam$), "honorar") = 1 Then
            cmd$ = "select * from usr_" & utabn(atyp$) & " where id='" & aufid$ & "'"
            Set stmp = New ADODB.Recordset
            stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            If Not stmp.EOF Then
              On Error Resume Next
              hon = CDbl(strrepl(ohnewaehrung(trm0(stmp.Fields(j%).value)), ".", ""))
              rrr = Err
              On Error GoTo 0
              If rrr <> 0 Then hon = 0
              hsum = hsum + hon
              If hon > 0 Then
              c$ = "": If isdigit(Right$(fnam$, 1)) Then c$ = Right$(fnam$, 1)
              cmd$ = "select FeldDaten from auftritthigru where auftrittsid='" & aufid$ & "' and FeldName='Provision" + c$ + "'"
              Set ttmp = New ADODB.Recordset
              ttmp.CursorLocation = adUseServer
rrr = form1.adoopen(ttmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
              If Not ttmp.EOF Then
                c$ = cut_d1(trm0(ttmp!felddaten), "/")
                If InStr(c$, "%") > 0 Then
                  prov = hon * CDbl(strrepl(c$, "%", "")) / 100
                  psum = psum + prov
                Else
                  On Error Resume Next
                  psum = psum + CDbl(c$)
                  rrr = Err
                  On Error GoTo 0
                  If rrr <> 0 Then
                    c$ = word1(c$)
                    On Error Resume Next
                    psum = psum + CDbl(c$)
                    rrr = Err
                    On Error GoTo 0
                    If rrr <> 0 And warnmeondata Then MsgBox ("could not handle this number: " + c$)
                  End If
                End If
              End If
              End If
            End If
          End If
        Next j%
        c$ = "0%"
        If hsum <> 0 Then
          'c$ = fixeur(psum * 100 / hsum) + "%"
          c$ = trm(psum * 100 / hsum) + "%"
        End If
        ProvisionFuerAuftrittByAdr = c$
        Exit Function
      End If
    End If

End If
    
    
    honi% = 0
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "' and felddaten='" & adrid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      atyp$ = stmp!auftrittstyp
      If LCase(atyp$) = "dirigentenauftritt" Then
        If InStr(LCase(honfn$), "dirigent") = 1 Then
          honfn$ = "provision" + onlynums(honfn$)
        Else
          honfn = ""
        End If
      Else
        If LCase(atyp$) = "orchesterauftritt" Then
          If honfn$ = "orchester" Then honfn$ = ""
        Else
          If InStr(LCase(honfn$), "künstler") = 1 Then
            honfn$ = "provision" + onlynums(honfn$)
          Else
            honfn = ""
          End If
        End If
      End If
      For j% = 1 To sqla.TableDefs("usr_" & utabn(atyp$)).Fields.Count - 1
        fnam$ = LCase$(sqla.TableDefs("usr_" & utabn(atyp$)).Fields(j%).name)
        If LCase(fnam$) = honfn$ Then
          honi% = j%
          cmd$ = "select * from usr_" & utabn(atyp$) & " where id='" & aufid$ & "'"
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If Not stmp.EOF Then
            ProvisionFuerAuftrittByAdr = trm("" & stmp.Fields(honi%).value)
            Exit Function
          End If
        End If
      Next j%
    End If

End Function

Function HonorarVonAuftrittByAdr(aufid$, adrid$) As String
Dim rrr
Dim stmp As ADODB.Recordset, fnam$, j%, honfn$, honi%, atyp$, cmd$, hsum As Double, c$, hon As Double

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "HonorarVonAuftrittByAdr"
HonorarVonAuftrittByAdr = ""

If adrid$ = "" Then
    hsum = 0
    j% = 0
    Do
      c$ = "": If j% > 0 Then c$ = trm(j%)
      c$ = "select * from auftritthigru where auftrittsid='" & aufid$ & "' and FeldName='Honorar" + c$ + "'"
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
      rrr = form1.adoopen(stmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If rrr = 0 Then
        If Not stmp.EOF Then
          On Error Resume Next
          hon = CDbl(ohnewaehrung(trm0(stmp!felddaten)))
          rrr = Err
          On Error GoTo 0
          If rrr <> 0 Then hon = 0
          hsum = hsum + hon
        End If
      End If
      j% = j% + 1
    Loop Until j% > 5 Or rrr <> 0
    HonorarVonAuftrittByAdr = fixeur(hsum)
    Exit Function
End If

    honi% = 0
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "' and felddaten='" & adrid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      atyp$ = stmp!auftrittstyp
      If LCase(atyp$) = "orchesterauftritt" Then
        If honfn$ = "orchester" Then honfn$ = ""
      Else
        If InStr(LCase(honfn$), "künstler") = 1 Then
          honfn$ = "honorar" + onlynums(honfn$)
        Else
          honfn = "honorar"
        End If
      End If
      For j% = 1 To sqla.TableDefs("usr_" & utabn(atyp$)).Fields.Count - 1
        fnam$ = LCase$(sqla.TableDefs("usr_" & utabn(atyp$)).Fields(j%).name)
        If LCase(fnam$) = honfn$ Then
          honi% = j%
          cmd$ = "select * from usr_" & utabn(atyp$) & " where id='" & aufid$ & "'"
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If Not stmp.EOF Then
            HonorarVonAuftrittByAdr = trm("" & stmp.Fields(honi%).value)
            Exit Function
          End If
        End If
      Next j%
    End If

End Function

Function MwStFuerAuftritt(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double
Dim typ$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "MwStFuerAuftritt"
MwStFuerAuftritt = "0"
cmd$ = "SELECT * FROM finanzen where id='" + aid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  hon$ = "0"
  If Not IsNull(r!mwst) Then
    hon$ = r!mwst
    MwStFuerAuftritt = hon$
  End If
End If
End Function

Function HonorarVonAuftritt(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double
Dim typ$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "HonorarVonAuftritt"
HonorarVonAuftritt = "0.00 " + transe("€")
cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and (feldname='Honorar' or feldname='Gesamtpreis')"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then
  cmd$ = "SELECT auftrittstyp FROM auftritt where id='" + aid$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  typ$ = ""
  If Not r.EOF Then typ$ = r!auftrittstyp
  If InStr(LCase(typ$), "auftritt") > 0 Or LCase(typ$) = "dienstleistung" Or LCase(typ$) = "verkauf" Then
    Unload auftritt
    DoEvents
    Load auftritt
    Call auftritt.SetFocus
    Call auftritt.showrec(aid$, 0)
    Load fdet
    Call fdet.SetFocus
    fdet.fid = aid$
    DoEvents
    Call fdet.Command10_Click
    DoEvents
    Unload auftritt
  End If
End If
r.Close
cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and (feldname='Honorar' or feldname='Gesamtpreis')"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  hon$ = "0.00 " + transe("€")
  If Not IsNull(r!felddaten) Then
    hon$ = r!felddaten
    HonorarVonAuftritt = hon$
  End If
  cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and feldname='Betrag_pro_Stunde'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    bps$ = "0.00 " + transe("€")
    If Not IsNull(r!felddaten) Then bps$ = r!felddaten
    cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and feldname='Dauer'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then
      dau$ = "0 h"
      If Not IsNull(r!felddaten) Then
        waehr$ = nurdiewaehrung(bps$)
        dau$ = r!felddaten
        dau$ = word1(dau$): wert2 = 0
        On Error Resume Next: wert2 = var2dbl(dau$): On Error GoTo 0
        bps$ = word1(bps$): wert1 = 0
        On Error Resume Next: wert1 = var2dbl(bps$): On Error GoTo 0
        h1on$ = Format$(wert1 * wert2, "0.00") + " " + waehr$
        If h1on$ <> hon$ Then
          cmd$ = "update auftritthigru set felddaten='" + h1on$ + "' where auftrittsid='" + aid$ + "' and feldname='Honorar'"
          Call sqlqry(cmd$)
          HonorarVonAuftritt = h1on$
        End If
      End If
    End If
  End If
End If
End Function

Function auftrittsstatus(aid$) As Integer
Dim rrr, iast As Integer
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittsstatus"
auftrittsstatus = -1
cmd$ = "SELECT astatus FROM auftritt where id='" + aid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
iast = -1
If Not IsNull(r!astatus) Then
  On Error Resume Next
  iast = r!astatus
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then iast = -1
End If
auftrittsstatus = iast

End Function

Function auftrittstyp(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittstyp"
auftrittstyp = ""
cmd$ = "SELECT auftrittstyp FROM auftritt where id='" + aid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
auftrittstyp = r!auftrittstyp

End Function

Function auftrittszeit(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittszeit"
auftrittszeit = ""
cmd$ = "SELECT zeit FROM auftritt where id='" + aid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
auftrittszeit = trm(r!zeit)

End Function

Sub c1add(w$)
Dim i%

'd2infile = "Form1": d2insub = "c1add"
For i% = 0 To Combo1.ListCount - 1
  If w$ = Combo1.List(i%) Then Exit Sub
Next i%
Combo1.AddItem w$

End Sub

Public Function get_atabkz(evid$) As String
Dim rrr
Dim at As ADODB.Recordset, i%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_atabkz"
get_atabkz = evid$
For i% = 0 To atabkzcachepointer% - 1
'  If atabkzcacheid$(i%) = evid$ Then
  If atabkz$(i%) = evid$ Then
    get_atabkz = atabkzcacheid$(i%)
    Exit Function
  End If
Next i%
Set at = New ADODB.Recordset
at.CursorLocation = adUseServer
rrr = form1.adoopen(at, "SELECT * FROM auftrittstypen where id='" + evid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If Not at.EOF Then
  If Not IsNull(at!abkz) Then
    atabkzcacheid$(atabkzcachepointer%) = at!abkz
    get_atabkz = atabkzcacheid$(atabkzcachepointer%)
    atabkzcachepointer% = atabkzcachepointer% + 1
    If atabkzcachepointer% > 99 Then atabkzcachepointer% = 99
  End If
End If

End Function

Public Function get_hordeabkz(evid$) As String
Dim typ As String

get_hordeabkz = evid$
typ = get_atabkz(evid$)
get_hordeabkz = getusersetting("hordetyp_" + typ, typ)

End Function


Public Function get_projectid_by_aid(a$) As String
Dim rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_projectid_by_aid"
get_projectid_by_aid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT TourneeplanID as rc FROM auftritt where id='" + a$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!rc) Then Exit Function
get_projectid_by_aid = rtmp!rc

End Function


Public Sub ExportOneTableToExcel(tn$)
Dim rrr
Dim i%

'EXPORTS TABLE IN ACCESS DATABASE TO EXCEL
'REFERENCE TO DAO IS REQUIRED

Dim strExcelFile As String
Dim strTxtFile As String
Dim strWorksheet As String
Dim strTable As String
Dim c$, r As ADODB.Recordset, o%
'Dim objDB As Database

'Change Based on your needs, or use
'as parameters to the sub
strExcelFile = tn$ + ".xls"
strTxtFile = tn$ + ".txt"
strWorksheet = Left$(tn$, 8)
strTable = tn$

'Set objDB = OpenDatabase(strDB)

' 'If excel file already exists, you can delete it here
If Dir(strExcelFile) <> "" Then Kill strExcelFile

Call sqlqry( _
  "SELECT * INTO [Excel 8.0;DATABASE=" & strExcelFile & _
   "].[" & strWorksheet & "] FROM " & "[" & strTable & "]" _
   )
c$ = "select * from " & strTable
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
o% = FreeFile
Open "xp_" & strTxtFile For Output As #o%
While Not r.EOF
  For i% = 0 To r.Fields.Count - 1
    Print #o%, trm("'" & r.Fields(i%).value); "',";
  Next i%
  Print #o%,
  r.MoveNext
Wend
Close #o%

End Sub

Function wordsuchen() As String
Dim cdn$, p%, fn$

'd2infile = "Form1": d2insub = "wordsuchen"
wordsuchen = ""

p% = FreeFile
fn$ = mylocaldatadir() + "such.dir"
Open fn$ For Output As #p%
Print #p%, "c:\"
Print #p%, "d:\"
Close #p%

While wordsuchen = "" And exist(fn$) = 1
  wordsuchen = fsuchrekursiv("winword.exe")
Wend

End Function

Function fsuchrekursiv(fn$) As String
Dim tr As String, o%, cdn$, lcnt%, p%, rrr, erg

'd2infile = "Form1": d2insub = "fsuchrekursiv"
fsuchrekursiv = ""
o% = FreeFile
fn$ = mylocaldatadir() + "such.dir"
On Error Resume Next
Kill fn$ + ".bak"
Name fn$ As fn$ + ".bak"
On Error GoTo 0
lcnt% = 0
Open fn$ + ".bak" For Input As #o%
Line Input #o%, cdn$
'Debug.Print "fsuchrekuriv in "; cdn$; " ";
dbupgrade.List1.AddItem cdn$
dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1
DoEvents
If exist(cdn$ + "winword.exe") = 1 Then
  fsuchrekursiv = cdn$ + "winword.exe"
  Close #p%
  Close #o%
  On Error Resume Next
  Kill fn$
  Kill fn$ + ".bak"
  On Error GoTo 0
  Exit Function
End If

p% = FreeFile
Open fn$ For Output As #p%
On Error Resume Next
tr = Dir(cdn$, vbDirectory)
rrr = Err
On Error GoTo 0
Do While tr <> "" And rrr = 0
  On Error Resume Next
  erg = (GetAttr(cdn$ + tr) And vbDirectory)
  rrr = Err
  On Error GoTo 0
  If rrr = 0 And erg = vbDirectory Then
    If Left(tr, 1) <> "." Then
      lcnt% = lcnt% + 1
      Print #p%, cdn$ & tr & "\"
    End If
  End If
  tr = Dir()
Loop
While Not EOF(o%)
  Line Input #o%, cdn$
  Print #p%, cdn$
  lcnt% = lcnt% + 1
Wend
Close #p%
Close #o%
'Debug.Print lcnt%
If lcnt% = 0 Then
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
End If

End Function
Public Function defaultbesetzt(id) As String
Dim rrr
Dim r As ADODB.Recordset
Dim rc As String

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "defaultbesetzt"
defaultbesetzt = ""
If IsNull(id) Or id = "NULL" Then Exit Function
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id FROM b_loc where wid='" & id & "' and dflt=1", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  rc = bestzstr(r!id)
  defaultbesetzt = Mid(rc, InStr(rc, ":") + 1)
End If

End Function
Public Function bestzstr(id) As String
Dim rrr
Dim r As ADODB.Recordset
Dim i%, p%, rc As String

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "bestzstr"
bestzstr = "": rc = ""
If IsNull(id) Or id = "NULL" Then Exit Function
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM b_loc where id='" & id & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  For p% = 2 To r.Fields.Count - 2
    If rc <> "" Then rc = rc & "/"
    rc = rc & trm(r.Fields(p%).value)
    If InStr(trm(r.Fields(p%).name), "sonst") > 0 Then
      bestzstr = bestzstr & "  " & rc
      rc = ""
    End If
  Next p%
  If rc <> "" Then bestzstr = bestzstr & "  " & rc
  bestzstr = trm(bestzstr)
  If Not IsNull(r!dflt) And r!dflt > 0 Then bestzstr = "Standard:" & bestzstr
End If

End Function

Public Function neumaxbesetz(mb$, b$) As String
Dim g$(4), s$, rr$(4), i%, wa$, rc$, n%
Dim mg$(4), mrr$(4)

'd2infile = "Form1": d2insub = "neumaxbesetz"
If Left(b$, 9) = "Standard:" Then b$ = trm(Mid$(b$, 10))

neumaxbesetz = mb$
If trm(mb$) = "" Then
  neumaxbesetz = b$
  Exit Function
End If

i% = 0
s$ = b$
While Len(trm(s$)) > 0
  g$(i%) = word1(s$)
  s$ = trm(Mid$(s$, Len(g$(i%)) + 1))
  While Len(trm(s$)) > 0 And InStr(word1(s$), "/") = 0
    wa$ = word1(s$)
    rr$(i%) = rr$(i%) + " " + wa$
    s$ = trm(Mid$(s$, Len(wa$) + 1))
  Wend
  i% = i% + 1
Wend
i% = 0
s$ = mb$
While Len(trm(s$)) > 0
  mg$(i%) = word1(s$)
  s$ = trm(Mid$(s$, Len(mg$(i%)) + 1))
  While Len(trm(s$)) > 0 And InStr(word1(s$), "/") = 0
    wa$ = word1(s$)
    mrr$(i%) = mrr$(i%) + " " + wa$
    s$ = trm(Mid$(s$, Len(wa$) + 1))
  Wend
  i% = i% + 1
Wend
rc$ = ""
For i% = 0 To 3
  If InStr(g$(i%), "/") > 0 Then
    n% = Len(g$(i%))
    While n% > 0 And Mid$(g$(i%), n%, 1) <> "/": n% = n% - 1: Wend
    rr$(i%) = trm(Mid$(g$(i%), n% + 1) + " " + rr$(i%))
    g$(i%) = Left$(g$(i%), n% - 1)
  End If
  If InStr(mg$(i%), "/") > 0 Then
    n% = Len(mg$(i%))
    While n% > 0 And Mid$(mg$(i%), n%, 1) <> "/": n% = n% - 1: Wend
    mrr$(i%) = trm(Mid$(mg$(i%), n% + 1) + " " + mrr$(i%))
    mg$(i%) = Left$(mg$(i%), n% - 1)
  End If
'Debug.Print g$(i%); "<==>"; mg$(i%);
  mg$(i%) = neumaxbesetzgrp(g$(i%), mg$(i%))
'Debug.Print "==>"; mg$(i%)
  rr$(i%) = trm(strrepl(rr$(i%), "-", ""))
'Debug.Print rr$(i%); "+"; mrr$(i%); "==>";
  If trm(rr$(i%)) <> "" Then rr$(i%) = rr$(i%) & ","
  mrr$(i%) = trm(rr$(i%) & " " & strrepl(mrr$(i%), "-", ""))
'Debug.Print mrr$(i%)
'Debug.Print
  If rc$ <> "" Then rc$ = rc$ + " "
  rc$ = rc$ + trm((mg$(i%) + "/" + mrr$(i%))): If Right$(rc$, 1) = "/" Then rc$ = rc$ + "-"
'Debug.Print "rc="; rc$
Next i%
neumaxbesetz = rc$
End Function

Private Function neumaxbesetzgrp(b$, mb$) As String
Dim i%, rc$, a$, l$, b0%, m0%

'd2infile = "Form1": d2insub = "neumaxbesetzgrp"
If trm(b$) = "" Then
  neumaxbesetzgrp = mb$
  Exit Function
End If
If trm(mb$) = "" Then
  neumaxbesetzgrp = b$
  Exit Function
End If
a$ = b$: l$ = mb$: rc$ = ""
While a$ <> "" And l$ <> ""
  b0% = 0: If a$ <> "" Then b0% = Val(a$)
  m0% = 0: If l$ <> "" Then m0% = Val(l$)
  If b0% > m0% Then m0% = b0%
  If rc$ <> "" Then rc$ = rc$ + "/"
  rc$ = rc$ + trm(m0%)
  i% = InStr(a$, "/")
  If i% > 0 Then
    a$ = Mid$(a$, i% + 1)
  Else
    a$ = ""
  End If
  i% = InStr(l$, "/")
  If i% > 0 Then
    l$ = Mid$(l$, i% + 1)
  Else
    l$ = ""
  End If
Wend

neumaxbesetzgrp = rc$
End Function
Public Function mylastFormVar(frm$, var$, def$) As String
Dim inifile As String, l$, o%

'd2infile = "Form1": d2insub = "mylastFormVar"
l$ = def$
inifile = form1.mylocaldatadir() + "\positions\" + frm$ + "." & var$
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  If Not EOF(o%) Then
    Line Input #o%, l$
  End If
  Close #o%
End If
On Error Resume Next
mylastFormVar = l$
On Error GoTo 0

End Function
Public Sub setmylastFormVar(f$, V$, wert$)
Dim inifile As String, o%, rrr


'd2infile = "Form1": d2insub = "setmylastFormVar"
inifile = form1.mylocaldatadir() + "\positions\" + f$ + "." & V$
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  Print #o%, wert$
  Close #o%
End If
End Sub

Function HinweisVonAuftritt(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "HinweisVonAuftritt"
HinweisVonAuftritt = ""
cmd$ = "SELECT felddaten FROM auftritthigru where auftrittsid='" + aid$ + "'" & _
   " and ( instr(feldname,'Hinweis')>0 " & _
   " or feldname='Nachricht' " & _
   " or instr(feldname,'Anmerkun')>0 " & _
   " or feldname='Bemerkung' )"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then HinweisVonAuftritt = r!felddaten

End Function

Sub reftest(wert)
Dim i As Integer, j As Integer, rrr
Dim s As Database, stn$, fn$
Dim r As ADODB.Recordset, t As TableDef, tn$, f As Field
Dim idx As New Index

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "reftest"
fn$ = "referenz.mdb"
If exist(fn$) = 0 Then Exit Sub
If InStr(LCase(dbname$), ".mdb") = 0 Then Exit Sub
On Error Resume Next
Set s = OpenDatabase(fn$, True, False)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
Debug.Print s.name; " mit"; s.TableDefs.Count - 1; " Tabellen"
For i = 0 To s.TableDefs.Count - 1
  If Left$(s.TableDefs(i).name, 4) <> "MSys" Then
    On Error Resume Next
    Debug.Print s.TableDefs(i).name; " und "; sqla.TableDefs(s.TableDefs(i).name).name
    rrr = Err
    On Error GoTo 0
    If rrr = 3265 Then
      Set t = sqla.CreateTableDef(s.TableDefs(i).name)
      For j = 0 To s.TableDefs(i).Fields.Count - 1
        Set f = t.CreateField(s.TableDefs(i).Fields(j).name, s.TableDefs(i).Fields(j).Type)
        f.Size = s.TableDefs(i).Fields(j).Size:
        t.Fields.Append f
      Next j
      sqla.TableDefs.Append t
    Else
      For j = 0 To s.TableDefs(i).Fields.Count - 1
        If wert = 1 Then
          If j = 0 Then
            idx.name = "PrimaryKey"
            idx.Fields = s.TableDefs(i).Fields(j).name
            idx.Primary = True
            idx.Unique = True
            s.TableDefs(i).Indexes.Append idx
          End If
        End If
        On Error Resume Next
        tn$ = sqla.TableDefs(s.TableDefs(i).name).Fields(j).name
        rrr = Err
        On Error GoTo 0
        If rrr = 3265 Then
          tn$ = s.TableDefs(i).Fields(j).name
          Set t = sqla.TableDefs(s.TableDefs(i).name)
          Set f = t.CreateField(tn$, s.TableDefs(i).Fields(j).Type, s.TableDefs(i).Fields(j).Size)
          t.Fields.Append f
        Else
          If s.TableDefs(i).Fields(j).name <> tn$ Then
            Debug.Print tn$; ":"; s.TableDefs(i).Fields(j).name; " <--> "; sqla.TableDefs(s.TableDefs(i).name).Fields(j).name
          End If
          If s.TableDefs(i).Fields(j).Type <> sqla.TableDefs(s.TableDefs(i).name).Fields(j).Type Then
            Debug.Print tn$; ":"; s.TableDefs(i).Fields(j).Type; " <--> "; sqla.TableDefs(s.TableDefs(i).name).Fields(j).Type
          End If
          If s.TableDefs(i).Fields(j).Size <> sqla.TableDefs(s.TableDefs(i).name).Fields(j).Size Then
            Call errhdl(s.TableDefs(i).name & ", " & tn$ & ":" & s.TableDefs(i).Fields(j).Size & " <--> " & sqla.TableDefs(s.TableDefs(i).name).Fields(j).Size)
            Set f = sqla.TableDefs(s.TableDefs(i).name).Fields(j)
          End If
        End If
      Next j
    End If
  End If
Next i

End Sub
Public Function getsystemsetting(fldn$, Optional vifnull As String) As String
Dim r As ADODB.Recordset, vin As String, rrr, c$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getsystemsetting"
vin = ""
If vifnull <> "" Then vin = vifnull
getsystemsetting = vin

c$ = "SELECT wert as rc FROM sysvars where owner='sysvar_system_" & fldn$ & "'"
If isstarting And fldn$ <> "localdir" Then Call startlog(getuserid(), "getsystemsetting:" + trm(c$))
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Function
If Not r.EOF Then
  If Not IsNull(r!rc) Then getsystemsetting = internaldecrypt(trm(r!rc))
  If isstarting And fldn$ <> "localdir" And InStr(LCase(fldn$), "passw") = 0 Then Call startlog(getuserid(), "getsystemsetting returns:" + trm(r!rc))
  Exit Function
End If
If isstarting And InStr(LCase(fldn$), "passw") = 0 And fldn$ <> "localdir" Then Call startlog(getuserid(), "getsystemsetting returns:" + trm(getsystemsetting))
End Function

Public Function dictionarylookup(w$) As String
Dim rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "dictionarylookup"
dictionarylookup = w$

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM dictionary where id='" + w$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then
  dictionarylookup = transe(w$)
  Exit Function
End If
If IsNull(rtmp!translat) Then
  dictionarylookup = transe(w$)
  Exit Function
End If
dictionarylookup = trm(rtmp!translat)
rtmp.Close
End Function

Public Function dictionarylookupmonth(i%) As String
Dim rrr
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "dictionarylookupmonth"
dictionarylookupmonth = MonthName$(i%)

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT translat FROM dictionary where id='" + MonthName$(i%) + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If IsNull(rtmp!translat) Then Exit Function
dictionarylookupmonth = trm(rtmp!translat)
rtmp.Close
End Function

Public Function engnameof(na$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "engnameof"
engnameof = na$

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresse where name='" + na$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Function
If IsNull(rtmp!id) Then Exit Function
cmd$ = "SELECT Felddaten From auftritthigru where auftrittstyp='Daten-Engl' and auftrittsid='" & rtmp!id & "' and feldname='Name'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
engnameof = trm(r!felddaten)
End Function

Public Function myfontsize() As Integer

'd2infile = "Form1": d2insub = "myfontsize"
If ufsze% < 8 Or ufsze% > 12 Then ufsze% = 8
myfontsize = ufsze%

End Function
Public Function getmyhomepath()
Dim uoutlk$

'd2infile = "Form1": d2insub = "getmyhomepath"
If uhomepath$ <> "" Then
  getmyhomepath = uhomepath$
  Exit Function
End If
getmyhomepath = form1.mydatadir()

End Function

Public Function getmyoutlook()
Dim uoutlk$

'd2infile = "Form1": d2insub = "getmyoutlook"
uoutlk$ = form1.getusersetting("outlook", "")
If exist(uoutlk$) = 0 Then
  uoutlk$ = ""
End If
getmyoutlook = uoutlk$

End Function

Public Function mkupdcmd(t$, idn$, id$, n$, typ%, wert)
Dim c$, i%

'd2infile = "Form1": d2insub = "mkupdcmd"
c$ = ""
If LCase(n$) <> "tstamp" Then
  Select Case sqla.TableDefs(t$).Fields(i%).Type
        Case 8: c$ = "update " & t$ & " set " & n$ & "='" & wert & "' where " & idn$ & "='" & id$ & "'"
        Case 10: c$ = "update " & t$ & " set " & n$ & "='" & wert & "' where " & idn$ & "='" & id$ & "'"
        Case 12: c$ = "update " & t$ & " set " & n$ & "='" & wert & "' where " & idn$ & "='" & id$ & "'"
        Case Else:
               c$ = typ% & " " & wert
  End Select
End If
mkupdcmd = c$
End Function
Public Sub sqlex_adresse(t$, idn$, id$)
Dim rrr, rhig As ADODB.Recordset, j%
Dim c$, r As ADODB.Recordset, o%, fn$, r1 As ADODB.Recordset, r2 As ADODB.Recordset, i%
Dim kontaktexport As Boolean, hinweisexport As Boolean

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "sqlex_adresse"
If t$ = "usr_neuer auftritt" Then Exit Sub
fn$ = form1.mydatadir() & "\" & t$ & "_" & mkfn(id$) & ".sql"
If exist(fn$) = 0 Then
  kontaktexport = True
  hinweisexport = True
  If getusersetting("adressekontaktexport", "nein") <> "ja" Then kontaktexport = False
  If getusersetting("adressehinweisexport", "nein") <> "ja" Then hinweisexport = False
  c$ = "select * from " & t$ & " where " & idn$ & "='" & id$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    o% = FreeFile
    Open fn$ For Output As #o%
    c$ = "insert into " & t$ & " (" & idn$ & ") values('" & r.Fields(0).value & "');"
    Print #o%, c$
    For i% = 1 To r.Fields.Count - 1
      If trm(r.Fields(i%).value) <> "" Then
        If hinweisexport Or (Not hinweisexport And InStr(LCase(r.Fields(i%).name), "hinweis") = 0) Then
          c$ = mkupdcmd(t$, idn$, id$, r.Fields(i%).name, r.Fields(i%).Type, r.Fields(i%).value) & ";"
          Print #o%, c$
        End If
      End If
    Next i%
    If t$ = "hblist" Then
      c$ = "select * from hbplist where hid='" & r!hid & "' and pgid='" & r!pgid & "'"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r1.EOF
        c$ = "insert into hbplist (id) values('" & r1!id & "');"
        Print #o%, c$
        For i% = 1 To r1.Fields.Count - 1
          If trm(r1.Fields(i%).value) <> "" Then
            c$ = mkupdcmd("hbplist", "id", r1!id, r1.Fields(i%).name, r1.Fields(i%).Type, r1.Fields(i%).value) & ";"
            Print #o%, c$
          End If
        Next i%
        r1.MoveNext
      Wend
    End If
    If t$ = "adresse" Then
      c$ = "select * from adresstyp where vid='" & id$ & "' and (kid='-1' or isnull(kid))"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r1.EOF
        c$ = "insert into adresstypen (id,sortierung) values('" & r1!typ & "',999);"
        Print #o%, c$
        c$ = "insert into adresstyp (id,kid) values('" & r1!id & "','-1');"
        Print #o%, c$
        For i% = 1 To r1.Fields.Count - 1
          If trm(r1.Fields(i%).value) <> "" _
             And LCase(r1.Fields(i%).name) <> "kid" Then
            c$ = mkupdcmd("adresstyp", "id", r1!id, r1.Fields(i%).name, r1.Fields(i%).Type, r1.Fields(i%).value) & ";"
            Print #o%, c$
          End If
        Next i%
        c$ = "select * from auftritthigru where auftrittsid='" & id$ & "' and auftrittstyp='" + r1!typ + "'"
        Set rhig = New ADODB.Recordset
        rhig.CursorLocation = adUseServer
rrr = form1.adoopen(rhig, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not rhig.EOF
          c$ = "insert into auftritthigru (id) values('" & rhig!id & "');"
          Print #o%, c$
          For j% = 1 To rhig.Fields.Count - 1
            If trm(rhig.Fields(j%).value) <> "" Then
              c$ = mkupdcmd("auftritthigru", "id", rhig!id, rhig.Fields(j%).name, rhig.Fields(j%).Type, rhig.Fields(j%).value) & ";"
              Print #o%, c$
            End If
          Next j%
          rhig.MoveNext
        Wend
        r1.MoveNext
      Wend
      If hinweisexport Then
        c$ = "select * from auftritthigru where auftrittsid='" & id$ & "'"
        Set r1 = New ADODB.Recordset
        r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not r1.EOF
          c$ = "insert into auftritthigru (id) values('" & r1!id & "');"
          Print #o%, c$
          For i% = 1 To r1.Fields.Count - 1
            If trm(r1.Fields(i%).value) <> "" Then
              c$ = mkupdcmd("auftritthigru", "id", r1!id, r1.Fields(i%).name, r1.Fields(i%).Type, r1.Fields(i%).value) & ";"
              Print #o%, c$
            End If
          Next i%
          r1.MoveNext
        Wend
      End If
      If kontaktexport Then
        c$ = "select * from kontakt where vid='" & id$ & "'"
        Set r2 = New ADODB.Recordset
        r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not r2.EOF
          c$ = "insert into kontakt (id) values('" & r2!id & "');"
          Print #o%, c$
          For i% = 1 To r2.Fields.Count - 1
            If trm(r2.Fields(i%).value) <> "" Then
              c$ = mkupdcmd("kontakt", "id", r2!id, r2.Fields(i%).name, r2.Fields(i%).Type, r2.Fields(i%).value) & ";"
              Print #o%, c$
            End If
          Next i%
          c$ = "select * from adresstyp where vid='" & id$ & "' and kid='" & r2!id & "'"
          Set r1 = New ADODB.Recordset
          r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not r1.EOF
            c$ = "insert into adresstyp (id) values('" & r1!id & "');"
            Print #o%, c$
            For i% = 1 To r1.Fields.Count - 1
              If trm(r1.Fields(i%).value) <> "" Then
                c$ = mkupdcmd("adresstyp", "id", r1!id, r1.Fields(i%).name, r1.Fields(i%).Type, r1.Fields(i%).value) & ";"
                Print #o%, c$
              End If
            Next i%
            r1.MoveNext
          Wend
          c$ = "select * from auftritthigru where auftrittsid='" & id$ & r2!id & "'"
          Set r1 = New ADODB.Recordset
          r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not r1.EOF
            c$ = "insert into auftritthigru (id) values('" & r1!id & "');"
            Print #o%, c$
            For i% = 1 To r1.Fields.Count - 1
              If trm(r1.Fields(i%).value) <> "" Then
                c$ = mkupdcmd("auftritthigru", "id", r1!id, r1.Fields(i%).name, r1.Fields(i%).Type, r1.Fields(i%).value) & ";"
                Print #o%, c$
              End If
            Next i%
            r1.MoveNext
          Wend
          r2.MoveNext
        Wend
      End If
    End If
    Close #o%
  End If
End If

End Sub

Public Function abosimblock(halle$, block$, dtg$) As Integer
Dim rrr
'd2infile = "Form1": d2insub = "abosimblock"
abosimblock = 0

Dim c$, r As ADODB.Recordset

c$ = "SELECT count(*) as cnt " & _
     "FROM ((hbpstatus INNER JOIN (hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid) ON (hbpstatus.dtg = hbabotermine.dtg) AND (hbpstatus.pstatus2 = hbabotermine.aboid)) INNER JOIN hbplist ON (hbabotermine.adrid = hbplist.hid) AND (hbabotermine.pid = hbplist.pgid) AND (hbpstatus.hbpid = hbplist.id)) INNER JOIN hblist ON (hblist.pgid = hbplist.pgid) AND (hbplist.hid = hblist.hid) " & _
     "WHERE (((hblist.hid)='" & halle$ & "') AND ((hblist.pgid)='" & block$ & "') AND ((hbpstatus.dtg)='" & dtg$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
abosimblock = r!cnt

End Function
Public Function abosimraum(halle$, raum$, dtg$) As Integer
Dim rrr
'd2infile = "Form1": d2insub = "abosimraum"
abosimraum = 0

Dim c$, r As ADODB.Recordset

c$ = "SELECT count(*) as cnt " & _
     "FROM ((hbpstatus INNER JOIN (hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid) ON (hbpstatus.pstatus2 = hbabotermine.aboid) AND (hbpstatus.dtg = hbabotermine.dtg)) INNER JOIN hbplist ON (hbpstatus.hbpid = hbplist.id) AND (hbabotermine.pid = hbplist.pgid) AND (hbabotermine.adrid = hbplist.hid)) INNER JOIN hblist ON (hbplist.hid = hblist.hid) AND (hbplist.pgid = hblist.pgid) " & _
     "WHERE (((hblist.raum)='" & raum$ & "') AND ((hblist.hid)='" & halle$ & "') AND ((hbpstatus.dtg)='" & dtg$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
abosimraum = r!cnt

End Function


Public Sub kassenabrechnung(thema$, von$, bis$)
Dim rrr
Dim o%, c$, vorlage$, p%, fn$, l$, q%, t$, ln$, pb%
Dim epn As Double, epm As Double, epb As Double
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "kassenabrechnung"
vorlage$ = "kvk_abrechnung.rtf"
If vorlage$ = "" Then vorlage$ = meineprgdruckvorlage()
If exist(s0d$ & "\" + dbname$ + ".rtf\" & vorlage$) = 0 Then
  MsgBox "Vorlage unbekannt: " + s0d$ & "\" & vorlage$
  Exit Sub
End If

epn = 0: epb = 0: epm = 0
o% = FreeFile
Open s0d$ & "\" & dbname$ + ".rtf\" & vorlage$ For Input As #o%
p% = FreeFile
fn$ = trm(form1.myuniquedocname("noask"))
fn$ = DirName(fn$) & "\kassenzettel.rtf"
If fn$ = "" Then Exit Sub
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  q% = InStr(l$, "Liste der Auftritte:")
  If q% > 0 Then
    c$ = "SELECT * FROM kassenbuch where ((thema='" & thema$ + "') and (dtg>='" & von$ + "')  and (dtg<='" & bis$ + "')) order by dtg;"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
      Print #p%, "\trowd \trgaph70\trleft-70 \cellx1834\cellx6987\cellx7337\cellx8213\cellx9141 \pard\plain \intbl \f4\fs20\lang1031 ";
      Print #p%, r!dtg & "\cell " & r!bezeichnung & "\cell \pard \qr\intbl ";
      Print #p%, "1\cell " & fixeur(r!epreisnetto) & "\cell " & fixeur(r!mwst) & "\cell \pard \intbl \row \pard ";
'      Print #p%, "\trowd \trgaph70\trleft-70 \cellx2410\cellx9476 \pard \intbl "
      epn = epn + r!epreisnetto
      epm = epm + r!mwst
      epb = epb + r!epreisnetto + r!mwst
      r.MoveNext
    Wend
  Else
    While Len(l$) > 0
      q% = InStr(l$, bkmstart$)
      If q% > 0 Then
        t$ = Mid$(l$, q% + Len(bkmstart$))
        Print #p%, Left$(l$, q% - 1)
        t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
        Select Case t$
          Case "bezeichnung": Print #p%, thema$
          Case "von": Print #p%, datfromsql(word1(von$)) & " " & word2(von$)
          Case "enddatum": Print #p%, datfromsql(word1(bis$)) & " " & word2(bis$)
          Case "system__datum": Print #p%, Date
          Case "summe_honorar1_netto": Print #p%, fixeur(epn)
          Case "summe_honorar1_mwst": Print #p%, fixeur(epm)
          Case "summe_honorar1_brutto": Print #p%, fixeur(epb)
          Case Else
        End Select
        ln$ = Mid$(l$, q% + 1)
        Do
            pb% = InStr(ln$, bkmend$ + t$)
            If pb% = 0 Then Line Input #o%, ln$
        Loop Until pb% > 0
        ln$ = Mid$(ln$, pb%)
        If InStr(ln$, "}") = 0 Then
            l$ = ""
        Else
            l$ = Mid$(ln$, InStr(ln$, "}") + 1)
        End If
      Else
        Print #p%, l$
        l$ = ""
      End If
    Wend
  End If

Wend
Close #o%
Close #p%

Call form1.openthisdoc(fn$, "")

End Sub

Public Sub kassenzettel(thema$, von$, bis$)
Dim rrr
Dim o%, c$, vorlage$, fn$, trgp$, l$, q%, t$, ln$, pb%
Dim epn As Double, epm As Double, epb As Double, kn$, ks$, ko$, p%, po%
Dim r As ADODB.Recordset, aboid$, aboplid$, r1 As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "kassenzettel"
vorlage$ = "kvk_rechnung.rtf"
If vorlage$ = "" Then vorlage$ = meineprgdruckvorlage()
If exist(s0d$ & "\" + dbname$ + ".rtf\" & vorlage$) = 0 Then
  MsgBox "Vorlage unbekannt: " + s0d$ & "\" & vorlage$
  Exit Sub
End If

epn = 0: epb = 0: epm = 0
o% = FreeFile
Open s0d$ & "\" & dbname$ + ".rtf\" & vorlage$ For Input As #o%
p% = FreeFile
fn$ = trm(form1.myuniquedocname("noask"))
fn$ = DirName(fn$) & "\kassenzettel.rtf"
If fn$ = "" Then Exit Sub
kn$ = "": ks$ = "": ko$ = ""
c$ = "SELECT * FROM kassenbuch where ((thema='" & thema$ + "') and (dtg>='" & von$ + "')  and (dtg<='" & bis$ + "')) order by dtg;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  If trm(r!kontaktname) <> "" And r!kontaktname <> "-1" Then kn$ = r!kontaktname
  c$ = "SELECT * FROM adresse where ((id='" & r!vonid & "'));"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    trgp$ = form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(r!id)
    fn$ = trm(form1.myuniquedocnameinpath(trgp$, ""))
c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,doctyp) values('" & _
              form1.newid("dochist", "id", 19) & "','" & _
              r!id & "','" & r!name & "','" & _
              fn$ & "','" & _
              Date & " " & Time & "','" & _
              uId$ & "','" & _
              "Kartenverkauf" & " ','" & _
              "Kassenzettel" & "')"
Call form1.sqlqry(c$)
    If kn$ = "" Then
      kn$ = r!name
    Else
      kn$ = r!name & "\par " & kn$
    End If
    ks$ = trm(r!strasse)
    ko$ = trm(trm(r!plz) & " " & r!ort)
  End If
End If
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  q% = InStr(l$, "Liste der Auftritte:")
  If q% > 0 Then
    Print #p%, "\trowd \trgaph70\trleft-70 \cellx1834\cellx6987\cellx7337\cellx8213\cellx9141 \pard\plain \intbl \f4\fs20\lang1031 ";
    Print #p%, "Kaufdatum" & "\cell " & "für Datum und Platz" & "\cell \pard \qr\intbl ";
    Print #p%, " \cell Netto \cell MwSt\cell \pard \intbl \row \pard ";
    c$ = "SELECT * FROM kassenbuch where ((thema='" & thema$ + "') and (dtg>='" & von$ + "')  and (dtg<='" & bis$ + "')) order by dtg;"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not r.EOF
      Print #p%, "\trowd \trgaph70\trleft-70 \cellx1834\cellx6987\cellx7337\cellx8213\cellx9141 \pard\plain \intbl \f4\fs20\lang1031 ";
      Print #p%, r!dtg & "\cell " & r!bezeichnung & "\cell \pard \qr\intbl ";
      Print #p%, "1\cell " & fixeur(r!epreisnetto) & "\cell " & fixeur(r!mwst) & "\cell \pard \intbl \row \pard ";
'      Print #p%, "\trowd \trgaph70\trleft-70 \cellx2410\cellx9476 \pard \intbl "
      epn = epn + r!epreisnetto
      epm = epm + r!mwst
      epb = epb + r!epreisnetto + r!mwst
      c$ = trm(r!zahlstatus)
      If Len(c$) > 1 Then
        po% = InStr(c$, "/")
        If po% > 1 Then
          aboid$ = Left$(c$, po% - 1)
          aboplid$ = Mid$(c$, po% + 1)
          c$ = "SELECT hbplist.platzname, hbabotermine.dtg, hbplist.hid, hbplist.pgid, hbabotermine.aboid, hbpstatus.aboplatzid " & _
               "FROM (hbabotermine INNER JOIN hbpstatus ON (hbabotermine.dtg = hbpstatus.dtg) AND (hbabotermine.aboid = hbpstatus.pstatus2)) INNER JOIN hbplist ON (hbpstatus.hbpid = hbplist.id) AND (hbabotermine.pid = hbplist.pgid) AND (hbabotermine.adrid = hbplist.hid) " & _
               "WHERE (((hbabotermine.aboid)='" & aboid$ & "') AND ((hbpstatus.aboplatzid)=" & aboplid$ & "));"
          Set r1 = New ADODB.Recordset
          r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not r1.EOF
            Print #p%, "\trowd \trgaph70\trleft-70 \cellx1834\cellx6987\cellx7337\cellx8213\cellx9141 \pard\plain \intbl \f4\fs20\lang1031 ";
            Print #p%, " " & "\cell " & r1!dtg & " " & r1!hid & " " & r1!pgid & " " & r1!platzname & "\cell \pard \qr\intbl ";
            Print #p%, " \cell " & " " & "\cell " & " " & "\cell \pard \intbl \row \pard ";
            r1.MoveNext
          Wend
        End If
      End If
      r.MoveNext
    Wend
  Else
    While Len(l$) > 0
      q% = InStr(l$, bkmstart$)
      If q% > 0 Then
        t$ = Mid$(l$, q% + Len(bkmstart$))
        Print #p%, Left$(l$, q% - 1)
        t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
        Select Case t$
          Case "bezeichnung": Print #p%, thema$
          Case "von": Print #p%, datfromsql(word1(von$)) & " " & word2(von$)
          Case "enddatum": Print #p%, datfromsql(word1(bis$)) & " " & word2(bis$)
          Case "system__datum": Print #p%, Date
          Case "kunde__name": Print #p%, kn$
          Case "kunde__strasse": Print #p%, ks$
          Case "kunde__ort": Print #p%, ko$
          Case "summe_honorar1_netto": Print #p%, fixeur(epn)
          Case "summe_honorar1_mwst": Print #p%, fixeur(epm)
          Case "summe_honorar1_brutto": Print #p%, fixeur(epb)
          Case Else
        End Select
        ln$ = Mid$(l$, q% + 1)
        Do
            pb% = InStr(ln$, bkmend$ + t$)
            If pb% = 0 Then Line Input #o%, ln$
        Loop Until pb% > 0
        ln$ = Mid$(ln$, pb%)
        If InStr(ln$, "}") = 0 Then
            l$ = ""
        Else
            l$ = Mid$(ln$, InStr(ln$, "}") + 1)
        End If
      Else
        Print #p%, l$
        l$ = ""
      End If
    Wend
  End If
Wend
Close #o%
Close #p%

Call form1.openthisdoc(fn$, "")

End Sub

Public Function mkfn(fnin$) As String
Dim rc$, i%, z$

'd2infile = "Form1": d2insub = "mkfn"
mkfn = "unbenannt"
rc$ = ""
For i% = 1 To Len(fnin$)
  z$ = LCase(Mid$(fnin$, i%, 1))
  If (z$ >= "0" And z$ <= "9") Or (z$ >= "a" And z$ <= "z") Then rc$ = rc$ & z$
Next i%
If rc$ <> "" Then mkfn = rc$

End Function
Public Function mkkompdn(lfnin$) As String
Dim rc$, fnin$, i%, z$

'd2infile = "Form1": d2insub = "mkkompdn"
mkkompdn = "unbenannt"
rc$ = ""
fnin$ = strrepl(LCase(lfnin$), "ö", "oe")
fnin$ = strrepl(fnin$, "ü", "ue")
fnin$ = strrepl(fnin$, "ä", "ae")
fnin$ = strrepl(fnin$, "ß", "ss")
For i% = 1 To Len(fnin$)
  z$ = Mid$(fnin$, i%, 1)
  If (z$ = " ") Or (z$ = "/") Then
    rc$ = rc$ & "_"
  Else
    If (z$ >= "0" And z$ <= "9") Or (z$ >= "a" And z$ <= "z") Or InStr(".-+", z$) > 0 Then rc$ = rc$ & z$
  End If
Next i%
If rc$ <> "" Then mkkompdn = rc$

End Function

Public Sub rereadsomesysvars()
Dim r As ADODB.Recordset, i As Integer, abc As Integer, s1d$, rrr, cmd$, tr

Dim d2infile As String, d2insub As String
Call clr_usr_setting
d2infile = "Form1": d2insub = "rereadsomesysvars"
'(wider)einlesen bestimmter sysvars bei load und änderungen der benutzereinstellungen zur vermeidung eunes neustarts
s1d$ = form1.getusersetting("Agencyprof")
If exist_by_dir(s1d$ & "\Agencyprof" & trm(App.Major) & ".exe") = 1 Then
  s0d$ = s1d$
End If
On Error Resume Next
lnkcolor = Val(form1.getusersetting("linkcolor", "&H8000000D"))
rrr = Err
On Error GoTo 0
If rrr <> 0 Then lnkcolor = &H8000000D
DoEvents
warnmeondata = True
grantptr = -1
If form1.getusersetting("dataprintwarning", "ja") = "nein" Then warnmeondata = False
datchgmode = form1.getusersetting("datumsformat", "de")
ttmode = getusersetting("tooltipmodus", "2")
tzoffset = CLng(getusersetting("timezoneoffset", "7200"))
usemenu = form1.getusersetting("usemenu", "ja")
Call set_datchgmode(datchgmode)
ichspreche = form1.getusersetting("sprechen", "")
backslashhandler = getusersetting("backslashhandler", "an"): Call chk_backslashhandler(trm(backslashhandler))
srchhigru = form1.getusersetting("searchhigru", "")
dttrenn = form1.getusersetting("datumstrenner", ".")
crlfrepl = getusersetting("crlfreplace", "")
vorlagencache = getusersetting("vorlagencache", "")
If vorlagencache <> "" Then
  On Error Resume Next
  MkDir vorlagencache
  On Error GoTo 0
'  tr = Dir(vorlagencache + "\*.rtf")
'  While tr <> ""
'    On Error Resume Next
'    Kill vorlagencache + "\" + tr
'    On Error GoTo 0
'    tr = Dir
'  Wend
End If
autocheckmail = True
If getusersetting("AutoMailTest", "nein") = "nein" Then autocheckmail = False
If getusersetting("ledszeigen", "nein") = "ja" Then
  cb1.Visible = True
  cb2.Visible = True
  shwled = True
Else
  cb1.Visible = False
  cb2.Visible = False
  shwled = False
End If
If nexist(ichspreche & "\0.wav") Then ichspreche = ""
'If exist_by_dir(s00d$ & "\Agencyprof" & trm(App.Major) & ".exe") = 0 And InStr(LCase(s0d$), "\src") = 0 Then
'  MsgBox "Die Pfadeinstellungen sind falsch. Bitte überprüfen:" & vbCrLf & "s0dir=" & s0d$
'  DoEvents
'  End
'End If
doc0dir$ = "docs" & "." & dbname$: s1d$ = form1.getusersetting("userdocdir"): If s1d$ <> "" Then doc0dir$ = s1d$
m0dir$ = "medien" & "." & dbname$: s1d$ = form1.getusersetting("mediendir"): If s1d$ <> "" Then m0dir$ = s1d$
On Error Resume Next
MkDir m0dir
MkDir m0dir & "\__PROJEKTE__"
MkDir s0d$ & "\Agencyprof" & "\" & m0dir & "\__PROJEKTE__"
On Error GoTo 0
sys_mwst = var2dbl(trm(getusersetting("MwSt", "1900")))
docdup$ = getusersetting("docdup")
For i = 0 To 199: granttab(i) = "": Next i
grantptr = -1
abc = Len("sysvar_" + uId$ + "_grant_") + 1
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT owner,wert FROM sysvars where instr(owner,'sysvar_" + uId$ + "_grant_')=1", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  grantptr = grantptr + 1
  granttab(grantptr) = Mid(r!Owner, abc)
  r.MoveNext
Wend
alertdb = getusersetting("alertdb", "wkruf")
alertdbuser = getusersetting("alertdbuser", "wkrufer")
alertdbpsswd = "wkrAgprOf"
alertdbhost = getusersetting("alertdbhost", "www.agencyprof.com")
mustfield$ = ""
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT * FROM sysvars where instr(owner,'sysvar_" & uId$ & "_mussfeld')>0 or instr(owner,'sysvar_system_mussfeld')>0"
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
While Not r.EOF
  mustfield$ = mustfield$ + " " + trm(r!wert) + " "
  i = i + 1
  r.MoveNext
Wend
End If

End Sub

Public Function darf_ich_sprechen() As Boolean

'd2infile = "Form1": d2insub = "darf_ich_sprechen"
darf_ich_sprechen = False
If ichspreche <> "" Then darf_ich_sprechen = True

End Function

Public Function fixfilename(longname$)
Dim ln$

'd2infile = "Form1": d2insub = "fixfilename"
fixfilename = longname$
If getusersetting("KurzeDateinamen") = "ja" Then
  ln$ = longname$
  fixfilename = GetShortName(ln$)
End If

End Function

Public Function filenamekurz(longname$)
Dim ln$

filenamekurz = longname$
ln$ = longname$
filenamekurz = GetShortName(ln$)

End Function

Public Function mimetype(e$) As String
Dim tg$, rrr, o%, fn$, mex$, mt$, l$, brk As Boolean, ext$, p%

'd2infile = "Form1": d2insub = "mimetype"
mimetype = "application/x-unknown-content-type;"
o% = FreeFile
On Error Resume Next
Open form1.s0dir() & "\mime.typ" For Input As #o%
rrr = Err
On Error GoTo 0
If rrr = 0 Then
  While Not EOF(o%) And brk = False
    Line Input #o%, l$: l$ = trm(strrepl(l$, Chr$(9), " "))
    p% = InStr(l$, " ")
    If p% > 0 Then
      mex$ = trm(Mid$(l$, p% + 1))
      If InStr(mex$, e$) > 0 Then
        brk = True
        mt$ = trm(Left$(l$, p%))
      End If
    End If
  Wend
  Close #o%
End If
If mt$ <> "" Then mimetype = mt$
End Function

Public Function saveasBox(templ$) As String

'd2infile = "Form1": d2insub = "saveasBox"
  saveasBox = ""
  If getusersetting("saveas") <> "comdlg" Then
    With saveas
      .fname = templ$
      .Show vbModal, Me
      If (.SelectionOK) Then
        saveasBox = (.SelectedName)
      End If
    End With
    Unload saveas
  Else
    On Error Resume Next
    With cdlg1
    'Bei "Abbruch" Fehler raisen lassen:
    .CancelError = True
    'Suchpfad einstellen:
    .InitDir = DirName(templ$)
    .FileName = FileName(templ$)
    .DialogTitle = "Speichern unter ..."
    'und endlich den Dialog anzeigen:
    .ShowOpen

    'Auswertung:
    If Err = cdlCancel Then Exit Function
    saveasBox = .FileName

    End With
    On Error GoTo 0

  End If

End Function

Public Sub iCalUpdate(adrid$, aid$)
Dim ical$, icalneu$, icali%, icalo%, l$
Dim alldone As Boolean, wmode%

'd2infile = "Form1": d2insub = "iCalUpdate"
ical$ = form1.s0dir() & "\" & form1.medien() & "\" & adrid$ & ".ics"
If exist(ical$) = 0 Then Exit Sub
Exit Sub
'probably failing:
icali% = FreeFile
Open ical$ For Input As #icali%
icalo% = FreeFile
icalneu$ = ical$ & ".neu"
wmode% = 0
Open icalneu$ For Output As #icalo%
alldone = False
While Not EOF(icali%)
  Line Input #icali%, l$
  If UCase(l$) = "BEGIN:VCALENDAR" Then
    wmode% = 1
  End If
  If UCase(l$) = "END:VCALENDAR" Then
      wmode% = 0
  End If
  Select Case wmode%
    Case 0:  Print #icalo%, l$
    Case 1:
        If Left(UCase(l$), 3) = "UID" Then
          Print #icalo%, l$
          Line Input #icali%, l$
'Debug.Print l$, aid$
          wmode% = 0
          If trm(l$) = ":" & aid$ Then
            alldone = True
            Call writeiCal2openfile(icalo%, aid$)
            Do
              Line Input #icali%, l$
            Loop Until l$ = "END:VCALENDAR"
          Else
            Print #icalo%, l$
          End If
        Else
          Print #icalo%, l$
        End If
    Case Default:
  End Select
Wend
Close #icali%
If Not alldone Then     'neuer eintrag
  Print #icalo%, "BEGIN:VCALENDAR"
  Print #icalo%, "Version"
  Print #icalo%, " :2.0"
  Print #icalo%, "PRODID"
  Print #icalo%, " :-//Agencyprof.de/NONSGML Agencyprof Calendar V0.3//EN"
  Print #icalo%, "BEGIN:VEVENT"
  Print #icalo%, "UID"
  Call writeiCal2openfile(icalo%, aid$)
End If
Close #icalo%
Kill ical$
Name icalneu$ As ical$
End Sub

Public Sub writeiCal2openfile(icalo%, aid$)
Dim rrr
Dim r As ADODB.Recordset, c$, cat$, dtt$, dtg$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "writeiCal2openfile"
'achtung: ausgabe beginnt erst bei uid, header muss bereits geschieben sein

        c$ = "SELECT * FROM auftritt where id='" & aid$ & "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r.EOF Then

        Print #icalo%, " :" & aid$
        Print #icalo%, "SUMMARY"
        Print #icalo%, " :" & strrepl(nouml(r!bezeichnung), ",", "")
        Print #icalo%, "DESCRIPTION"
        Print #icalo%, " :-"
        cat$ = trm(r!ort)
        If cat$ <> "" Then
          Print #icalo%, "LOCATION"
          Print #icalo%, " :" & cat$
        End If
        cat$ = form1.getusersetting("MOZCAT_" & trm(r!auftrittstyp))
        If cat$ = "" Then cat$ = form1.getusersetting("MOZCAT")
        If cat$ = "" Then cat$ = "Miscellaneous"
        Print #icalo%, "CATEGORIES"
        Print #icalo%, " :" & cat$
        cat$ = form1.getusersetting("MOZSTAT_" & form1.get_eventstatusname(trm(r!astatus)))
        If cat$ = "" Then cat$ = form1.getusersetting("MOZSTAT")
        If cat$ = "" Then cat$ = "Tentative"
        Print #icalo%, "STATUS"
        Print #icalo%, " :" & cat$
        Print #icalo%, "CLASS"
        Print #icalo%, " :PUBLIC"
        dtt$ = fixl(strrepl("" & onlynums(trm(r!zeit)), ":", ""), 6)
        dtg$ = strrepl("" & trm(r!datum), "-", "") & "T" & strrepl(dtt$, " ", "0")
        Print #icalo%, "DTSTART"
        Print #icalo%, " :" & dtg$
        Print #icalo%, "DTEND"
        Print #icalo%, " :" & dtg$
        dtg$ = strrepl("" & datum2sql(Date), "-", "") & "T" & strrepl("" & Time, ":", "")
        Print #icalo%, "DTSTAMP"
        Print #icalo%, " :" & dtg$
        Print #icalo%, "END:VEVENT"
        Print #icalo%, "END:VCALENDAR"

        End If


End Sub
Public Sub iCalDelTermin(adrid$, aid$)
Dim ical$, icalneu$, icali%, icalo%, l$
Dim t$, delme As Boolean

'd2infile = "Form1": d2insub = "iCalDelTermin"
ical$ = form1.s0dir() & "\" & form1.medien() & "\" & adrid$ & ".ics"
If exist(ical$) = 0 Then Exit Sub
icali% = FreeFile
Open ical$ For Input As #icali%
icalo% = FreeFile
icalneu$ = ical$ & ".neu"
Open icalneu$ For Output As #icalo%

While Not EOF(icali%)
  Line Input #icali%, l$
  If UCase(l$) = "BEGIN:VCALENDAR" Then
    t$ = "BEGIN:VCALENDAR"
    delme = False
    Do
      Line Input #icali%, l$
      t$ = t$ & vbCrLf & l$
      If Left(UCase(l$), 3) = "UID" Then
        Line Input #icali%, l$
        t$ = t$ & vbCrLf & l$
        If trm(l$) = ":" & aid$ Then delme = True
      End If
    Loop Until UCase(l$) = "END:VCALENDAR"
    If Not delme Then
      Print #icalo%, t$
    Else
      While Not EOF(icali%)
        Line Input #icali%, l$
        Print #icalo%, l$
      Wend
    End If
  Else
    Print #icalo%, l$
  End If
Wend
Close #icali%
Close #icalo%
Kill ical$
Name icalneu$ As ical$
End Sub

Function userformatphone(t) As String
Dim i%, rc$, f$, c$

'd2infile = "Form1": d2insub = "userformatphone"
rc$ = ""
f$ = getusersetting("Nummernzeichen")
If f$ = "" Then
  rc$ = t
Else
  If trm(t) <> "" Then
    For i% = 1 To Len(t)
      c$ = Mid(t, i%, 1)
      If InStr(f$, c$) > 0 Then rc$ = rc$ & c$
    Next i%
    rc$ = strrepl(rc$, "  ", " ")
  End If
End If
userformatphone = rc$
End Function

Public Function instrumentvon(na$) As String

'd2infile = "Form1": d2insub = "instrumentvon"
instrumentvon = xadrdata(na$, "Künstler", "Instrument")

End Function

Public Function xadrdata(na$, typ$, feld$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "xadrdata"
xadrdata = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresse where name='" + na$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Function
If IsNull(rtmp!id) Then Exit Function
cmd$ = "SELECT Felddaten From auftritthigru where auftrittstyp='" & typ$ & "' and auftrittsid='" & rtmp!id & "' and feldname='" & feld$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
xadrdata = trm(r!felddaten)

End Function

Public Function vorlagendir() As String
Dim rc$, tr, rrr

rc$ = s0d$ & "\" & dbname$ & ".rtf"
vorlagendir = getusersetting("vorlagenverzeichnis", rc$)
On Error Resume Next
tr = Dir(rc$ + "\*.rtf")
rrr = Err
On Error GoTo 0
If rrr <> 0 Or tr = "" Then vorlagendir = getusersetting("vorlagencache", rc$)

End Function

Public Function vorlagenverzeichnis() As String
vorlagenverzeichnis = vorlagendir()
End Function

Sub inchonorarlcount()
hontrue = True
honorarlcount% = honorarlcount% + 1
If honorarlcount% > 499 Then
  honorarlcount% = 499
  hontrue = False
  honerr$ = "zu viele Einzelhonorare (>500)"
End If
honvalid = 0
End Sub
Public Sub delalias()
Dim i%, k%

If skip1del Then
  skip1del = False
  Exit Sub
End If
For i% = 0 To 9
  For k% = 0 To 1
   aliasfeld$(i%, k%) = ""
   aliastext$(i%, k%) = ""
  Next k%
Next i%
End Sub

Public Sub addalias(fld$, ali$)
Dim i%

i% = 0
While aliasfeld$(i%, 0) <> "": i% = i% + 1: Wend
aliasfeld$(i%, 0) = fld$
aliasfeld$(i%, 1) = ali$

End Sub

Public Function readalias(f$) As String
Dim o%, l$, i%, rmode As Integer, p%

'd2infile = "Form1": d2insub = "readalias"
readalias = ""
Call delalias
i% = 0
o% = FreeFile
If exist(f$) = 0 Then Exit Function
Open f$ For Input As #o%
Line Input #o%, l$: readalias = l$
rmode = 0
While Not EOF(o%)
  If i% > 98 Then Exit Function
  Line Input #o%, l$
  If trm(l$) = "" Then
    rmode = 1
    i% = 0
  Else
    If rmode = 0 Then
      p% = InStr(l$, "--")
      If p% > 1 Then
        aliasfeld$(i%, 0) = LCase(trm(Left(l$, p% - 1)))
        aliasfeld$(i%, 1) = LCase(trm(Mid(l$, p% + 2)))
        If aliasfeld$(i%, 0) = "fy__var__hauptperson" Then
          listenhauptperson = aliasfeld(i%, 1)
        End If
        i% = i% + 1
      End If
    Else
      p% = InStr(l$, "--")
      If p% > 1 Then
        aliasfeld$(i%, 0) = trm(Left(l$, p% - 1))
        aliasfeld$(i%, 1) = trm(Mid(l$, p% + 2))
        i% = i% + 1
      End If
    End If
  End If
Wend
Close #o%

End Function

Function getaliasfeld(fld$) As String
Dim i%

'd2infile = "Form1": d2insub = "getaliasfeld"
getaliasfeld = fld$
i% = 0
While aliasfeld$(i%, 0) <> "" And i% < 99
  If aliasfeld$(i%, 0) = LCase(fld$) Then
    getaliasfeld = aliasfeld$(i%, 1)
    Exit Function
  End If
  i% = i% + 1
Wend
End Function


Function replacealiastext(lne$) As String
Dim r$, i%, p%, l0$, l2$, al%

'd2infile = "Form1": d2insub = "replacealiastext"
replacealiastext = lne$
r$ = lne$
i% = 0

While aliastext$(i%, 0) <> "" And i% < 99
  p% = InStr(r$, aliastext$(i%, 0))
  While p% > 0
    l0$ = "": l2$ = ""
    al% = Len(aliastext$(i%, 0))
    If p% > 1 Then l0$ = Left(r$, p% - 1)
    If p% + al% < Len(r$) Then l2$ = Mid$(r$, p% + al%)
    r$ = l0$ & aliastext$(i%, 1) & l2$
    p% = InStr(r$, aliastext$(i%, 0))
  Wend
  i% = i% + 1
Wend
replacealiastext = r$

End Function

Public Function auftrittshonorarbyname(aufid$, adrid$) As String
Dim rrr
Dim stmp As ADODB.Recordset, fnam$, j%, honfn$, honi%, atyp$, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittshonorarbyname"
auftrittshonorarbyname = ""

    honi% = 0
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "' and felddaten='" & adrid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      atyp$ = stmp!auftrittstyp
      If LCase(atyp$) = "orchesterauftritt" Then
        If honfn$ = "orchester" Then honfn$ = ""
      Else
        honfn = ""
      End If
      For j% = 1 To sqla.TableDefs("usr_" & utabn(atyp$)).Fields.Count - 1
        fnam$ = LCase$(sqla.TableDefs("usr_" & utabn(atyp$)).Fields(j%).name)
        If LCase(fnam$) = "honorar" & honfn$ Then
          honi% = j%
          cmd$ = "select * from usr_" & utabn(atyp$) & " where id='" & aufid$ & "'"
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If Not stmp.EOF Then
            auftrittshonorarbyname = trm("" & stmp.Fields(honi%).value)
            Exit Function
          End If
        End If
      Next j%
    End If

End Function
Public Function auftrittshonorarfeldbyname(aufid$, adrid$) As String
Dim rrr
Dim stmp As ADODB.Recordset, fnam$, j%, honfn$, honi%, atyp$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittshonorarfeldbyname"
auftrittshonorarfeldbyname = ""

    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "' and felddaten='" & adrid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      auftrittshonorarfeldbyname = "honorar" & honfn$
      Exit Function
    End If

End Function

Public Function auftrittshonorarfeldbyadrid(aufid$, adrid$) As String
Dim rrr
Dim stmp As ADODB.Recordset, fnam$, j%, honfn$, honi%, atyp$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittshonorarfeldbyadrid"
auftrittshonorarfeldbyadrid = ""

    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "select * from auftritthigru where auftrittsid='" & aufid$ & "' and felddaten='" & adrid$ & "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not stmp.EOF Then
      honfn$ = LCase(stmp!feldname)
      If InStr(LCase(honfn$), "künstler") = 1 Then
        honfn$ = "honorar" + onlynums(honfn$)
      End If
      auftrittshonorarfeldbyadrid = honfn$
      Exit Function
    End If

End Function

Function ratefeldaustabelle(tabelle As String, feldname As String) As String
Dim i%, fna$

'd2infile = "Form1": d2insub = "ratefeldaustabelle"
ratefeldaustabelle = feldname
For i% = 1 To sqla.TableDefs(tabelle).Fields.Count - 1
  fna$ = LCase(sqla.TableDefs(tabelle).Fields(i%).name)
  If InStr(fna$, feldname) = 1 Then
    ratefeldaustabelle = fna$
    Exit Function
  End If
Next i%

End Function

Public Function get_kontaktid_by_name(vid$, kid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "get_kontaktid_by_name"
get_kontaktid_by_name = ""

If InStr(kid$, "(") > 0 Then kid$ = trm(cut_d1(kid$, "("))
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id,name FROM kontakt where vid='" + vid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  If trm(rtmp!name) = kid$ Then
    get_kontaktid_by_name = rtmp!id
    Exit Function
  End If
  rtmp.MoveNext
Wend

End Function

Sub l1a(txt$)
Dim tx$, p%, i%, c$

'd2infile = "Form1": d2insub = "l1a"
tx$ = txt$
p% = InStr(tx$, "{")
If p% > 0 Then
  tx$ = trm(Left(tx$, p% - 1))
  For i% = 0 To List2.ListCount - 1
    If Left(List2.List(i%), Len(tx$)) = tx$ Then Exit Sub
  Next i%
  List2.AddItem form1.crlffake(tx$)
  Exit Sub
End If
tx$ = form1.crlffake(tx$)
For i% = 0 To List1.ListCount - 1
  c$ = trm(List1.List(i%))
  If InStr(c$, ":") = 0 Then
    If c$ = tx$ Then Exit Sub
  Else
    If Left(List1.List(i%), Len(tx$) + 1) = tx$ + ":" Then Exit Sub
  End If
Next i%
List1.AddItem tx$
End Sub

Function getplzort(rland, rplz, rort) As String
Dim plzort$, land$, plzmode$, z$, i%

'd2infile = "Form1": d2insub = "getplzort"
getplzort = ""
land$ = trm(rland)
plzmode$ = getusersetting("plzort-" & rland, "")
If plzmode$ = "" Then
  plzmode$ = "P O|L"
  z$ = "insert into sysvars (id,owner,wert) values('" & newid("sysvars", "id", 16) & "','sysvar_system_plzort-" & rland & "','P O|L')"
  Call sqlqry(z$)
End If
'If LCase(land$) = LCase(getusersetting("meinland")) Then land$ = ""
'If trm(rort) <> "" Then plzort$ = trm(rort)
'If trm(rplz) <> "" Then plzort$ = trm(trm(rplz) & " " & plzort$)
'If land$ <> "" Then plzort$ = trm(trm(land$) & " " & plzort$)
For i% = 1 To Len(plzmode$)
  z$ = Mid$(plzmode$, i%, 1)
  Select Case LCase(z$)
    Case "l":
        plzort$ = plzort$ & rland
    Case "p":
        plzort$ = plzort$ & rplz
    Case "o":
        plzort$ = plzort$ & rort
    Case "|":
        plzort$ = plzort$ & vbCrLf
    Case "<":
        plzort$ = trm(plzort$)
    Case Else:
        plzort$ = plzort$ & z$
  End Select
Next i%
getplzort = trm(plzort$)

End Function

Public Function composeemlname(M$) As String
Dim dm$, de$, fn$, mbase$, ebase$, i%, f0$

composeemlname = ""
mbase$ = getusersetting("msgstorebase", ""): If mbase$ = "" Then Exit Function
ebase$ = getusersetting("emlstorebase", ""): If ebase$ = "" Then Exit Function
f0$ = FileName(M$)
dm$ = DirName(M$)
i% = Len(mbase$)
If Len(dm$) < i% Then Exit Function
de$ = "": If Len(dm$) > i% Then de$ = Mid$(dm$, i% + 1)
fn$ = emlfilename(f0$)
f0$ = ebase$ + de$ + "\" + fn$
If fn$ <> "" And Not nexist(f0$) Then composeemlname = f0$
End Function

Public Function dupcheck(fn$) As String
Dim l%, f2$, dhrepl$, s$, e$, ext$

'd2infile = "Form1": d2insub = "dupcheck"
dupcheck = fn$: If Not nexist(fn$) Then Exit Function
ext$ = FileExtension(fn$)
f2$ = getusersetting("msgstorebase", "")
e$ = getusersetting("emlstorebase", "")
If ext$ = "msg" Then
  f2$ = composeemlname(fn$)
  If f2$ <> "" Then
    fn$ = f2$
  End If
End If
If docdup$ <> "" Then
  dupcheck = docdup$ & Mid(fn$, Len(docdup$) + 1)
  If Not nexist(dupcheck) Then Exit Function
End If
If docequiv1$ <> "" Then
  dupcheck = strrepl(fn$, docequiv1$, docequiv2$)
  If Not nexist(dupcheck) Then Exit Function
  dupcheck = strrepl(fn$, docequiv2$, docequiv1$)
  If Not nexist(dupcheck) Then Exit Function
  dupcheck = strrepl(LCase(fn$), LCase(docequiv1$), LCase(docequiv2$))
  If Not nexist(dupcheck) Then Exit Function
  dupcheck = strrepl(LCase(fn$), LCase(docequiv2$), LCase(docequiv1$))
  If Not nexist(dupcheck) Then Exit Function
End If
dhrepl$ = getusersetting("dhreplace", "")
If dhrepl <> "" Then
  s$ = LCase(cut_d1(dhrepl$, "|"))
  e$ = cut_d2bis(dhrepl$, "|")
  dupcheck = strrepl(fn$, s$, e$)
  If Not nexist(dupcheck) Then Exit Function
  dupcheck = strrepl(fn$, s$, e$)
  If Not nexist(dupcheck) Then Exit Function
  dupcheck = strrepl(fn$, UCase(s$), e$)
  If Not nexist(dupcheck) Then Exit Function
End If
dupcheck = fn$
End Function

Public Function kommasettings(l$, mode$)
Dim l1$

'd2infile = "Form1": d2insub = "kommasettings"
l1$ = trm(l$)
kommasettings = l1$
If l1$ <> "" And LCase(getusersetting("anredekomma", "ja")) = "ja" Then
  l1$ = trm(l$)
  If Right$(l1$, 1) <> "," Then
    If mode$ = "an" Or (mode$ = "ab" And LCase(getusersetting("abredekomma", "nein")) <> "nein") Then
      l1$ = l1$ & ","
    End If
  End If
End If
kommasettings = l1$
End Function


Public Function iml(txt$) As String
'd2infile = "Form1": d2insub = "iml"
iml = inmylanguage(txt$)
End Function
Public Function inmylanguage(txt$) As String
Dim ltxt$, i%, o%, l$, fn$, rrr

'd2infile = "Form1": d2insub = "inmylanguage"
inmylanguage = txt$
If currentlanguage = "de" Then Exit Function
ltxt$ = LCase(txt$)
For i% = 0 To ttabptr% - 1
  If LCase(transtab(0, i%)) = ltxt$ Then
    inmylanguage = transtab(1, i%)
    Exit Function
  End If
Next i%
o% = FreeFile
fn$ = s0d$ + "\_untranslated.txt"
If Not nexist(fn$) Then
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    If InStr(LCase(l$), LCase(txt$) + "|") = 1 Then
      Close #o%
      Exit Function
    End If
  Wend
  Close #o%
End If
o% = FreeFile
On Error Resume Next
Open fn$ For Append As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Function
Print #o%, txt$ + "|" + txt$
Close #o%
Debug.Print txt$
End Function

Public Function outmylanguage(txt$) As String
Dim ltxt$, i%

'd2infile = "Form1": d2insub = "outmylanguage"
outmylanguage = txt$
ltxt$ = LCase(txt$)
For i% = 0 To ttabptr% - 1
  If LCase(transtab(1, i%)) = ltxt$ Then
    outmylanguage = transtab(0, i%)
    Exit Function
  End If
Next i%

End Function

Public Sub transtabinit(sprache As String)
Dim o%, p%, l$, fn$

'd2infile = "Form1": d2insub = "transtabinit"
currentlanguage = sprache
o% = FreeFile
fn$ = s0d$ & "\transtab" & "-" & sprache & ".txt"
If Not nexist(fn$) Then
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    p% = InStr(l$, "\|")
    If p% = 0 Then
      p% = InStr(l$, "|")
    Else
      p% = p% + 2
      While p% < Len(l$)
        If Mid(l$, p%, 1) = "|" And Mid(l$, p% - 1, 1) <> "\" Then
          GoTo trtinit_wendbreak
        End If
        p% = p% + 1
      Wend
    End If
trtinit_wendbreak:
    If p% > 0 Then
      transtab(0, ttabptr%) = strrepl(Left(l$, p% - 1), "\|", "|")
      transtab(1, ttabptr%) = strrepl(Mid(l$, p% + 1), "\|", "|")
      If transtab(1, ttabptr%) = "" Then
        transtab(1, ttabptr%) = transtab(0, ttabptr%)
      End If
      ttabptr% = ttabptr% + 1
    End If
  Wend
  Close #o%
  meinesprache = sprache
Else
  meinesprache = "de"
End If
End Sub

Function mmkwe(w) As String

Static es$(0 To 9), ss$(0 To 8), zs$(0 To 9)

es$(0) = ""
es$(1) = "one"
es$(2) = "two"
es$(3) = "three"
es$(4) = "four"
es$(5) = "five"
es$(6) = "six"
es$(7) = "seven"
es$(8) = "eight"
es$(9) = "nine"
ss$(0) = "ten"
ss$(1) = "eleven"
ss$(2) = "twelve"
ss$(3) = "thirteen"
ss$(4) = "fourteen"
ss$(5) = "fifteen"
ss$(6) = "sixteen"
ss$(7) = "seventeen"
ss$(8) = "eighteen"
zs$(0) = ""
zs$(1) = "xxx"
zs$(2) = "twenty"
zs$(3) = "thirty"
zs$(4) = "fourty"
zs$(5) = "fifty"
zs$(6) = "sixty"
zs$(7) = "seventy"
zs$(8) = "eighty"
zs$(9) = "ninety"

Dim e%, z%, h%
Dim hstr$, zstr$

e% = w Mod 10
z% = Int(((w Mod 100) - e) / 10)
h% = Int(w / 100)
If h% > 0 Then
  If z > 0 Or e > 0 Then
    hstr$ = es$(h%) + "hundredand"
  Else
    hstr$ = es$(h%) + "hundred"
  End If
End If
If z% > 0 Then
  If z% = 1 Then
    If e% < 9 Then
      zstr$ = ss$(e%)
    Else
      zstr$ = es$(e%) + "teen"
    End If
  Else
    If e% > 0 Then
      'zstr$ = es$(e%) + "and" + zs$(z%)
      zstr$ = zs$(z%) & es$(e%)
    Else
      zstr$ = zs$(z%)
    End If
  End If
Else
  If e% = 1 Then
    zstr$ = "one"
  Else
    zstr$ = es$(e%)
  End If
End If
mmkwe = hstr$ + zstr$

End Function

Function inwords(l As Long) As String
Dim erg1$
Dim i%, j%

Static p$(0 To 2)
ReDim z%(0 To 2), erg$(0 To 2)


If InStr(l, ",") > 1 Then
  l = Left$(l, InStr(l, ",") - 1)
End If
p$(2) = "million": p$(1) = "thousand": p$(0) = ""
l = Abs(l) ' negatve Zahlen schenken wir uns
erg1$ = Mid$(str$(l), 2)
While Len(erg1$) < 9
  erg1$ = " " + erg1$
Wend
For i% = 7 To 1 Step -3
  z%(j%) = Val(Mid$(erg1$, i%, 3))
  j% = j% + 1
Next i%
For j% = 2 To 0 Step -1
  erg$(j%) = mmkwe(z%(j%))
  If Len(erg$(j%)) > 0 Then erg$(j%) = erg$(j%) + p$(j%)
  If j% > 0 And InStr(erg$(j%), "one") = Len(erg$(j%)) - 3 Then
    erg$(j%) = Left$(erg$(j%), Len(erg$(j%)) - 1)
  End If
Next j%
inwords = erg$(2) + erg$(1) + erg$(0)

End Function

Public Function spambetreff(spamline$) As Boolean
Dim s$, z$, i%, spl$

'd2infile = "Form1": d2insub = "spambetreff"
spl$ = strrepl(spamline$, ".", "")
For i% = 1 To Len(spl$)
  z$ = LCase(Mid$(strrepl(spl$, "\/", "v"), i%, 1))
  If z$ < "a" Or z$ > "z" Then
    Select Case z$:
      Case "ä":
      Case "ö":
      Case "ü":
      Case "ß":
      Case "0": z$ = "o"
      Case "1": z$ = "i"
      Case "5": z$ = "s"
      Case "ì": z$ = "i"
      Case "@": z$ = "a"
      Case Else: z$ = ""
    End Select
  End If
  s$ = s$ & z$
Next i%

spambetreff = False
i% = 0
While i% < 99 And spmlst$(i%) <> "" And Not spambetreff
  If InStr(s$, spmlst$(i%)) > 0 Then spambetreff = True
  i% = i% + 1
Wend

End Function

Public Sub read_spmlst()
Dim i%, fn$, o%

'd2infile = "Form1": d2insub = "read_spmlst"
For i% = 0 To 99: spmlst$(i%) = "": Next i%
fn$ = form1.s0dir() & "\" & "spamwrds.txt"
i% = 0
If Not nexist(fn$) Then
  o% = FreeFile
  Open fn$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, fn$
    fn$ = LCase(trm(fn$))
    If fn$ <> "" Then
      spmlst$(i%) = fn$
      i% = i% + 1
    End If
  Wend
  Close #o%
End If
End Sub

Public Function inccounter() As Long

'd2infile = "Form1": d2insub = "inccounter"
globcount = globcount + 1
inccounter = globcount

End Function
Public Sub sqlex_adresseas(t$, idn$, id$, neuid$)
Dim rrr
Dim c$, r As ADODB.Recordset, fn$, r1 As ADODB.Recordset, r2 As ADODB.Recordset, r3 As ADODB.Recordset
Dim i%, neuaid$, wert$, neuanid$, nkid$, ntid$, o%

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "sqlex_adresseas"
fn$ = form1.mydatadir() & "\" & t$ & "_" & mkfn(id$) & ".sql"
If exist(fn$) = 0 Then
  c$ = "select * from " & t$ & " where " & idn$ & "='" & id$ & "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    c$ = "insert into " & t$ & " (" & idn$ & ") values('" & neuid$ & "');"
    Call form1.sqlqry(c$)
    For i% = 1 To r.Fields.Count - 1
      If trm(r.Fields(i%).value) <> "" Then
        c$ = mkupdcmd(t$, idn$, neuid$, r.Fields(i%).name, r.Fields(i%).Type, r.Fields(i%).value) & ";"
        Call form1.sqlqry(c$)
      End If
    Next i%
    If t$ = "adresse" Then
      c$ = "select * from adresstyp where vid='" & id$ & "' and (kid='-1' or isnull(kid))"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r1.EOF
        neuaid$ = form1.newid("adresstyp", "id", 18)
        c$ = "insert into adresstyp (id,kid) values('" & neuaid$ & "','-1');"
        Call form1.sqlqry(c$)
        For i% = 1 To r1.Fields.Count - 1
          If trm(r1.Fields(i%).value) <> "" _
             And LCase(r1.Fields(i%).name) <> "kid" Then
            wert$ = trm(r1.Fields(i%).value)
            If r1.Fields(i%).name = "vid" Then wert$ = neuid$
            c$ = mkupdcmd("adresstyp", "id", neuaid$, r1.Fields(i%).name, r1.Fields(i%).Type, wert$) & ";"
            Call form1.sqlqry(c$)
          End If
        Next i%
        r1.MoveNext
      Wend
      c$ = "select * from auftritthigru where auftrittsid='" & id$ & "'"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r1.EOF
        neuaid$ = form1.newid("auftritthigru", "id", 36)
        c$ = "insert into auftritthigru (id) values('" & neuaid$ & "');"
        Call form1.sqlqry(c$)
        For i% = 1 To r1.Fields.Count - 1
          If trm(r1.Fields(i%).value) <> "" Then
            wert$ = trm(r1.Fields(i%).value)
            If r1.Fields(i%).name = "auftrittsid" Then wert$ = neuid$
            c$ = mkupdcmd("auftritthigru", "id", neuaid$, r1.Fields(i%).name, r1.Fields(i%).Type, wert$) & ";"
            Call form1.sqlqry(c$)
          End If
        Next i%
        r1.MoveNext
      Wend
      c$ = "select * from anreden where kid='-1." & id$ & "'"
      Set r1 = New ADODB.Recordset
      r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r1.EOF
        neuanid$ = form1.newid("anreden", "id", 16)
        c$ = "insert into anreden (id) values('" & neuanid$ & "');"
        Call form1.sqlqry(c$)
        For i% = 1 To r1.Fields.Count - 1
          If trm(r1.Fields(i%).value) <> "" Then
            wert$ = trm(r1.Fields(i%).value)
            If r1.Fields(i%).name = "kid" Then wert$ = "-1." & neuid$
            c$ = mkupdcmd("anreden", "id", neuanid$, r1.Fields(i%).name, r1.Fields(i%).Type, wert$) & ";"
            Call form1.sqlqry(c$)
          End If
        Next i%
        r1.MoveNext
      Wend
      c$ = "select * from kontakt where vid='" & id$ & "'"
      Set r2 = New ADODB.Recordset
      r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not r2.EOF
        nkid$ = form1.newid("kontakt", "id", 18)
        c$ = "insert into kontakt (id) values('" & nkid$ & "');"
        Call form1.sqlqry(c$)
        For i% = 1 To r2.Fields.Count - 1
          If trm(r2.Fields(i%).value) <> "" Then
            wert$ = trm(r2.Fields(i%).value)
            If r2.Fields(i%).name = "vid" Then wert$ = neuid$
            c$ = mkupdcmd("kontakt", "id", nkid$, r2.Fields(i%).name, r2.Fields(i%).Type, wert$) & ";"
            Call form1.sqlqry(c$)
          End If
        Next i%
        c$ = "select * from anreden where kid='" & r2!id & "'"
        Set r3 = New ADODB.Recordset
        r3.CursorLocation = adUseServer
rrr = form1.adoopen(r3, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not r3.EOF
          neuanid$ = form1.newid("anreden", "id", 16)
          c$ = "insert into anreden (id) values('" & neuanid$ & "');"
          Call form1.sqlqry(c$)
          For i% = 1 To r3.Fields.Count - 1
            If trm(r3.Fields(i%).value) <> "" Then
              wert$ = trm(r3.Fields(i%).value)
              If r3.Fields(i%).name = "kid" Then wert$ = nkid$
              c$ = mkupdcmd("anreden", "id", neuanid$, r3.Fields(i%).name, r3.Fields(i%).Type, wert$) & ";"
              Call form1.sqlqry(c$)
            End If
          Next i%
          r3.MoveNext
        Wend
        c$ = "select * from adresstyp where vid='" & id$ & "' and kid='" & r2!id & "'"
        Set r1 = New ADODB.Recordset
        r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not r1.EOF
          ntid$ = form1.newid("adresstyp", "id", 18)
          c$ = "insert into adresstyp (id) values('" & ntid$ & "');"
          Call form1.sqlqry(c$)
          For i% = 1 To r1.Fields.Count - 1
            If trm(r1.Fields(i%).value) <> "" Then
              wert$ = trm(r1.Fields(i%).value)
              If r1.Fields(i%).name = "vid" Then wert$ = neuid$
              If r1.Fields(i%).name = "kid" Then wert$ = nkid$
              c$ = mkupdcmd("adresstyp", "id", ntid$, r1.Fields(i%).name, r1.Fields(i%).Type, wert$) & ";"
              Call form1.sqlqry(c$)
            End If
          Next i%
          r1.MoveNext
        Wend
        c$ = "select * from auftritthigru where auftrittsid='" & id$ & r2!id & "'"
        Set r1 = New ADODB.Recordset
        r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        While Not r1.EOF
          ntid$ = form1.newid("auftritthigru", "id", 36)
          c$ = "insert into auftritthigru (id) values('" & ntid$ & "');"
          Call form1.sqlqry(c$)
          For i% = 1 To r1.Fields.Count - 1
            If trm(r1.Fields(i%).value) <> "" Then
              wert$ = trm(r1.Fields(i%).value)
              If r1.Fields(i%).name = "auftrittsid" Then wert$ = neuid$ & nkid$
              c$ = mkupdcmd("auftritthigru", "id", ntid$, r1.Fields(i%).name, r1.Fields(i%).Type, wert$) & ";"
              Call form1.sqlqry(c$)
            End If
          Next i%
          r1.MoveNext
        Wend
        r2.MoveNext
      Wend
    End If
    Close #o%
  End If
End If

End Sub

Public Sub delusersetting(se$)
Dim c$

c$ = "delete from sysvars where owner='sysvar_" & uId$ & "_" & se$ & "'"
Call sqlqry(c$)
If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "form1", "delusersetting")
End Sub

Public Sub setusersetting(se$, we$)
Dim i%, ls$, c$

c$ = "delete from sysvars where owner='sysvar_" & uId$ & "_" & se$ & "'"
Call sqlqry(c$)
If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "form1", "setusersetting")
Call sqlqry("insert into sysvars (id,owner,wert) values('" & _
        newid("sysvars", "id", 9) & "','sysvar_" & uId$ & "_" & se$ & "','" & we$ & "')")
If useusrcache = "ja" Then
  ls$ = LCase(se$)
  For i% = 0 To 199
    If LCase(usr_setting(0, i%)) = ls$ Then
      usr_setting(1, i%) = we$
      Exit Sub
    End If
  Next i%
End If

End Sub

Public Sub setusersetting4user(who As String, se$, we$)
Dim c$
'd2infile = "Form1": d2insub = "setusersetting4user"
c$ = "delete from sysvars where owner='sysvar_" & who & "_" & se$ & "'"
Call sqlqry(c$)
If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "form1", "setusersetting4user")
Call sqlqry("insert into sysvars (id,owner,wert) values('" & _
        newid("sysvars", "id", 9) & "','sysvar_" & who & "_" & se$ & "','" & we$ & "')")
End Sub

Public Sub setsystemsetting(se$, we$)
Dim c$
'd2infile = "Form1": d2insub = "setsystemsetting"
c$ = "delete from sysvars where owner='sysvar_system_" & se$ & "'"
Call sqlqry(c$)
If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "form1", "setsystemsetting")
Call sqlqry("insert into sysvars (id,owner,wert) values('" & _
        newid("sysvars", "id", 9) & "','sysvar_system_" & se$ & "','" & we$ & "')")
End Sub

Public Function rdprog(felddaten$) As String
Dim rrr
Dim rprog As ADODB.Recordset, k$, dau$, d$
Dim stmp As ADODB.Recordset, rc$, wid$, sid$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "rdprog"
rc$ = ""
rdprog = rc$
    Set rprog = New ADODB.Recordset
    rprog.CursorLocation = adUseServer
rrr = form1.adoopen(rprog, "SELECT werkid FROM programmliste where programmid='" + felddaten$ + "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rprog.EOF
      wid$ = trm(rprog!werkid): sid$ = ""
      If Left$(wid$, 4) = "SBZ:" Then
        sid$ = Mid$(wid$, 5)
        wid$ = form1.getsatzidbywerkid(sid$)
      End If
      k$ = form1.getkompvornamenamebywerkid(wid$)
      dau$ = form1.getdauerbywerkid(wid$): If sid$ <> "" Then dau$ = ""
      d$ = "(" + form1.getkompdatesbywerkid(wid$) + ")"
      If Left$(LCase$(k$), 7) = "pause p" Or Left$(LCase$(k$), 7) = "oder od" Then
          k$ = ""
          d$ = ""
      End If
      If k$ <> "" Then
        rc$ = rc$ & k$ & " " & d$ & ": "
      End If
      If sid$ = "" Then
        rc$ = rc$ & form1.getwerknamebyid("" & wid$ & "")
      Else
        rc$ = rc$ + form1.getsatznamebyid(sid$) + " " + transe("aus") + " " + form1.getwerknamebyid("" & wid$ & "")
      End If
      If trm(dau$) <> "" Then
        rc$ = rc$ & " (" & dau$
        If InStr(LCase(dau$), "min") = 0 Then rc$ = rc$ + " " + transe("Min.")
        rc$ = rc$ + ") "
      End If
      rc$ = rc$ & vbCrLf
      rprog.MoveNext
    Wend
rdprog = rc$
End Function

Public Function bkmktest1(lne$) As Boolean
Dim l$, p%, rrr

'd2infile = "Form1": d2insub = "bkmktest1"
bkmktest1 = True
l$ = lne$
p% = InStr(l$, "{\*\bkmkstart ")
If p% = 0 Then Exit Function
Do
  l$ = Mid$(l$, p% + 1)
  On Error Resume Next
  p% = InStr(l$, "{\*\bkmkstart ")
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then p% = 0
Loop Until p% = 0
bkmktest1 = False
If InStr(l$, "{\*\bkmkend ") > 0 Then bkmktest1 = True

End Function

Public Function getkompidbywerkid(wid$) As String
Dim rrr
Dim n$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkompidbywerkid"
getkompidbywerkid = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT KomponistenNummer FROM w_loc where id='" + wid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If rtmp.EOF Then Exit Function
If Not IsNull(rtmp!KomponistenNummer) Then getkompidbywerkid = rtmp!KomponistenNummer

End Function

Sub dbgu(wert)

'd2infile = "Form1": d2insub = "dbgu"
If trm(wert) = "" Then Exit Sub
If form1.getusersetting("Textmarkenverfolgen", "nein") <> "ja" Then Exit Sub
dbupgrade.List1.AddItem wert
dbupgrade.List1.ListIndex = dbupgrade.List1.ListCount - 1
DoEvents
End Sub

Sub startlog(wer$, was$)
Dim o%

'd2infile = "Form1": d2insub = "startlog"
If Not starting Then Exit Sub
If was$ = "" Then
  On Error Resume Next
  Kill "startlog_" & wer$ & ".txt"
  On Error GoTo 0
Else
  o% = FreeFile
  Open "strtlog_" & wer$ & ".txt" For Append As #o%
  Print #o%, was$
  Close #o%
End If

End Sub

Public Function isstarting() As Boolean
'd2infile = "Form1": d2insub = "isstarting"
isstarting = starting
End Function

Public Function new_rechnr(docname$, bezeichnung$) As String
Dim r As Long, c$, bez$

'd2infile = "Form1": d2insub = "new_rechnr"
r = Val(trm0(getsystemsetting("RechNr", "0")))
bez$ = "RechNr.: " + trm(r) + ", " + bezeichnung$
If LCase(Right(docname$, 8)) <> "temp.rtf" Then
  Call setsystemsetting("RechNr", trm(r + 1))
  new_rechnr = r
  c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" & _
            form1.newid("dochist", "id", 18) & "','system','-1','" & docname$ & "','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() + "','" & bez$ + "','Rechnungsnummer: " + trm(r) + "','Rechnungsnummer')"
Else
  new_rechnr = "TEST**" + trm(r) + "**TEST"
  c$ = "insert into dochist (id,adresse,kontakt,docname,erstellt,owner,betreff,memoinhalt,doctyp) values('" & _
            form1.newid("dochist", "id", 18) & "','system','-1','" & docname$ & "','" & _
            datum2sql(Date) & " " & Time & "','" & form1.getuserid() + "','" & bez$ + "','TEST-RechNr: " + trm(r) + "','Rechnungsnummer')"
End If
Call form1.sqlqry(c$)
End Function

Public Function projekttyp(id$) As String
Dim rrr
Dim rtmp As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "projekttyp"
projekttyp = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT Hauptperson FROM tplan where id='" + id$ + "'"
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then projekttyp = trm(rtmp!hauptperson)

End Function

Public Function mydir() As String
'd2infile = "Form1": d2insub = "mydir"
mydir = s0d$ & "\" & docs() & "\" & uId$

End Function

Public Function getkontaktidbyname(a$, kn$) As String
Dim rrr
Dim rtmp As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkontaktidbyname"
getkontaktidbyname = "-1"
c$ = "SELECT id FROM kontakt where name='" + kn$ + "' and vid='" + a$ + "';"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then getkontaktidbyname = trm(rtmp!id)

End Function

Public Function mailresend(mailfile$) As Boolean
Dim infn%, l$, mbox$, o%, dtg$, lockfile As String, optfile As String
Dim p%, mboxfile As String, recvln As Boolean, hdr As Boolean

mailresend = False
mbox$ = getusersetting("mailserver")
If mbox$ <> "dir:Outbox" Then Exit Function
MousePointer = 11: DoEvents
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
        lockfile = mboxfile$ + ".lck"
        mboxfile$ = mboxfile$ + ".amf"
        p% = p% + 1
      Loop Until nexist(mboxfile)
      o% = FreeFile: Open lockfile For Output As #o%: Close #o%
      o% = FreeFile
      Open optfile For Output As #o%
      Print #o%, "quittung=0";
      Close #o%

o% = FreeFile
Open mboxfile For Output As #o%
Print #o%, "Return-Path: <" + form1.getuseremail(form1.getuserid()) + ">"
infn% = FreeFile
Open mailfile$ For Input As #infn%
recvln = False: hdr = True
mbox$ = "Received: from Agencyprof-Resender (localhost [127.0.0.1])"
mbox$ = mbox$ + vbCrLf + "    by Agencyprof 1.0 with id " + mboxfile
mbox$ = mbox$ + vbCrLf + "    this is a forwarded, restored mail without original recieve-lines"
While Not EOF(infn%)
  Line Input #infn%, l$
  If hdr Then
    Do
    If InStr(LCase(l$), "received: ") = 1 And hdr = True Then
      Do
        Line Input #infn%, l$
      Loop Until Left$(l$, 1) <> " " And Left$(l$, 1) <> Chr$(9)
      If mbox$ <> "" Then
        Print #o%, mbox$
        mbox$ = ""
      End If
    End If
    If InStr(LCase(l$), "from: ") = 1 And hdr = True Then
      l$ = cut_d2bis(l$, ":")
      Print #o%, "X-Original-From: " + l$
      Do
        Line Input #infn%, l$
      Loop Until Left$(l$, 1) <> " " And Left$(l$, 1) <> Chr$(9)
      Print #o%, "From: Mailsafe<" + form1.getuseremail(form1.getuserid()) + ">"
    End If
    If InStr(LCase(l$), "delivered-to: ") = 1 And hdr = True Then
      Do
        Line Input #infn%, l$
      Loop Until Left$(l$, 1) <> " " And Left$(l$, 1) <> Chr$(9)
    End If
    If InStr(LCase(l$), "return-path: ") = 1 And hdr = True Then
      Do
        Line Input #infn%, l$
      Loop Until Left$(l$, 1) <> " " And Left$(l$, 1) <> Chr$(9)
    End If
    Loop Until (InStr(LCase(l$), "return-path: ") <> 1 And InStr(LCase(l$), "delivered-to: ") <> 1 And InStr(LCase(l$), "received: ") <> 1) Or l$ = ""
    If trm(l$) = "" Or (Left$(l$, 1) = " " And InStr(l$, ":") = 0) Or InStr(trm(LCase(l$)), "content") = 1 Then
      hdr = False
    Else
      Print #o%, l$
    End If
  End If
  If Not hdr Then Print #o%, l$
  DoEvents
Wend
Close #infn%
Close #o%
On Error Resume Next
Kill lockfile
On Error GoTo 0
MousePointer = 0
End Function

Public Function excelfeldtrenner() As String
'd2infile = "Form1": d2insub = "excelfeldtrenner"
excelfeldtrenner = exceldelim$
End Function

Public Sub addmissingfield(tabelle$, feld$)

'd2infile = "Form1": d2insub = "addmissingfield"
missingfields = missingfields + "<" + LCase(tabelle$) + "|" + LCase(feld$) + ">"

End Sub

Public Function isfieldmissing(tabelle$, feld$) As Boolean
'd2infile = "Form1": d2insub = "isfieldmissing"
isfieldmissing = False
If InStr(missingfields, "<" + LCase(tabelle$) + "|" + LCase(feld$) + ">") > 0 Then isfieldmissing = True
End Function

Public Sub fieldcheck(tabe$, fld$)
Dim ty, rrr

'd2infile = "Form1": d2insub = "fieldcheck"
On Error Resume Next
ty = form1.sqla.TableDefs(tabe$).Fields(fld$).Type
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Call addmissingfield(tabe$, fld$)
End If

End Sub

Public Function feldtest(tabe$, fld$) As Boolean
Dim ty, rrr

feldtest = True
On Error Resume Next
ty = form1.sqla.TableDefs(tabe$).Fields(fld$).Type
rrr = Err
On Error GoTo 0
If rrr <> 0 Then feldtest = False

End Function

Public Sub subcall(fn$, sndk$)
Dim hier$, X

'd2infile = "Form1": d2insub = "subcall"
hier$ = CurDir$
Call ChDrive(Left$(fn$, 1) & ":")
Call ChDir(DirName(Mid(fn$, 3)))
X = Shell(fn$, 1)
DoEvents
Call ChDrive(Left$(hier$, 1) & ":")
Call ChDir(hier$)
DoEvents
If sndk$ <> "" Then SendKys sndk$, 1
End Sub

Public Sub SendKys(ByVal txt As String, wmode As Integer)

'd2infile = "Form1": d2insub = "SendKys"
If getusersetting("sendkeysbyevent", "nein") = "ja" Then
  Call SendKeysEx(txt)
Else
  Call SendKeys(txt, wmode)
End If

End Sub

Public Function tagesname(nr%) As String

'd2infile = "Form1": d2insub = "tagesname"
tagesname = dayname(nr%)

End Function
Public Function langtagesname(nr%) As String

'd2infile = "Form1": d2insub = "langtagesname"
langtagesname = longdayname(nr%)

End Function

Sub formload(name$)

'd2infile = "Form1": d2insub = "formload"
If name$ = "dayvw" Then
  Load dayvw
  On Error Resume Next
  Call dayvw.SetFocus
  On Error GoTo 0
  dayvw.Text1.text = CDate(Date)
End If

End Sub
Public Function auftrittsende(aid$, cmode As String) As String
Dim c$, rtmp As ADODB.Recordset, z As String, hh As String, mm As String, rrr, dmm As Integer

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "auftrittsende"
auftrittsende = ""
  c$ = "select felddaten from auftritthigru where auftrittsid='" + aid$ + "' and feldname='zzzsysez' and auftrittstyp='" + auftrittstyp(aid$) + "';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    c$ = trm(rtmp!felddaten)
    If InStr(LCase(cmode$), "noconvert") = 0 Then
      z = Right(c$, 1)
      If z = "m" Or z = "h" Then
        If z = "h" Then
          dmm = Val(c$) * 60
        Else
          dmm = Val(c$)
        End If
        z = form1.auftrittszeit(aid$)
        hh = cut_d1(z, ":"): mm = cut_d2bis(z, ":")
        On Error Resume Next
        mm = Val(hh) * 60 + Val(mm)
        rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then mm = 0
        mm = mm + dmm: dmm = Int(mm / 60): mm = mm - 60 * dmm
        c$ = fixl0(trm(dmm), 2) + ":" + fixl0(trm(mm), 2)
      End If
    End If
  Else
    c$ = ""
  End If
  If c$ = "" Then
    c$ = "select felddaten from auftritthigru where auftrittsid='" + aid$ + "' and feldname='Terminende' and auftrittstyp='" + auftrittstyp(aid$) + "';"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not rtmp.EOF Then
      c$ = datum2sql(trm(rtmp!felddaten))
      c$ = Mid$(c$, 9, 2) & "." & Mid$(c$, 6, 2) & "." & Mid$(c$, 1, 4)
    Else
      c$ = ""
    End If
  End If

auftrittsende = c$
End Function

Public Function terminvizlist(ByVal id$) As String
Dim rrr
Dim c$, rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "terminvizlist"
terminvizlist = ""
  c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisviz';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  c$ = ""
  While Not rtmp.EOF
    c$ = c$ + "," + trm(rtmp!felddaten)
    rtmp.MoveNext
  Wend
  c$ = c$ + ","
terminvizlist = c$
End Function

Public Function getinternalkey() As String
'd2infile = "Form1": d2insub = "getinternalkey"
getinternalkey = internalkey
End Function


Public Function termininvizlist(ByVal id$) As String
Dim rrr
Dim c$, rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "termininvizlist"
termininvizlist = ""
  c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + "' and feldname='zzzsysisinviz';"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  c$ = ""
  While Not rtmp.EOF
    c$ = c$ + "," + trm(rtmp!felddaten)
    rtmp.MoveNext
  Wend
  c$ = c$ + ","
termininvizlist = c$
End Function

Public Function terminisviz4me(id, iam As String) As Boolean
Dim rrr
Dim vzl As String
Dim c$, rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "terminisviz4me"
terminisviz4me = True
vzl = terminvizlist(id)
If trm(vzl) = "," Then Exit Function
If InStr(vzl, "," + iam + ",") > 0 Then Exit Function

c$ = "select groupid from benutzergruppen where userid='" + iam + "';"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  If InStr(vzl, "," + trm(rtmp!groupid) + ",") > 0 Then Exit Function
  rtmp.MoveNext
Wend

terminisviz4me = False

End Function

Public Function terminisinviz4me(id, iam As String) As Boolean
Dim rrr
Dim vzl As String
Dim c$, rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "terminisinviz4me"
terminisinviz4me = True
vzl = termininvizlist(id)
If InStr(vzl, "," + iam + ",") > 0 Then Exit Function
c$ = "select groupid from benutzergruppen where userid='" + iam + "';"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  If InStr(vzl, "," + trm(rtmp!groupid) + ",") > 0 Then Exit Function
  rtmp.MoveNext
Wend
terminisinviz4me = False

End Function

Public Function internaldecrypt(wert As String) As String
Dim n As String

'd2infile = "Form1": d2insub = "internaldecrypt"
n = wert
If Left(n, 8) = "decrypt:" Then
  n = decrypt(Mid$(n, 9), getinternalkey())
End If
internaldecrypt = n

End Function

Function islistfeld(atyp As String, fnam As String) As Boolean
Dim rrr
Dim r As ADODB.Recordset, cmd$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "islistfeld"
islistfeld = False

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
cmd$ = "SELECT zeilen FROM auftrittsfelder where lcase(typ)='" + atyp + "' and (lcase(FeldName)='" + LCase(fnam) + "' or instr(lcase(FeldName),'." + LCase(fnam) + "')>0);"
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  If r!zeilen > 9 Then
    islistfeld = True
    Exit Function
  End If
End If

End Function

Public Function getusersettingfromuser(who As String, fldn$, Optional vifnull As String) As String
Dim r As ADODB.Recordset, vin As String, c1md$, rrr, c$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getusersettingfromuser"
vin = ""
If vifnull <> "" Then vin = vifnull
getusersettingfromuser = vin
c1md$ = "SELECT " + fldn$ + " as rc FROM benutzerdaten where id='" + who + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c1md$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
  If Not r.EOF Then
    If Not IsNull(r!rc) Then getusersettingfromuser = r!rc
    r.Close
    Exit Function
  End If
End If

c$ = "SELECT wert as rc FROM sysvars where owner='sysvar_" & who & "_" & fldn$ & "'"
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Function
If Not r.EOF Then
  If Not IsNull(r!rc) Then getusersettingfromuser = internaldecrypt((r!rc))
  r.Close
  Exit Function
End If
getusersettingfromuser = getsystemsetting(fldn$, vin)
End Function

Public Function adressbeziehung(adrid As String, bez As String, feld As String) As String
Dim c$, r As ADODB.Recordset, r1 As ADODB.Recordset, a2id As String, t1$, t2$, rrr

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "adressbeziehung"
adressbeziehung = "Bez.: " + bez + ", Feld: " + feld + " nicht gefunden."

t1$ = "rel:" + bez
t2$ = "rel:" + strrepl(bez, "_", " ")
c$ = "select wert from adresstyp where typ='" + t1$ + "' or typ='" + t2$ + "' and vid='" + adrid + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Function
If Not r.EOF Then
  a2id = r!wert
  c$ = "select " + feld + " as erg from adresse where id='" + a2id + "'"
  Set r1 = New ADODB.Recordset
  r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then Exit Function
  If Not r1.EOF Then
    adressbeziehung = trm(r1!erg)
  End If
End If

End Function

Public Sub setpoplock(w As Boolean)

'd2infile = "Form1": d2insub = "setpoplock"
If w Then
  cb2.BackColor = RGB(255, 0, 0)
Else
  cb2.BackColor = RGB(0, 255, 0)
End If
cb2.Cls

poplock = w
End Sub

Public Function adoopen(r As ADODB.Recordset, cmd As String, conn As ADODB.Connection, opm1 As Integer, opm2 As Integer, Optional infile$, Optional inproc$) As Variant
Dim rrr, tooktime As Long

'Debug.Print infile$; ":"; inproc$
Call tm_start(9)
If shwled Then
  cb1.BackColor = RGB(255, 0, 0)
  cb1.Cls
End If
Call dbg2f(cmd, infile$, inproc$)
On Error Resume Next
'r.Open cmd, conn, opm1, opm2
r.Open cmd, conn, adOpenDynamic, adLockReadOnly
rrr = Err
On Error GoTo 0
adoopen = rrr
If shwled Then
  cb1.BackColor = RGB(0, 255, 0)
  cb1.Cls
End If
tooktime = tm_stop(9)
If Label10.Visible Then
  adostats_tsum = adostats_tsum + tooktime
  adostats_samples = adostats_samples + 1
'Debug.Print cmd + vbCrLf + trm(tooktime) + " ms, avg=" + trm(Int(10 * adostats_tsum / adostats_samples) / 10)
  Label10.Caption = trm(Int(10 * adostats_tsum / adostats_samples) / 10) + " ms"
End If
End Function

Private Sub wrk_open_Click()
Call Command3_Click
End Sub

Private Sub wrk_prgopn_Click()
Call Command10_Click
End Sub

Sub signaturinclude()
Dim r As ADODB.Recordset, rrr, sig$
Dim o%, l$, anz As Integer, c$

anz = 0
c$ = "select count(*) as anz from sysvars where instr(owner,'" + uId$ + "_vorsignatur" + "')>0 order by owner;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
  If Not r.EOF Then
    anz = r!anz
  End If
End If
If anz > 0 Then
  anz = Int(Rnd * anz)
  c$ = "select wert from sysvars where instr(owner,'" + uId$ + "_vorsignatur" + "')>0 order by owner;"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
  c$ = ""
  If rrr = 0 Then
    If Not r.EOF Then
      Do
        c$ = r!wert
        r.MoveNext
        anz = anz - 1
      Loop Until r.EOF Or anz < 0
    End If
  End If
  If c$ <> "" Then smtp.txtMessageText.text = smtp.txtMessageText.text + Chr$(13) + Chr$(10) + c$
End If
sig$ = ""
o% = FreeFile
sig$ = getusersetting("mailsignatur", s0d$ & "\" + docs() + "\" & form1.getuserid() & "\signatur.txt")
On Error Resume Next
Open sig$ For Input As #o%
rrr = Err
On Error GoTo 0
sig$ = ""
If rrr = 0 Then
  While Not EOF(o%)
    Line Input #o%, l$
    sig$ = sig$ + Chr$(13) + Chr$(10) + l$
  Wend
  Close #o%
  If sig$ <> "" And InStr(smtp.txtMessageText.text, sig$) = 0 Then smtp.txtMessageText.text = smtp.txtMessageText.text + Chr$(13) + Chr$(10) + sig$
End If
anz = 0
c$ = "select count(*) as anz from sysvars where instr(owner,'" + uId$ + "_nachsignatur" + "')>0 order by owner;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
  If Not r.EOF Then
    anz = r!anz
  End If
End If
If anz > 0 Then
  anz = Int(Rnd * anz)
  c$ = "select wert from sysvars where instr(owner,'" + uId$ + "_nachsignatur" + "')>0 order by owner;"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
  c$ = ""
  If rrr = 0 Then
    If Not r.EOF Then
      Do
        c$ = r!wert
        r.MoveNext
        anz = anz - 1
      Loop Until r.EOF Or anz < 0
    End If
  End If
  If c$ <> "" Then smtp.txtMessageText.text = smtp.txtMessageText.text + Chr$(13) + Chr$(10) + c$
End If

End Sub

Public Sub xmysettings()
Dim r As ADODB.Recordset, rrr
Dim o%, l$, c$, ofn$, p%, X, u$, enc$

Call killxmysettings
ofn$ = form1.getmyhomepath() + "\settings.agp"
o% = FreeFile
Open ofn$ For Output As #o%
Print #o%, "userid=" + uId$
Print #o%, "userdocdir=" + docs()
enc$ = "hihallohuhu4716"
c$ = getusersetting("Mailserver", "")
If c$ <> "" Then Print #o%, "Mailserver=" + trm(c$)
c$ = "SELECT * from sysvars WHERE InStr(owner,'_" + uId$ + "_')>0;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
  While Not r.EOF
    c$ = trm(r!Owner)
    p% = InStr(c$, "_" + uId$ + "_")
    If p% > 0 Then
      c$ = Mid(c$, p% + Len(uId$) + 2)
      Print #o%, c$ + "=" + trm(r!wert)
'Debug.Print c$ + "=" + trm(r!wert)
    End If
    r.MoveNext
  Wend
End If
c$ = "SELECT * from sysvars WHERE InStr(owner,'_system_')>0;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
If rrr = 0 Then
  While Not r.EOF
    c$ = trm(r!Owner)
    p% = InStr(LCase(c$), "_system_")
    If p% > 0 Then
      c$ = Mid(c$, p% + 8)
      Print #o%, c$ + "=" + trm(r!wert)
'Debug.Print c$ + "=" + trm(r!wert)
    End If
    r.MoveNext
  Wend
End If
If LCase(upop$) = "dir:inbox" Then
  c$ = "SELECT * FROM poplist where id='" & uId$ & "_DEFAULT'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
  If Not r.EOF Then
    Print #o%, "popuser=" + trm(r!user)
    Print #o%, "popserver=" + trm(r!server)
    Print #o%, "popport=" + trm(r!Port)
    Print #o%, "poppsswd=" + trm(r!psswd)
  End If
Else
  Print #o%, "popuser=" + upopid$
  Print #o%, "popserver=" + upop$
  Print #o%, "popport=" + trm(upopport%)
  Print #o%, "poppsswd=" + encrypt(upoppsswd$, enc$)
End If
Close #o%

ofn$ = form1.getmyhomepath() + "\poplist.agp"
o% = FreeFile
Open ofn$ For Output As #o%
u$ = uId$
c$ = "SELECT * FROM poplist where instr(id,'" & u$ & "_')=1 and id<>'" & u$ & "_PDFServer'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
While Not r.EOF
  If r!id <> "PDFServer" Then
    Print #o%, r!id; "|"; r!server; "|"; r!user; "|"; r!psswd; "|"; trm(r!Port)
  End If
  r.MoveNext
Wend
Close #o%
'x = Shell("notepad.exe " & ofn$, 1)
End Sub

Public Sub killxmysettings()
Dim ofn$

  ofn$ = form1.getmyhomepath() + "\settings.agp"
  On Error Resume Next
  Kill ofn$

End Sub

Public Function isTitleRunning(ttext As String) As Boolean
Dim Length As Long
Dim sTitel As String
Dim CurHwnd As Long

isTitleRunning = False
CurHwnd = GetWindow(hWnd, GW_HWNDFIRST)
Do While CurHwnd <> 0
  ' Fenstertitel ermitteln
  sTitel = Space$(255)
  Length = GetWindowText(CurHwnd, sTitel, Len(sTitel))
  sTitel = Left$(sTitel, Length)

  ' Fenstertitel prüfen
  If InStr(sTitel, ttext) > 0 Then
    isTitleRunning = True
    Exit Do
  End If

  ' Handle des nächsten Fensters
  ' 0, wenn kein weiteres Fenster vorhanden
  CurHwnd = GetWindow(CurHwnd, GW_HWNDNEXT)
Loop
End Function

Private Sub checkmail()
Dim ol$, sqa As Database, ttest As TableDef, rrr, i%, tn$
Dim intMessageNum As Integer, dn As String

If umchk$ <> "yes" Then Exit Sub

Call rlist3
If LCase(upop$) <> "dir:inbox" Then Exit Sub
Command8.Picture = mlstat(2).Picture
dn = form1.s0dir() + "\" + form1.docs() + "\" + uId$ + "\mail\inbox"
intMessageNum = InboxMessageCount(dn)
If intMessageNum = 0 Then
  Command8.Picture = mlstat(0).Picture
Else
  Command8.Picture = mlstat(1).Picture
End If
'If Not nexist(s0d + "\AgencyprofPOPClient.exe") Then
'  If Not isTitleRunning("Agencyprof - POPClient") Then
'    Command8.Picture = mlstat(2).Picture
'  End If
'End If
  
DoEvents
End Sub

Public Function projektfarbe(typ$) As Long

projektfarbe = RGB(255, 255, 255)
Select Case LCase(typ$)
    Case "künstler": projektfarbe = RGB(0, 255, 0)
    Case "orchester": projektfarbe = RGB(196, 196, 255)
    Case "kammermusik": projektfarbe = RGB(0, 255, 255)
    Case "crossover": projektfarbe = RGB(0, 255, 255)
    Case Else: projektfarbe = RGB(255, 0, 255)
End Select

End Function

Public Function formpos(f As Form)

If getusersetting("limitforms2screen", "ja") <> "ja" Then Exit Function
If f.Left + f.Width > Screen.Width Then f.Left = Screen.Width - f.Width
If f.Top + f.Height > Screen.Height Then f.Top = Screen.Height - f.Height - 200

End Function

Public Function neuevertragsnummer() As String
Dim vnr As String

vnr = newvnr("opt_vnr", "id")
neuevertragsnummer = trm(vnr)
End Function

Public Function saison(dtg As String) As String
Dim yy As Integer, mm As Integer, dd As Integer, l$, strt As Integer
Dim s1$, s2$, trenn$

l$ = dtg
trenn$ = "."
If InStr(dtg, trenn$) = 0 Then trenn$ = "/"
If InStr(dtg, trenn$) = 0 Then trenn$ = "-"
dd = Val(cut_d1(l$, trenn$)): l$ = cut_d2bis(l$, trenn$)
mm = Val(cut_d1(l$, trenn$)): l$ = cut_d2bis(l$, trenn$)
yy = Val(cut_d1(l$, trenn$)): l$ = cut_d2bis(l$, trenn$)
strt = trm0(getusersetting("saisonstart", "0"))
If strt < 2 Then
  saison = trm(yy)
  Exit Function
End If
If mm < strt Then
  s1$ = trm(yy - 1)
  s2$ = trm((yy) Mod 100)
Else
  s1$ = trm(yy)
  s2$ = trm((yy + 1) Mod 100)
End If
If Len(s1$) < 2 Then s1$ = "0" + s1$
If Len(s2$) < 2 Then s2$ = "0" + s2$
saison = s1$ + "/" + s2$
End Function

Private Sub adpprint(adm As Integer, kanal As Integer, wert As String)
Print #kanal, wert;
Debug.Print "printing: " + trm(wert)
If adm < 0 Or adm > 9 Then Exit Sub
adruckmerkwert(adm%) = wert
End Sub
Public Function newvnr(t$, key$) As String
Dim stmp As ADODB.Recordset, cmd$, rrr, nnr As String
Dim cnt As Long, tvnr$, sa$

Dim d2infile As String, d2insub As String

cnt = 1: sa$ = saison(trm(Date))
d2infile = "Form1": d2insub = "newvnr"
Do
  tvnr$ = sa$ + " " + trm(cnt)
  cmd$ = "SELECT id FROM " + t$ + " where id='" + tvnr$ + "'"
  Set stmp = New ADODB.Recordset
  stmp.CursorLocation = adUseServer
  rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr <> 0 Then
    newvnr = transe("Fehler")
    Exit Function
  End If
  cnt = cnt + 1
Loop Until stmp.EOF
nnr = tvnr$
cnt = 0
cmd$ = "SELECT max(sortnr) as mx FROM " + t$
Set stmp = New ADODB.Recordset
stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
cnt = trm0(stmp!mx)
cnt = cnt + 1
cmd$ = "insert into " + t$ + " (" + key$ + ",sortnr) values('" + nnr + "'," + trm(cnt) + ")"
Call sqlqry(cmd$)
newvnr = nnr

End Function

Public Function plzoofadr(adrid$) As String
Dim c$
Dim r As ADODB.Recordset, rrr

c$ = "select plz,ort from adresse where id='" + adrid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly)
c$ = ""
If rrr = 0 Then
  If Not r.EOF Then c$ = trm(trm(r!plz) + " " + trm(r!ort))
End If
plzoofadr = c$
End Function

Sub mailboxinit()
Dim myinbox As String, tr

myinbox = form1.s0dir() + "\" + form1.docs() + "\" + uId$ + "\mail\inbox"
tr = Dir(myinbox + "\*.lck")
While tr <> ""
  On Error Resume Next
  Kill myinbox + "\" + tr
  On Error GoTo 0
  tr = Dir
Wend

End Sub

Public Sub neuart(vid, kid, art, wert)
Dim idt As Recordset, c$, nid As String

c$ = "select * from adresstypen where id='" + art + "'"
Set idt = sqla.OpenRecordset(c$, dbOpenDynaset, dbOpenDynaset)
If idt.EOF Then
  c$ = "select * from adressgruppenindex where id='" + art + "'"
  Set idt = sqla.OpenRecordset(c$, dbOpenDynaset, dbOpenDynaset)
  If idt.EOF Then
    c$ = "insert into adressgruppenindex (id) values('" + art + "')"
    Call form1.sqlqry(c$)
  End If
  nid = form1.newid("adressgruppen", "id", 20)
  c$ = "insert into adressgruppen (id,adressid,grpid,kid) values('" + nid + "','" + vid + "','" + art + "','" + kid + "')"
  Call form1.sqlqry(c$)
  Exit Sub
End If
c$ = "select * from adresstyp where typ='" + art + "' and vid='" + vid + "' and kid='" + kid + "'"
Set idt = sqla.OpenRecordset(c$, dbOpenDynaset, dbOpenDynaset)
If Not idt.EOF Then
  Exit Sub
End If
c$ = "insert into adresstyp (id,vid,typ,kid,wert) values('" & _
  form1.newid("adresstyp", "id", 18) & "','" & _
  vid & "','" & _
  art & "','" & _
  kid & "','" & _
  wert & "')"
Call form1.sqlqry(c$)
End Sub

Public Function higruinsert(auftrittsid$, auftrittstyp$, feldname$, felddaten$) As Integer
Dim c$

higruinsert = False
If trm(felddaten$) = "" Then Exit Function
If trm(felddaten$) = "(null)" Then felddaten$ = ""
c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,Feldname,felddaten) values('" & _
         form1.newid("auftritthigru", "id", 9) & "','" & _
         auftrittsid$ & "','" & auftrittstyp$ & "','" & feldname$ & "','" & _
         felddaten$ & "')"
Call form1.sqlqry(c$)
higruinsert = True

End Function

Private Function needrecomptxt(rtf$, txt$) As Boolean
Dim dtg1, dtg2

needrecomptxt = False

If nexist(rtf$) Then Exit Function
dtg1 = FileDateTime(rtf$)
If nexist(txt$) Then
  needrecomptxt = True
Else
  dtg2 = FileDateTime(txt$)
  If dtg1 > dtg2 Then needrecomptxt = True
End If
End Function

Public Function crlffake(txt As String) As String
If crlfrepl <> "" Then
  crlffake = strrepl(txt, vbCrLf, crlfrepl)
Else
  crlffake = txt
End If
End Function

Public Function getkontaktnamebyid(kid$) As String
Dim rrr
Dim n$, vid$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "getkontaktabteilungbyid"
getkontaktnamebyid = ""

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT name FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then Exit Function
vid$ = trmx1(rtmp!name)
rtmp.Close
If vid$ = "" Then Exit Function
getkontaktnamebyid = vid$

End Function

Public Function higruget(id$, kid$, typ$, f$) As String
Dim c$, k$, r As ADODB.Recordset

higruget = ""
k$ = kid$: If k$ = "-1" Then k$ = ""
c$ = "select felddaten from auftritthigru where auftrittsid='" + id$ + k$ + "' and auftrittstyp='" + typ$ + "' and feldname='" + f$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.adoc, adOpenDynamic, adLockReadOnly
If Not r.EOF Then higruget = trm(r!felddaten)

End Function

Public Function getstimmtonbywerkid(id$) As String
Dim rrr
Dim n$, p%
Dim rtmp As ADODB.Recordset

getstimmtonbywerkid = ""
If isfieldmissing("opt_stimmton", "id") Then Exit Function
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT stimmton FROM opt_stimmton where id='" + id$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")

If rtmp.EOF Then Exit Function
If IsNull(rtmp!stimmton) Then Exit Function
getstimmtonbywerkid = rtmp!stimmton

End Function

Sub clear_honorarliste()
Dim i%, j%

  honorarlcount% = 0
  honvalid = 0
  For i% = 0 To 6: For j% = 0 To 499: Honorarliste$(i%, j%) = "": Next j%: Next i%
End Sub

Private Sub clr_usr_setting()
Dim i%, j%
For i% = 0 To 199
  usr_set_hits(i%) = 0
  For j% = 0 To 1
    usr_setting(j, i%) = ""
  Next j%
Next i%
End Sub

Public Function getanabrede(adrid$, kid$) As String
Dim rtmp As ADODB.Recordset, anred$, abred$, rrr

anred$ = "": abred$ = ""
If kid$ <> "" And kid$ <> "-1" Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='" + kid$ + "' and user='" + anredeuser$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")
  If Not rtmp.EOF Then
    If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
    If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
  Else
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='" + kid$ + "' and user='system'", adoc, adOpenDynamic, adLockReadOnly, "", "")
    If Not rtmp.EOF Then
      If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
      If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
    End If
  End If
Else
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='-1." + adrid$ + "' and user='" + anredeuser$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")
  If Not rtmp.EOF Then
    If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
    If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
  Else
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM anreden where kid='-1." + adrid$ + "' and user='system'", adoc, adOpenDynamic, adLockReadOnly, "", "")
    If Not rtmp.EOF Then
      If Not IsNull(rtmp!an) Then anred$ = kommasettings(rtmp!an, "an")
      If Not IsNull(rtmp!Ab) Then abred$ = kommasettings(rtmp!Ab, "ab")
    End If
  End If
End If
If anred$ = "" Then anred$ = getusersetting("StandardAnrede", "")
If abred$ = "" Then abred$ = getusersetting("StandardAbrede", "")
getanabrede = anred$ + "|" + abred$
End Function

Public Function pfstrasse(plz As String, plzpostfach As String, postfach As String, strasse As String) As String
Dim pferg$, stra$, pfadr As Boolean

pferg$ = getusersetting("postfachergänzen", "")
pfstrasse = strasse
stra$ = strasse
  If shwAdrDetail.Check3.value = 1 And trm(postfach) <> "" And trm(plzpostfach) <> "" Then
    stra$ = postfach
    pfadr = True
    If pferg$ <> "" Then
      If InStr(LCase(stra$), pferg$) = 0 Then
        stra$ = pferg$ & " " & stra$
      End If
    End If
  End If
  pfstrasse = stra$
End Function

Public Function plzortpostfach(land As String, plz As String, plzpostfach As String, ort As String) As String

plzortpostfach = getplzort(land, plz, ort)
If shwAdrDetail.Check3.value = 1 And trm(plzpostfach) <> "" Then plzortpostfach = getplzort(land, plzpostfach, ort)
End Function

Public Sub switchdb(u$, dbn$, dbpara$, adopara$)
Dim i%

dbswitch = False
Call unloadall
Unload shwAdrDetail

If uId$ <> u$ Or dbname$ <> dbn Then
  dbswitch = True
  uId$ = u$
  Call clr_usr_setting
End If
If dbname$ <> dbn Then
  dbname$ = dbn$
  Set sqla = wrkJet.OpenDatabase(dbname$, dbDriverCompleteRequired, False, dbpara$)
  Set pub_sqla = wrkJet.OpenDatabase(dbname$, dbDriverCompleteRequired, False, dbpara$)
  Set adoc = New ADODB.Connection
  adoc.ConnectionString = adopara$
  adoc.Open
End If
If dbswitch Then
  Call form1.rereadsomesysvars
  autocheckmail = False
  fallbackdir$ = "": fallbackserver$ = "": fallbackserverpath$ = ""
End If
Load shwAdrDetail
End Sub

Public Function check_tst(aid$)
Dim rrr, ckow(0 To 99) As String, ckl(0 To 99) As String, cko(0 To 99) As Double
Dim ckid(0 To 99) As String, ckn%, sid$
Dim r As ADODB.Recordset, s As ADODB.Recordset, rdtg$
Dim c$, cmd$, i%, dtg$, dto As Variant, rc$, c1$
Dim d2infile As String, d2insub As String, confmode$
'd2infile = "Form1": d2insub = "auftrittstyp"

check_tst = False
confmode$ = currentconfmode$
If isfieldmissing("opt_checks", "id") Then Exit Function
cmd$ = "SELECT auftrittstyp,Datum FROM auftritt where id='" + aid$ + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
rdtg$ = datfromsql(r!datum)
c$ = "select * from opt_checklists where auftrittstyp='" + trm(r!auftrittstyp) + "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If s.EOF Then Exit Function
ckn% = 0
While Not s.EOF
  sid$ = trm(s!id)
  c$ = "select * from opt_checks where auftrittsid='" + aid$ + "' and checkid='" + sid$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If r.EOF Then
Debug.Print "add check " + sid$
    dto = CDate(rdtg$) + s!offset
    dtg$ = datum2sql(trm(dto))
    c$ = "insert into opt_checks (id,auftrittsid,checkid,dtg,ownr,confirmed) values("
    c$ = c$ + "'" + newid("opt_checks", "id", 19) + "',"
    c$ = c$ + "'" + aid$ + "',"
    c$ = c$ + "'" + sid$ + "',"
    c$ = c$ + "'" + dtg$ + "',"
    rc$ = trm(s!ownr)
    If Left$(rc$, 1) = "{" Then
      c1$ = "select FeldDaten as wert from auftritthigru where auftrittsid='" + aid$ + "' and FeldName='" + Mid$(rc$, 2) + "'"
      c1$ = form1.get1erg(c1$)
      If c1$ <> "" Then c1$ = form1.APUsernameByAddressID(c1$)
      If c1$ <> "" Then rc$ = c1$
    End If
    c$ = c$ + "'|" + rc$ + "|','" + confmode$ + "')"
    Call sqlqry(c$)
  End If
  s.MoveNext
Wend
If confmode$ = "ok, deleted" Then Exit Function
If Not check_tst Then
  c$ = "select * from opt_checks where auftrittsid='" + aid$ + "' and dtg<='" + datum2sql(Date) + "' and confirmed not like 'ok, confirme%'"
  Set s = New ADODB.Recordset
  s.CursorLocation = adUseServer
  rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not s.EOF Then
    check_tst = True
  End If
End If
End Function

Public Function check_pointbyid(chkid$) As String
Dim rrr
Dim s As ADODB.Recordset
Dim c$

check_pointbyid = ""
c$ = "select * from opt_checklists where id='" + chkid$ + "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If Not s.EOF Then check_pointbyid = trm(s!checkpoint)

End Function

Public Function check_optpointbyid(chkid$) As String
Dim rrr
Dim s As ADODB.Recordset
Dim c$

check_optpointbyid = ""
c$ = "select * from opt_checks where id='" + chkid$ + "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If Not s.EOF Then check_optpointbyid = trm(s!checkpoint)

End Function

Public Function reptest(vid$, wid$) As Integer
Dim rrr
Dim s As ADODB.Recordset
Dim c$

reptest = -1
c$ = "select neverever from opt_repertoire where wid='" + wid$ + "' and vid='" + vid$ + "'"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If Not s.EOF Then
  reptest = trm(s!neverever)
End If

End Function

Public Function checkdates(aid$, von As String, bis As String) As String
Dim rrr
Dim s As ADODB.Recordset
Dim r As ADODB.Recordset
Dim c$, fld$
Dim dtg$, prv$

'Debug.Print aid$
checkdates = ""
Exit Function
'Termintest:
c$ = "SELECT auftritthigru.FeldName, auftritthigru.FeldDaten,auftritt.Datum "
c$ = c$ + "FROM (auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid) INNER JOIN auftrittsfelder ON auftritthigru.auftrittstyp = auftrittsfelder.typ "
c$ = c$ + "WHERE auftritt.id='" + aid$ + "' AND InStr(auftrittsfelder.FeldName,concat(auftritthigru.FeldName,'.'))=11 "
c$ = c$ + "ORDER BY auftritthigru.FeldDaten"
Set s = New ADODB.Recordset
s.CursorLocation = adUseServer

rrr = form1.adoopen(s, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
prv$ = ""
If rrr <> 0 Then Exit Function
While Not s.EOF
  If prv$ <> trm(s!felddaten) Then
    prv$ = trm(s!felddaten)
    Debug.Print s!datum; " "; s!feldname; " "; s!felddaten
  
    c$ = "SELECT auftritt.id, auftritt.Datum, auftritt.Zeit, auftritthigru.FeldDaten, auftritt.Bezeichnung, auftritt.Auftrittstyp, auftritthigru_1.FeldName as fn2, auftritthigru_1.FeldDaten "
    c$ = c$ + "FROM ((auftritt INNER JOIN auftritthigru ON auftritt.id = auftritthigru.auftrittsid) INNER JOIN auftrittsfelder ON auftritthigru.auftrittstyp = auftrittsfelder.typ) INNER JOIN auftritthigru AS auftritthigru_1 ON auftritt.id = auftritthigru_1.auftrittsid "
    c$ = c$ + "WHERE auftritt.id<>'" + aid$ + "' AND auftritt.Datum='" + s!datum + "' AND auftritthigru_1.FeldDaten='" + s!felddaten + "' AND auftritthigru.FeldName='zzzsysez' AND auftrittsfelder.FeldName Like 'adrselect%'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
    If rrr = 0 Then
      While Not r.EOF
        If InStr(trm(r!fn1), r!fn2) = 11 Then
        End If
        r.MoveNext
      Wend
    End If
  End If
  s.MoveNext
Wend

End Function

Sub showprios()
Dim c$, cmd$, rtmp As ADODB.Recordset, rrr, p As Integer

If isfieldmissing("opt_prios", "id") Then Exit Sub

cmd$ = "SELECT * FROM opt_prios where userid='" + form1.getuserid() + "' order by prio,evnt"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not rtmp.EOF
    If cut_d1(trm(rtmp!evnt), ":") = "A" Then
      List1.AddItem cut_d2bis(trm(rtmp!evnt), ":")
    Else
      List3.AddItem "Projekt: " + cut_d2bis(trm(rtmp!evnt), ":")
    End If
    rtmp.MoveNext
  Wend
End If
c$ = "select owner from sysvars where owner like 'sysvar_" + form1.getuserid() + "_zzzadr_sticky_%' order by owner"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not rtmp.EOF
    c$ = trm(rtmp!Owner): p = InStr(c$, "sticky_")
    If p > 0 Then
      On Error Resume Next
      List1.AddItem Mid$(c$, p + 7)
      On Error GoTo 0
      Call form1.c1add(" " + Mid$(c$, p + 7))
    End If
    rtmp.MoveNext
  Wend
End If
End Sub

Public Sub chkallnums(vid$, kid$, typ$, num$)
Dim cmd$, rtmp As ADODB.Recordset, rrr

If trm(vid$) = "" Or trm(typ$) = "" Or trm(num$) = "" Then Exit Sub
If isfieldmissing("opt_allenummern", "vid") Then Exit Sub

cmd$ = "SELECT id FROM opt_allenummern where "
If kid$ <> "-1" Then
  cmd$ = cmd$ + "kid='" + kid$ + "' "
Else
  cmd$ = cmd$ + "vid='" + vid$ + "' and kid='-1' "
End If
cmd$ = cmd$ + "and numtyp='" + typ$ + "' and num='" + num$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If rtmp.EOF Then
    cmd$ = "insert into opt_allenummern (id,vid,kid,numtyp,num) values("
    cmd$ = cmd$ + "'" + form1.newid("opt_allenummern", "id", 60) + "',"
    cmd$ = cmd$ + "'" + vid$ + "',"
    cmd$ = cmd$ + "'" + kid$ + "',"
    cmd$ = cmd$ + "'" + typ$ + "',"
    cmd$ = cmd$ + "'" + num$ + "')"
    Call sqlqry(cmd$)
  End If
End If
End Sub

Function emlfilename(fin$) As String
Dim i%, r$, z$, l$, bsfn$, f$

f$ = FileName(fin$)
emlfilename = f
r$ = ""
l$ = f

If LCase(Right$(l$, 4)) = ".msg" Then
  bsfn$ = Left(l$, Len(l$) - 4)
Else
  emlfilename = ""
  Exit Function
End If
For i% = 1 To Len(bsfn$)
  z$ = Mid$(f, i%, 1)
  If z$ = "-" Or (LCase(z$) >= "a" And LCase(z$) <= "z") Or (LCase(z$) >= "0" And LCase(z$) <= "9") Then
      r$ = r$ + z$
  Else
      r$ = r$ + "_"
  End If
Next i%
emlfilename = r$ + ".eml"

End Function

Function HonorarVA(aid$) As String
Dim rrr
Dim r As ADODB.Recordset, cmd$, hon$, bps$, dau$, waehr$, h1on$, wert1 As Double, wert2 As Double
Dim typ$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "HonorarVonAuftritt"
HonorarVA = "0.00 " + transe("€")
cmd$ = "SELECT * FROM auftritthigru where auftrittsid='" + aid$ + "' and (feldname='HonorarVA')"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If r.EOF Then Exit Function
On Error Resume Next
HonorarVA = CDbl(strrepl(trm(r!felddaten), ".", ""))
rrr = Err
On Error GoTo 0
r.Close
If rrr <> 0 Then HonorarVA = "0.00 " + transe("€")
End Function

Public Function allmailadresses(lic$, kid$) As String
Dim rtmp As ADODB.Recordset, anred$, abred$, rrr, rc$

rc$ = ""
If kid$ <> "" And kid$ <> "-1" Then
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT email FROM kontakt where id='" + kid$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    If Not rtmp.EOF Then
      If trm(rtmp!email) <> "" Then rc$ = rc$ + trm(rtmp!email) + ","
    End If
  End If
  If Not isfieldmissing("opt_allenummern", "id") Then
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rrr = form1.adoopen(rtmp, "SELECT num FROM opt_allenummern where kid='" + kid$ + "' and numtyp='email'", adoc, adOpenDynamic, adLockReadOnly, "", "")
    If rrr = 0 Then
      While Not rtmp.EOF
        If trm(rtmp!num) <> "" Then rc$ = rc$ + trm(rtmp!num) + ","
        rtmp.MoveNext
      Wend
    End If
  End If
Else
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT email FROM adresse where id='" + lic$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    If Not rtmp.EOF Then
      If trm(rtmp!email) <> "" Then rc$ = rc$ + trm(rtmp!email) + ","
    End If
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, "SELECT email FROM kontakt where vid='" + lic$ + "'", adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 Then
    While Not rtmp.EOF
      If trm(rtmp!email) <> "" Then rc$ = rc$ + trm(rtmp!email) + ","
      rtmp.MoveNext
    Wend
  End If
  If Not isfieldmissing("opt_allenummern", "id") Then
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rrr = form1.adoopen(rtmp, "SELECT num FROM opt_allenummern where vid='" + lic$ + "' and numtyp='email'", adoc, adOpenDynamic, adLockReadOnly, "", "")
    If rrr = 0 Then
      While Not rtmp.EOF
        If trm(rtmp!num) <> "" Then rc$ = rc$ + trm(rtmp!num) + ","
        rtmp.MoveNext
      Wend
    End If
  End If
End If
allmailadresses = rc$
End Function

Function xkurs(dtg$, whr$, amount As Double) As Double
Dim mwhr$, xbhfremd$, hK As Double, hkdat As String, s1 As Double
Dim rrr

xkurs = amount
xbhfremd = ""
mwhr$ = form1.getusersetting("MeineWaehrung", transe("€"))
If whr$ <> "" And mwhr$ <> whr$ Then
  hK = var2dbl(strrepl(kursvom(whr$, dtg$), ".", ","))
  hkdat = kursdatum(whr$, dtg$)
  hK = CCur(hK)
  If hK = 0 Then hK = 10000000
  On Error Resume Next
  s1 = amount / hK
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then xkurs = s1
  xbhfremd = "(" + transe("Kurs") + ": " & trm(hK) & " " & mwhr$ & "/" & whr$ & " " + transe("am") + " " & hkdat & ")"
End If
xkurs_publicratestring = xbhfremd
End Function

Public Function cloudcreateadr(id$, n$, ownr$) As Boolean
Dim c$, i As Integer, r As ADODB.Recordset, rrr, kid$, s$, dtg$
Dim sid$, sidk$, sidp%, sida$, lo$, adrbkid$

cloudcreateadr = False
If Not cloud Then Exit Function

c$ = "select share_id,share_name,attribute_name from turba_sharesng where share_owner='" + ownr$ + "' and attribute_name='Agencyprof'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
If r.EOF Then
  cloud = False
  Exit Function
End If
adrbkid$ = r!share_name
r.Close
sid$ = trm(id$): sidk$ = "-1": sidp% = InStr(sid$, "{"): sida$ = sid$
If sidp% > 0 Then
  sidk$ = trm(Left(sid$, sidp% - 1))
  sida$ = trm(Mid(sid$, sidp% + 1))
  id$ = cut_d1(sida$, "}")
End If
      If turba_nosuchobject(adrbkid$, id$) Then
        c$ = "insert into turba_objects (object_id,owner_id,object_type,object_uid,object_lastname,object_alias) "
        c$ = c$ + "values('" + mkkey(24) + "',"
        c$ = c$ + "'" + adrbkid$ + "','Object',"
        'c$ = c$ + "'" + ownr$ + "','Object',"
        c$ = c$ + "'" + strrepl(datum2sql(Date), "-", "") + strrepl(Time, ":", "") + "." + mkkey(24) + "@" + cloudserver$ + "',"
        c$ = c$ + "'" + n$ + "',"
        dtg$ = strrepl(datum2sql(Date), "-", "") + strrepl(trm(Time), ":", "")
        c$ = c$ + "'" + id$ + "|00000000000000')"
        Call xhorde(c$)
      End If
'    Else
'      c$ = s$
'      lo$ = "SELECT stand as wert from adresse where id='" + id$ + "'"
'      s$ = get1erg(lo$)
'      If InStr(c$, "|") > 0 Then
'        c$ = cut_d2bis(c$, "|")
'        s$ = strrepl(datum2sql(word1(s$)), "-", "") + strrepl(word2bis(s$), ":", "")
'Debug.Print c$ + " vs " + s$
'        If c$ > s$ Then
'          Exit Function
'        End If
'      End If
'    End If

cloudcreateadr = True
  If Not form1.cloudworker(ownr$) And sidk$ = "-1" Then Exit Function
  c$ = "select id,Name from kontakt where vid='" + id$ + "'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  r.Open c$, adoc, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    While Not r.EOF
      s$ = trm(r!name)
      If s$ = sidk$ Or form1.cloudworker(ownr$) Then
        s$ = s$ + "{" + id$ + "}"
        c$ = LCase(typesof(s$))
        If InStr(c$, "|dummy|") = 0 And InStr(c$, "|internal|") = 0 And (InStr(c$, "|person|") > 0 Or InStr(c$, "|firma|") > 0 Or InStr(c$, "|saal|") > 0) Then
          kid$ = trm(r!id)
          Call cloudcreatekontakt(kid$, ownr$)
          Call kontakt2cloud(kid$)
          Exit Function
        End If
      End If
      r.MoveNext
    Wend
  End If

End Function

Public Function turbashareid(ownr$) As String
Dim c$, i As Integer, r As ADODB.Recordset, rrr, kid$, s$, dtg$
Dim sid$, sidk$, sidp%, sida$, lo$, adrbkid$

turbashareid = ""
c$ = "select share_name from turba_sharesng where share_owner='" + ownr$ + "' and attribute_name='Agencyprof'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
r.Open c$, form1.clddb, adOpenDynamic, adLockReadOnly
If r.EOF Then
  cloud = False
  Exit Function
End If
turbashareid = r!share_name
r.Close

End Function
Public Sub cloudcreatekontakt(id$, ownr$)
Dim c$, i As Integer, dtg$, tid$, hid$

If Not cloud Then Exit Sub
    
    tid$ = turbashareid(ownr$)
    hid$ = form1.get_kontaktname_by_id(id$) + "{" + form1.getadridbykontaktid(id$) + "}"
      If turba_nosuchobject(tid$, hid$) Then
        dtg$ = strrepl(datum2sql(Date), "-", "") + strrepl(trm(Time), ":", "")
        c$ = "insert into turba_objects (object_id,owner_id,object_type,object_uid,object_alias) "
        c$ = c$ + "values('" + mkkey(24) + "',"
        c$ = c$ + "'" + tid$ + "','Object',"
        c$ = c$ + "'" + strrepl(datum2sql(Date), "-", "") + strrepl(Time, ":", "") + "." + mkkey(24) + "@" + cloudserver$ + "',"
        c$ = c$ + "'" + hid$ + "|" + dtg$ + "')"
        Call xhorde(c$)
      End If


End Sub

Public Sub upd_turba_objects(id$, f$, w$)
Dim lo$, i, rrr

lo$ = "update turba_objects set " + f$ + "='" + w$ + "' where object_alias like '" + id$ + "|%'"
Call qhorde(lo$)

End Sub

Public Sub qhorde(cmd$)
Dim fn$, o%, rrr

If form1.getusersetting("serverfeedshorde", "nein") = "ja" Then
  fn$ = newcloudqfilex()
Else
  fn$ = newcloudqfile()
End If
If fn$ = "" Then Exit Sub
o% = FreeFile()
On Error Resume Next
Open fn$ For Append As #o%
rrr = Err
If rrr = 0 Then
  Print #o%, cmd$
  Close #o%
End If
End Sub

Public Sub xhorde(qdfTemp As String)
  Dim rrr, o%

  If shwled Then
    cb1.BackColor = RGB(255, 255, 0)
    cb1.Cls
  End If
  DoEvents
'  o% = FreeFile
'  On Error Resume Next
'  Open "c:\agencyprof\hordelog.txt" For Append As #o%
'  rrr = Err
'  On Error GoTo 0
'  If rrr = 0 Then
'    Print #o%, qdfTemp
'    Close #o%
'  End If
  On Error Resume Next
  clddb.Execute qdfTemp
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
'    MsgBox "Fehler beim Schreiben in die Replikationsdatenbank:" + vbCrLf + qdfTemp
'    connok = False
     Call form1.dbg2f("Fehlernummer: " & rrr & vbCr & Error$(rrr) & "statement=" & vbCrLf & qdfTemp)
  End If
  If shwled Then
    cb1.BackColor = RGB(0, 255, 0)
    cb1.Cls
  End If
  DoEvents
End Sub

Public Function cloudadressbuch(wem As String) As String
Dim c$, i As Integer, r As ADODB.Recordset, share_name$, rrr
  
  cloudadressbuch = ""
  c$ = "select share_name from turba_sharesng where share_owner='" + wem + "' and attribute_name='Agencyprof'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  r.Open c$, clddb, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    If r.EOF Then
      cloud = False
      Exit Function
    End If
    cloudadressbuch = r!share_name
  End If

End Function

Public Function cloudcreateevnt(id$, userbyid$) As String
Dim c$, i As Integer, r As ADODB.Recordset, share_name$, rrr, lngid As Long, tdiff

cloudcreateevnt = ""
'Debug.Print "creating event " + id$ + " for " + userbyid$
If Not cloud Then Exit Function
    
  On Error Resume Next
  tdiff = DateDiff("s", "1.1.1970 00:00", Date + Time)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then tdiff = DateDiff("s", "1/1/1970 00:00", Date + Time)
  c$ = "select share_name,share_owner,attribute_name from kronolith_sharesng where share_owner='" + userbyid$ + "' and attribute_name='Agencyprof'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  r.Open c$, clddb, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    If r.EOF Then
      cloud = False
      Exit Function
    End If
    share_name$ = r!share_name
  Else
    Exit Function
  End If
  cloudcreateevnt = share_name$
  c$ = "select calendar_id from kronolith_events where calendar_id='" + share_name$ + "' and event_creator_id='" + id$ + "@localnet'"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  On Error Resume Next
  r.Open c$, clddb, adOpenDynamic, adLockReadOnly
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then
    Exit Function
  End If
  
    c$ = "delete from kronolith_events where calendar_id='" + share_name$ + "' and event_creator_id='" + id$ + "@localnet'"
'Debug.Print c$
    Call qhorde(c$)
    c$ = "insert into kronolith_events (event_id,event_uid,calendar_id,event_creator_id) "
    lngid = Date + Time
    c$ = c$ + "values('" + trm(tdiff) + mkkey(14) + "',"
    c$ = c$ + "'" + strrepl(datum2sql(Date), "-", "") + strrepl(Time, ":", "") + "." + mkkey(24) + "@" + cloudserver$ + "',"
    c$ = c$ + "'" + share_name$ + "',"
    c$ = c$ + "'" + id$ + "@localnet')"
'Debug.Print c$
    Call qhorde(c$)

End Function

Public Sub upd_krono_objects(id$, f$, w$)
Dim lo$

If InStr(w$, "DATE_") <> 1 Then
  lo$ = "update kronolith_events set " + f$ + "='" + w$ + "' where event_creator_id='" + id$ + "@localnet'"
Else
  lo$ = "update kronolith_events set " + f$ + "=" + w$ + " where event_creator_id='" + id$ + "@localnet'"
End If
Call qhorde(lo$)

End Sub

Public Sub upd_many_krono_ids(id$, f$, w$)
Dim lo$, rrr
Dim rtmp As ADODB.Recordset

If Not cloud Then Exit Sub

lo$ = "SELECT calendar_id,event_id from kronolith_events where event_creator_id='" + id$ + "@localnet'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, lo$, clddb, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  While Not rtmp.EOF
    lo$ = "update kronolith_events set event_id='" + w$ + Left(trm(rtmp!calendar_id), 12) + "' where calendar_id='" + trm(rtmp!calendar_id) + "' and event_creator_id='" + id$ + "@localnet'"
'Debug.Print lo$
    Call qhorde(lo$)
    rtmp.MoveNext
  Wend
End If
End Sub

Public Sub upd_krono_managementobjects(id$, f$, w$)
Dim lo$, restr$, l$, h$

restr$ = "": l$ = supershares_krono
While l$ <> ""
  h$ = cut_d1(l$, "|"): l$ = cut_d2bis(l$, "|")
  If h$ <> "" Then
    If restr$ <> "" Then restr$ = restr$ + " or "
    restr$ = restr$ + "calendar_id='" + h$ + "'"
  End If
Wend
lo$ = "update kronolith_events set " + f$ + "='" + w$ + "' where event_creator_id='" + id$ + "@localnet' and (" + restr$ + ")"
'Debug.Print lo$
Call qhorde(lo$)

End Sub

Public Sub adr2cloud(id$)
Dim lo$, rrr, w$
Dim c$, i As Integer
Dim rtmp As ADODB.Recordset

If Not cloud Then Exit Sub

lo$ = "SELECT * from adresse where id='" + id$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, lo$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not rtmp.EOF Then

w$ = trm(rtmp!name): c$ = "object_lastname": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!strasse): c$ = "object_homestreet": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!ort): c$ = "object_homecity": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!fax): c$ = "object_fax": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!tel): c$ = "object_homephone": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!handy): c$ = "object_cellphone": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!email): c$ = "object_email": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!url): c$ = "object_url": Call upd_turba_objects(id$, c$, w$)
'w$ = trm(rtmp!hinweise): c$ = "object_notes": Call upd_turba_objects(id$, c$, w$,mde$)
w$ = trm(rtmp!plz): c$ = "object_homepostalcode": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!land): c$ = "object_homecountry": Call upd_turba_objects(id$, c$, w$)
w$ = trm(rtmp!postanrede): c$ = "object_nameprefix": Call upd_turba_objects(id$, c$, w$)
w$ = id$ + "|" + strrepl(datum2sql(Date), "-", "") + strrepl(trm(Time), ":", "")
c$ = "object_alias": Call upd_turba_objects(id$, c$, w$)

  End If
End If

End Sub

Public Sub kontakt2cloud(id$)
Dim lo$, rrr, w$, halias As String, tfake As String
Dim c$, i As Integer
Dim rtmp As ADODB.Recordset

If Not cloud Then Exit Sub

lo$ = "SELECT * from kontakt where id='" + id$ + "'"
tfake = strrepl(datum2sql(Date), "-", "") + strrepl(trm(Time), ":", "")
halias = form1.get_kontaktname_by_id(id$) + "{" + form1.getadridbykontaktid(id$) + "}"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, lo$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not rtmp.EOF Then

    w$ = trm(rtmp!name): c$ = "object_lastname": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!strasse): c$ = "object_homestreet": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!ort): c$ = "object_homecity": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!fax): c$ = "object_fax": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!tel): c$ = "object_homephone": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!handy): c$ = "object_cellphone": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!email): c$ = "object_email": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!url): c$ = "object_url": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!plz): c$ = "object_homepostalcode": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!lkz): c$ = "object_homecountry": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!postanrede): c$ = "object_nameprefix": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!vid): c$ = "object_company": Call upd_turba_objects(halias, c$, w$)
    w$ = trm(rtmp!Position): c$ = "object_role": Call upd_turba_objects(halias, c$, w$)
    w$ = halias + "|" + tfake
    c$ = "object_alias": Call upd_turba_objects(halias, c$, w$)

    lo$ = "SELECT * from adresse where id='" + rtmp!vid + "'"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
    rrr = form1.adoopen(rtmp, lo$, adoc, adOpenDynamic, adLockReadOnly, "", "")
    If rrr = 0 Then
      If Not rtmp.EOF Then
        w$ = trm(rtmp!strasse): c$ = "object_workstreet": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!ort): c$ = "object_workcity": Call upd_turba_objects(id$, c$, w$)
        'nosuch: w$ = trm(rtmp!fax): c$ = "object_fax": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!tel): c$ = "object_workphone": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!handy): c$ = "object_workphone2": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!email): c$ = "object_workemail": Call upd_turba_objects(id$, c$, w$)
        'nosuch: w$ = trm(rtmp!url): c$ = "object_url": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!plz): c$ = "object_workpostalcode": Call upd_turba_objects(id$, c$, w$)
        w$ = trm(rtmp!land): c$ = "object_workcountry": Call upd_turba_objects(id$, c$, w$)
      End If
    End If
  End If
End If

End Sub

Public Sub event2cloudremove(id$)
Dim lo$, rrr
Dim c$
Dim rtmp As ADODB.Recordset

If Not cloud Then Exit Sub
lo$ = "SELECT * from auftritt where id='" + id$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, lo$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not rtmp.EOF Then
    c$ = "delete from kronolith_events where event_creator_id='" + id$ + "@localnet'"
'Debug.Print c$
    Call qhorde(c$)
  End If
End If

End Sub
Public Sub event2cloud(id$)
Dim lo$, rrr, w$, desc$, usr$, typ$, dtg$, adrb$, aid$, gmtdiff$, timedirect$, prp$, prpn$
Dim c$, i As Integer, aelist As String, aalist As String, aalistlimited As String
Dim rtmp As ADODB.Recordset, t1$, t2$, t3$, tx1$, ttt As String, desclimited$, adatum As String
Dim ftst As ADODB.Recordset, axlist As String, tlist As String, tdiff
Dim fi As ADODB.Recordset, j, sharethis$, wht$, pdelim$
Dim fa As ADODB.Recordset, favalid As Boolean, nocreats As Boolean
Dim ra As ADODB.Recordset, r As ADODB.Recordset

If Not cloud Then Exit Sub
Debug.Print "event " + id$
  
  On Error Resume Next
  tdiff = DateDiff("s", "1.1.1970 00:00", Date + Time)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then tdiff = DateDiff("s", "1/1/1970 00:00", Date + Time)

lo$ = "SELECT * from auftritt where id='" + id$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, lo$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not rtmp.EOF Then
    If rtmp!astatus = 4 Then
      Call event2cloudremove(id$)
      Exit Sub
    End If
    typ$ = rtmp!auftrittstyp

'pushing to cloud
aelist = ""
nocreats = True
aalistlimited = ""
  c$ = "select * from auftrittsfelder where typ='" + typ$ + "'"
  c$ = c$ + " order by position"
  Set ftst = New ADODB.Recordset
  ftst.CursorLocation = adUseServer
  rrr = form1.adoopen(ftst, c$, adoc, adOpenDynamic, adLockReadOnly, "", "")
  If rrr = 0 And (Not ftst.EOF) Then
    aelist = "": aalist = ""
    While Not ftst.EOF
      t1$ = cut_d1(trm(ftst!feldname), ".")
      t2$ = cut_d2bis(trm(ftst!feldname), ".")
      t3$ = cut_d2bis(t2$, ".")
      t2$ = cut_d1(t2$, ".")
      If t1$ = "programm" Then
        t1$ = t2$: t2$ = ""
      End If
      If Not form1.isfieldmissing("auftrittsfelder", "opthordeshare") Then
        If ftst!opthordeshare = 1 Then
          sharethis$ = ""
          If Not form1.isfieldmissing("auftrittsfelder", "opthordesharewhat") Then
            wht$ = trm(ftst!opthordesharewhat)
            If wht$ <> "" Then
              sharethis$ = "(" + wht$ + ")"
            End If
          End If
          aalistlimited = aalistlimited + "|" + t2$
          aelist = aelist + "|" + t2$
          aalist = aalist + "|" + sharethis$ + t2$
          If t2$ = "" Or t1$ = "Programm" Then
            c$ = LCase(t1$)
            If InStr(c$, "auszahl") = 0 And InStr(c$, "klauseln") = 0 And InStr(c$, "provision") = 0 And InStr(c$, "honorar") = 0 And InStr(c$, "intern") = 0 Then
              aalist = aalist + "|" + t1$
              If Not form1.isfieldmissing("auftrittsfelder", "opthordeshare") Then
                If ftst!opthordeshare = 1 Then aalistlimited = aalistlimited + "|" + t1$ + "|"
              End If
            End If
          End If
        End If
      End If
      ftst.MoveNext
    Wend
    If aelist = "" And aalist = "" Then Exit Sub
    desclimited$ = ""
    If aelist <> "" Then aelist = aelist + "|"
    If aalist <> "" Then aalist = aalist + "|"
    If aalistlimited <> "" Then aalistlimited = aalistlimited + "|"
    desc$ = ""
          favalid = False
          c$ = "select * from usr_" + LCase(typ$) + " where id='" + id$ + "'"
          Set fa = New ADODB.Recordset
          fa.CursorLocation = adUseServer
          rrr = form1.adoopen(fa, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          If rrr = 0 And (Not fa.EOF) Then favalid = True

    While aalist <> ""
        DoEvents
        t1$ = cut_d1(aalist, "|"): aalist = cut_d2bis(aalist, "|")
        If t1$ <> "" Then
          t2$ = ""
          If favalid Then
            wht$ = ""
            If Left(t1$, 1) = "(" Then
              wht$ = cut_d2bis(cut_d1(t1$, ")"), "(")
              t1$ = cut_d2bis(t1$, ")")
            End If
            For j = 0 To fa.Fields.Count - 1
              If LCase(t1$) = LCase(fa.Fields(j).name) Then
                t2$ = trm(fa.Fields(j).value)
                Exit For
              End If
            Next j
            If wht$ <> "" Then
              w$ = ""
              While wht$ <> ""
                c$ = cut_d1(wht$, ","): wht$ = cut_d2bis(wht$, ",")
                prpn$ = c$: prp$ = c$
                If InStr(c$, "=") Then
                  prp$ = cut_d1(c$, "=")
                  prpn$ = cut_d2bis(c$, "=")
                End If
                pdelim$ = ", ": If LCase(prp$) = "ort" Then pdelim$ = " "
                prp$ = getAdrProperty(t2$, prp$)
                If prp$ <> "" Then
                  If w$ <> "" Then w$ = w$ + pdelim$
                  w$ = w$ + trm(prpn$ + " " + prp$)
                End If
              Wend
              t2$ = trm(w$)
            End If
          Else
            c$ = "select FeldDaten from auftritthigru where FeldName='" + t1$ + "' and auftrittsid='" + id$ + "'"
            Set fi = New ADODB.Recordset
            fi.CursorLocation = adUseServer
            rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
            If rrr = 0 And (Not fi.EOF) Then
              If trm(fi!felddaten) <> "" Then t2$ = trm(fi!felddaten)
            End If
          End If
          If LCase(t1$) = "programm" And t2$ <> "" Then
            t2$ = getwerke(t2$)
          End If
          If t2$ <> "" Then
            If desc$ <> "" Then desc$ = desc$ + "\r\n"
            desc$ = desc$ + UcaseFirstLetter(transe(t1$)) + ": " + strrepl(strrepl(trm(t2$), "|", ", "), vbCrLf, "\r\n")
            If InStr(aalistlimited$, "|" + t1$ + "|") > 0 Then
              If desclimited$ <> "" Then desclimited$ = desclimited$ + "\r\n"
              desclimited$ = desclimited$ + UcaseFirstLetter(transe(t1$)) + ": " + strrepl(strrepl(trm(t2$), "|", ", "), vbCrLf, "\r\n")
            End If
          Else
            If InStr(aelist, "|" + t1$ + "|") > 0 Then
'Debug.Print "igno: "; t1$
              aelist = strrepl(aelist, "|" + t1$ + "|", "|||")
            End If
          End If
      End If
    Wend
    While InStr(aelist, "||") > 0: aelist = strrepl(aelist, "||", "|"): Wend
    
    pgb1.Min = 0
    
    If aelist = "" Then aelist = form1.cloudmanager + form1.cloudstaff
    j = Len(aelist)
    If j <= 1 Then
      j = 1
      pgb1.Visible = False
    End If
    pgb1.Max = j: pgb1.value = pgb1.Max:  pgb1.Visible = True
    If form1.cloud And aelist <> "" Then
      axlist = aelist
      While aelist <> ""
        pgb1.value = Len(aelist)
        DoEvents
        t1$ = cut_d1(aelist, "|"): aelist = cut_d2bis(aelist, "|")
        If t1$ <> "" Then
          If favalid Then
            For j = 0 To fa.Fields.Count - 1
              If LCase(t1$) = LCase(fa.Fields(j).name) Then
                t2$ = trm(fa.Fields(j).value)
                Exit For
              End If
            Next j
          Else
            c$ = "select FeldDaten from auftritthigru where FeldName='" + t1$ + "' and auftrittsid='" + id$ + "'"
            Set fi = New ADODB.Recordset
            fi.CursorLocation = adUseServer
            rrr = form1.adoopen(fi, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
            If rrr = 0 And (Not fi.EOF) Then t2$ = trm(fi!felddaten)
          End If
 
            If InStr(t2$, "{") > 0 Then
              t3$ = form1.getadridbykontaktid(t2$)
            Else
              t3$ = t2$
            End If
            c$ = "select FeldDaten as wert from auftritthigru where FeldName='cloud' and auftrittsid='" + t3$ + "' and auftrittstyp='webcal'"
'            Debug.Print c$
            usr$ = form1.get1erg(c$)
            If usr$ <> "" Then
              c$ = form1.cloudcreateevnt(id$, usr$)
              nocreats = False
'              adrb$ = cloudadressbuch(usr$)
              adrb$ = usr$
              If adrb$ <> "" Then
                tlist = axlist
                While tlist <> ""
                  DoEvents
                  tx1$ = cut_d1(tlist, "|"): tlist = cut_d2bis(tlist, "|")
                  If tx1$ <> "" Then
                    c$ = "select FeldDaten as wert from auftritthigru where FeldName='" + tx1$ + "' and auftrittsid='" + id$ + "'"
                    aid$ = strrepl(form1.get1erg(c$), vbCrLf, "")
                    ttt$ = "x"
                    While ttt$ <> ""
                      ttt$ = cut_d1(aid$, "|"): aid$ = cut_d2bis(aid$, "|")
                      If ttt$ <> "" Then
                        DoEvents
                        c$ = ttt$ + "|" + adrb$
                        For j = 0 To hordex.ListCount - 1
                          If hordex.List(j) = c$ Then Exit For
                        Next j
                        If j >= hordex.ListCount Or hordex.ListCount = 0 Then
                          Call form1.add2hordex(c$)
                        End If
                        DoEvents
                      End If
                    Wend
                  End If
                Wend
              End If
            End If
          End If


      Wend
pgb1.Visible = False
                
                If Not nocreats Then
                Call upd_krono_objects(id$, "event_title", trm(trm(rtmp!ort) + " " + form1.get_hordeabkz(typ$)) + " " + trm(rtmp!bezeichnung))
                dtg$ = trm(strrepl(trm(rtmp!zeit), "Uhr", ""))
                dtg$ = strrepl(dtg$, "uhr", "")
                If dtg$ = "" Then dtg$ = "00:00:00"
                If InStr(dtg$, ":") = 0 Then
                  If Len(dtg$) < 4 Then dtg$ = "0" + dtg$
                  dtg$ = Left(dtg$, 2) + ":" + Mid$(dtg$, 2, 2) + ":00"
                Else
                  If Mid$(dtg$, 3, 1) <> ":" Then dtg$ = "0" + dtg$
                  If Len(dtg$) = 5 Then dtg$ = dtg$ + ":00"
                End If
                adatum = datum2sql(rtmp!datum)
                dtg$ = adatum + " " + dtg$
                Call upd_krono_objects(id$, "event_start", dtg$)
                timedirect = "DATE_ADD"
                gmtdiff$ = "ungesetzt"
                If ist_sommerzeit(adatum) Then
                  gmtdiff$ = getusersetting("GMTSZ_add", "ungesetzt")
                End If
                If gmtdiff$ = "ungesetzt" Then
                  gmtdiff$ = getusersetting("GMT_add", "0")
                End If
                If Left(gmtdiff$, 1) = "-" Then
                  timedirect = "DATE_SUB": gmtdiff$ = Mid(gmtdiff$, 2)
                End If
                If gmtdiff$ <> "0" Then
                  Call upd_krono_objects(id$, "event_start", timedirect$ + "(event_start,INTERVAL " + gmtdiff$ + " SECOND)")
                End If
                dtg$ = auftrittsende(id$, "")
                If dtg$ = "" Then dtg$ = "23:59:00"
                If InStr(dtg$, ":") = 0 Then
                  If Len(dtg$) < 4 Then dtg$ = "0" + dtg$
                  dtg$ = Left(dtg$, 2) + ":" + Mid$(dtg$, 2, 2) + ":00"
                Else
                  If Mid$(dtg$, 3, 1) <> ":" Then dtg$ = "0" + dtg$
                  If Len(dtg$) = 5 Then dtg$ = dtg$ + ":00"
                End If
                dtg$ = datum2sql(rtmp!datum) + " " + dtg$
                Call upd_krono_objects(id$, "event_end", dtg$)
                If gmtdiff$ <> "0" Then Call upd_krono_objects(id$, "event_end", timedirect$ + "(event_end,INTERVAL " + gmtdiff$ + " SECOND)")
                Call upd_krono_objects(id$, "event_location", trm(rtmp!ort))
                If desclimited$ <> "" Then Call upd_krono_objects(id$, "event_description", desclimited$)
                If desc$ <> "" Then Call upd_krono_managementobjects(id$, "event_description", desc$)
                Call upd_krono_objects(id$, "event_modified", trm(tdiff) + " " + form1.get_atabkz(typ$) + " " + trm(rtmp!bezeichnung))
                Call upd_many_krono_ids(id$, "event_id", trm(tdiff))
                End If
    End If
  End If


  End If
End If
End Sub

Public Function newcloudqfile() As String
Dim o%, fn$, rrr, i%

newcloudqfile = ""
fn$ = s0d$ & "\" + docs() + "\" + uId$ + "\cloudq"
On Error Resume Next
MkDir (fn$)
On Error GoTo 0
o% = FreeFile
fn$ = s0d$ & "\" + docs() + "\" & uId$ & "\cloudq\tst.tst"
On Error Resume Next
Open fn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  cloud = False
  
Else
  Close #o%
  Kill fn$
End If
i% = 1000
Do
  DoEvents
  i% = i% + 1
  fn$ = s0d$ & "\" + docs() + "\" & uId$ + "\cloudq\q" & Left$(Date, 2) & Mid$(Date, 4, 2) & strrepl(trm(Time), ":", "") & uId$ & trm(str$(i%)) & ".sql"
Loop Until exist(fn$) = 0 Or i > 10000
If i > 10000 Then
  cloud = False
Else
  newcloudqfile = fn$
End If

End Function

Public Function newcloudqfilex() As String
Dim o%, fn$, rrr, i%, tn$, dn$

newcloudqfilex = ""
dn$ = s0d$ & "\" + docs() + "\" + uId$ + "\cloudq"
On Error Resume Next
MkDir (dn$)
On Error GoTo 0
dn$ = dn$ + "\serverfeed"
On Error Resume Next
MkDir (dn$)
On Error GoTo 0
o% = FreeFile
tn$ = dn$ + "\tst.tst"
On Error Resume Next
Open tn$ For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  cloud = False
Else
  Close #o%
  Kill tn$
End If
i% = 10000
Do
  DoEvents
  i% = i% + 1
  fn$ = dn$ + "\q" & Left$(Date, 2) & Mid$(Date, 4, 2) & strrepl(trm(Time), ":", "") & uId$ & trm(str$(i%)) & ".sql"
Loop Until exist(fn$) = 0 Or i > 20000
If i > 20000 Then
  cloud = False
Else
  newcloudqfilex = fn$
End If

End Function

Function cloudworker(who$) As Boolean
cloudworker = False
If InStr(cloudstaff, "|" + who$ + "|") > 0 Or InStr(cloudmanager, "|" + who$ + "|") > 0 Then cloudworker = True
End Function

Public Function APUsernameByAddressID(adrid$) As String
Dim sid$, sida$, sidk$, sidp%, c$, id$, s$
Dim rtmp As ADODB.Recordset, rrr

APUsernameByAddressID = adrid$
sid$ = trm(adrid$): sidk$ = "-1": sidp% = InStr(sid$, "{"): sida$ = sid$
If sidp% > 0 Then
  sidk$ = trm(Left(sid$, sidp% - 1))
  sida$ = trm(Mid(sid$, sidp% + 1))
  id$ = cut_d1(sida$, "}")
  sidk$ = getkontaktnamebyid(get_kontaktid_by_name(id$, sidk$))
Else
  sidk$ = getnamebyid(sida$)
End If

s$ = "select ID from benutzerdaten where Name='" + sidk$ + "'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, s$, adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If Not rtmp.EOF Then
    APUsernameByAddressID = trm(rtmp!id)
  End If
End If

End Function

Public Function mname(z) As String
mname = mnams$(z)
End Function

Public Sub sqlin1file(fn$)
Dim o%, sq$, rrr, l$

o% = FreeFile
Open fn$ For Input As #o%
While Not EOF(o%)
  sq$ = ""
  Do
    On Error Resume Next
    Line Input #o%, l$
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then
      If rrr = 62 Then GoTo errrrx1
      MsgBox "Fehler Nr." & rrr & " beim Import" & vbCrLf & Error$(rrr) & vbCrLf & fn$
      End
    End If
    If Left(l$, 2) <> "--" Then
      If Len(sq$) > 0 Then sq$ = sq$ & vbCrLf
      sq$ = sq$ + l$
    End If
  Loop Until Right$(trm(sq$), 1) = ";"
  If trm(sq$) <> ";" Then
    If InStr(LCase$(sq$), "insert into ") = 1 Then
      form1.err_dupok% = 1
    End If
    DoEvents
    Call form1.sqlqry(sq$)
    err_dupok% = 0
  End If
Wend
errrrx1:
Close #o%
On Error Resume Next
Kill fn$
On Error GoTo 0
End Sub

Function refeddates(tpid$) As Integer
Dim c$, i%
Dim rtmp As ADODB.Recordset, rrr

refeddates = 0
If isfieldmissing("opt_othertplans", "id") Then Exit Function
i% = 0
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT aid FROM opt_othertplans where tpid ='" + tpid$ & "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr <> 0 Then
  Exit Function
End If
While Not rtmp.EOF
  i% = i% + 1
  rtmp.MoveNext
Wend
refeddates = i%

End Function

Sub debugwarn()
Dim addwarn As String

If currentlanguage = "de" Then
  If LCase(getusersetting("killlogonexit", "ja")) <> "ja" Then
    addwarn = vbCrLf + "Außerdem ist killogonexit aus (schlecht), die Logadtei wird permanent größer."
  Else
    addwarn = vbCrLf + "killogonexit ist an (gut), die Logadtei wird gelöscht wenn Agencyprof 'normal' beendet wird."
  End If
  Call MsgBox("debug2file ist eingeschaltet, was Agencyprof besonders nach längerer Laufzeit verlangsamen wird." + vbCrLf + "Die Logdatei " + s0d$ + "\debug2file_" + uId$ + ".txt kann jederzeit gelöscht werden, sie wird sofort neu erstellt." + addwarn, vbCritical, "Achtung, Debugging ...")
Else
  If LCase(getusersetting("killlogonexit", "ja")) <> "ja" Then
    addwarn = vbCrLf + "Additionally killogonexit is off (bad), the logfile will get bigger and bigger ..."
  Else
    addwarn = vbCrLf + "killogonexit is off (good), the logfile will be deleted when Agencyprof exits 'normaly'."
  End If
  Call MsgBox("debug2file ist on, that will slow down Agencyprof after some time running." + vbCrLf + "The logfile " + s0d$ + "\debug2file_" + uId$ + ".txt can be deleted anytime, it will be created again as soon as needed." + addwarn, vbCritical, "Warning, Debugging ...")
End If

End Sub

Sub chk_backslashhandler(hdl As String)
Dim erg As String, bhdl As String, ask%, qst$, shd As String

bhdl = getusersetting("backslashhandler", "")
Call setusersetting("tstbckslsh", "12\\34")
erg = getusersetting("tstbckslsh", "")
Call sqlqry("delete from sysvars where owner='sysvar_" & uId$ & "_tstbckslsh'")
If erg <> "12\\34" Then
  shd = "aus": If bhdl <> "an" Then shd = "an"
  qst$ = "incorrect setting: backslashhandler=" + bhdl + vbCrLf + "correct: backslashhandler=" + shd
  ask% = MsgBox(qst$, vbYesNo + vbCritical + vbDefaultButton2, "Change Setting?")
  If ask% = vbYes Then
    Call setusersetting("backslashhandler", shd)
    MsgBox "Setting changed." + vbCrLf + "Please restart Agencyprof."
    End
  End If
End If
End Sub

Public Function UseBrowser() As String
Dim try As String

UseBrowser = ""
try = getusersetting("UseBrowser", "")

If nexist(try) Then try = ""
If try = "" Then try = FindBrowser()
UseBrowser = try

End Function

Function iamdemo() As Boolean
iamdemo = False
If InStr(LCase(form1.computername), "wapdemo") = 1 And Len(form1.computername) = 8 Then iamdemo = True
End Function
Public Sub updateme()
Dim nfn$, url$, X As Boolean, o%, l0$, l1$, rrr, updl$
Dim apv As Long, aplv As Long, rstart As Boolean
Dim xapv As Long, xaplv As Long


rstart = False
apv = App.Revision
aplv = hexstring2dec(bas_getAPLibVersion())

updl$ = ""
If Not nexist(form1.s00dir() + "\AgencyprofRestart.exe") Then
  On Error Resume Next
  Kill form1.s00dir() + "\AgencyprofRestart.exe"
  On Error GoTo 0
End If
If nexist(form1.s00dir() + "\AgencyprofRestart.exe") Then
  'MsgBox "cannot find the new updater, downloading it ..."
  url$ = "http://www.agencyprof.de/download/update/AgencyprofRestart.exe"
  nfn$ = form1.s00dir() + "\AgencyprofRestart.exe"
  MousePointer = 11: DoEvents
  X = DownloadFileFromURL(url$, nfn$)
  MousePointer = 0: DoEvents
  If Not X Then
    MsgBox "Download failed, I give up."
    Exit Sub
  End If
  If nexist(form1.s00dir() + "\AgencyprofRestart.exe") Then
    MsgBox "... (still) cannot find the new updater, I give up."
    Exit Sub
  End If
End If
'check versions
'
url$ = "http://www.agencyprof.de/download/update/" & App.Major & "-" & App.Minor & "-Agencyprof1.ver"
nfn$ = form1.s00dir() + "\apversions.ver"
xapv = 0: xaplv = 0
If Not nexist(nfn$) Then
  On Error Resume Next
  Kill nfn$
  On Error GoTo 0
End If
MousePointer = 11: DoEvents
X = DownloadFileFromURL(url$, nfn$)
MousePointer = 0: DoEvents
If Not nexist(nfn$) Then
  On Error Resume Next
  o% = FreeFile
  Open nfn$ For Input As #o%
  Line Input #o%, l0$
  Line Input #o%, l1$
  Close #o%
  rrr = Err
  On Error GoTo 0
  If rrr = 0 Then
    On Error Resume Next
    xapv = CInt(l0$)
    xaplv = CInt(l1$)
    rrr = Err
    On Error GoTo 0
  End If
  If rrr <> 0 Then
    xapv = 0: xaplv = 0
  End If
  On Error Resume Next
  Kill nfn$
  On Error GoTo 0
End If
If xapv = 0 Then
  MsgBox "Cannot get public version numbers."
  Exit Sub
End If
'--------------
If xapv > apv Then
url$ = "http://www.agencyprof.de/download/update/Agencyprof1.exe"
nfn$ = form1.s00dir() + "\neu.Agencyprof1.exe"
If Not nexist(nfn$) Then
  On Error Resume Next
  Kill nfn$
  On Error GoTo 0
  If Not nexist(nfn$) Then
    MsgBox "cannot delete the old " + nfn$ + vbCrLf + "update failed"
    Exit Sub
  End If
End If
MousePointer = 11: DoEvents
X = DownloadFileFromURL(url$, nfn$)
MousePointer = 0: DoEvents
If Not X Then
  MsgBox "cannot download " + url$ + " to " + nfn$ + vbCrLf + "update failed"
  Exit Sub
Else
  rstart = True
End If
Else
  updl$ = "Agencprof1.exe: published #" + trm(xapv) + ", local #" + trm(apv)
End If
If xaplv > aplv Then
url$ = "http://www.agencyprof.de/download/update/agencyproflib.dll"
nfn$ = form1.s00dir() + "\neu.agencyproflib.dll"
If Not nexist(nfn$) Then
  On Error Resume Next
  Kill nfn$
  On Error GoTo 0
  If Not nexist(nfn$) Then
    MsgBox "cannot delete the old " + nfn$ + vbCrLf + "update failed"
    Exit Sub
  End If
End If
MousePointer = 11: DoEvents
X = DownloadFileFromURL(url$, nfn$)
MousePointer = 0: DoEvents
If Not X Then
  MsgBox "cannot download " + url$ + " to " + nfn$ + vbCrLf + "update failed"
  Exit Sub
Else
  rstart = True
End If
Else
  If updl$ <> "" Then updl$ = updl$ + vbCrLf
  updl$ = updl$ + "agencyproflib.dll: No update needed, published version " + trm(xaplv) + ", local version " + trm(aplv)
End If
If updl$ <> "" Then MsgBox updl$
If rstart Then
  MsgBox transe("restarting to finish update.")
  Call form1.unloadall
  DoEvents
  Call dbg2f("starte " + form1.s00dir() + "\AgencyprofRestart.exe")
  X = Shell(form1.s00dir() + "\AgencyprofRestart.exe", 1)
  End
Else
  MsgBox transe("No updates required.")
End If
End Sub

Public Function knownaddress(Address$) As Boolean
Dim i As Integer
Dim dh As ADODB.Recordset, rrr
Dim frome$, c$

  frome$ = Address$
  knownaddress = False
  If InStr(frome$, "<") > 0 Then
    frome$ = Mid$(frome$, InStr(frome$, "<") + 1)
    frome$ = Left$(frome$, InStr(frome$, ">") - 1)
  Else
    If InStr(trm(frome$), " ") > 0 Then
      frome$ = Left(trm(frome$), InStr(trm(frome$), " ") - 1)
      frome$ = trm(frome$)
    End If
  End If
  frome$ = strrepl(frome$, "'", "")
  c$ = "SELECT * FROM adresse where trim(lcase(email))='" + LCase(frome$) + "'"
  Set dh = New ADODB.Recordset
  dh.CursorLocation = adUseServer
rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "form1", "knownaddress")
  If rrr <> 0 Then
    knownaddress = True
    Exit Function
  End If
  If Not dh.EOF Then
    knownaddress = True
    Exit Function
  Else
    If Not form1.isfieldmissing("opt_allenummern", "id") Then
      c$ = "SELECT * FROM opt_allenummern where trim(lcase(num))='" + LCase(frome$) + "' and numtyp='email'"
      Set dh = New ADODB.Recordset
      dh.CursorLocation = adUseServer
      rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "form1", "knownaddress")
      If rrr <> 0 Then
        knownaddress = True
        Exit Function
      End If
      If Not dh.EOF Then
        knownaddress = True
        Exit Function
      End If
    End If
  End If
  c$ = "SELECT ID FROM kontakt where trim(lcase(email))='" + LCase(frome$) + "'"
  Set dh = New ADODB.Recordset
  dh.CursorLocation = adUseServer
  rrr = form1.adoopen(dh, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "form1", "knownaddress")
  If Not dh.EOF Then
    knownaddress = True
    Exit Function
  End If
  knownaddress = False
End Function

Public Function replicationfilename(cn As String) As String
Dim dbnme$, fn$, fn0$

dbnme$ = form1.getdbname()
fn$ = form1.s00dir + "\lreplik." + cn + "." + dbnme$ + ".dat"
If nexist(fn$) Then
  fn0$ = form1.s00dir + "\lreplik." + cn + ".dat"
  If Not nexist(fn0$) Then Call FileCopy(fn0$, fn$)
End If
replicationfilename = fn$
End Function

Public Function rdrep(felddaten$) As String
Dim rrr
Dim rprog As ADODB.Recordset, k$, dau$, d$
Dim stmp As ADODB.Recordset, rc$, wid$, sid$

Dim d2infile As String, d2insub As String
d2infile = "Form1": d2insub = "rdprog"
rc$ = ""
rdrep = rc$
    Set rprog = New ADODB.Recordset
    rprog.CursorLocation = adUseServer
rrr = form1.adoopen(rprog, "SELECT wid FROM opt_repertoire where neverever=0 and vid='" + felddaten$ + "' order by wid", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rprog.EOF
      wid$ = trm(rprog!wid): sid$ = ""
      If Left$(wid$, 4) = "SBZ:" Then
        sid$ = Mid$(wid$, 5)
        wid$ = form1.getsatzidbywerkid(sid$)
      End If
      k$ = form1.getkompvornamenamebywerkid(wid$)
      dau$ = form1.getdauerbywerkid(wid$): If sid$ <> "" Then dau$ = ""
      d$ = "(" + form1.getkompdatesbywerkid(wid$) + ")"
      If Left$(LCase$(k$), 7) = "pause p" Or Left$(LCase$(k$), 7) = "oder od" Then
          k$ = ""
          d$ = ""
      End If
      If k$ <> "" Then
        rc$ = rc$ & k$ & " " & d$ & ": "
      End If
      If sid$ = "" Then
        rc$ = rc$ & form1.getwerknamebyid("" & wid$ & "")
      Else
        rc$ = rc$ + form1.getsatznamebyid(sid$) + " " + transe("aus") + " " + form1.getwerknamebyid("" & wid$ & "")
      End If
      If trm(dau$) <> "" Then
        rc$ = rc$ & " (" & dau$
        If InStr(LCase(dau$), "min") = 0 Then rc$ = rc$ + " " + transe("Min.")
        rc$ = rc$ + ") "
      End If
      rc$ = rc$ & vbCrLf
      rprog.MoveNext
    Wend
rdrep = rc$
End Function

Public Sub add2hordex(ids$)
Dim i%
For i% = 0 To hordex.ListCount - 1
  If i% >= hordex.ListCount Then
    Exit Sub
  End If
  If hordex.List(i%) = ids$ Then
    Exit Sub
  End If
Next i%
hordex.AddItem ids$
End Sub

Function turba_nosuchobject(adrbkid$, id$) As Boolean
  Dim c$
  turba_nosuchobject = True
  c$ = "select object_alias as wert from turba_objects where owner_id='" + adrbkid$ + "' and object_alias like '" + id$ + "|%'"
  If get1hordeerg(c$) <> "" Then
    turba_nosuchobject = False
  Else
  Debug.Print c$
  End If
End Function

