VERSION 5.00
Begin VB.Form splan 
   Caption         =   "Saalplan"
   ClientHeight    =   7695
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command20 
      Caption         =   "go"
      Height          =   255
      Left            =   11520
      TabIndex        =   99
      Top             =   3960
      Width           =   375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   11880
      TabIndex        =   98
      Text            =   "Combo3"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "splan.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   97
      ToolTipText     =   "per Email an Agencyprof"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Termin ins Abo"
      Height          =   255
      Left            =   9720
      TabIndex        =   96
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ListBox abotermine 
      Height          =   2985
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   80
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command21 
      Caption         =   "neu zeichnen"
      Height          =   255
      Left            =   5880
      TabIndex        =   95
      Top             =   7440
      Width           =   1335
   End
   Begin VB.ComboBox rid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   93
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox rid1 
      Height          =   315
      Left            =   10320
      TabIndex        =   92
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox mwstvk 
      Height          =   285
      Left            =   11280
      TabIndex        =   90
      Text            =   "Text8"
      Top             =   7200
      Width           =   495
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   12360
      Picture         =   "splan.frx":00B2
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Saalplan löschen"
      Top             =   7200
      Width           =   495
   End
   Begin VB.ListBox aboliste 
      Height          =   1425
      Left            =   9600
      Sorted          =   -1  'True
      TabIndex        =   79
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox possabo 
      Height          =   840
      Left            =   1800
      TabIndex        =   88
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4800
      TabIndex        =   87
      Text            =   "Text7"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "zeige kundenplätze"
      Height          =   255
      Left            =   3120
      TabIndex        =   86
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text4 
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
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   85
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "?"
      Height          =   255
      Left            =   2160
      TabIndex        =   83
      ToolTipText     =   "Zeige Kundendaten"
      Top             =   5640
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Kunde"
      Height          =   255
      Left            =   0
      TabIndex        =   82
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   720
      TabIndex        =   81
      ToolTipText     =   "Kunde"
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abos bearbeiten"
      Height          =   255
      Left            =   9720
      TabIndex        =   78
      Top             =   7200
      Width           =   1455
   End
   Begin VB.ListBox termlist 
      Height          =   840
      Left            =   120
      TabIndex        =   77
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9120
      Picture         =   "splan.frx":1388
      Style           =   1  'Grafisch
      TabIndex        =   76
      ToolTipText     =   "bearbeiten"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      Picture         =   "splan.frx":24DA
      Style           =   1  'Grafisch
      TabIndex        =   75
      ToolTipText     =   "Alle Platzdaten dieser Aufführung löschen"
      Top             =   7200
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   74
      ToolTipText     =   "Zum Löschen aller Platzdaten dieses Termins deaktivieren"
      Top             =   7440
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Auswahl aufheben"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   73
      Top             =   6840
      Width           =   2895
   End
   Begin VB.ListBox knownwbegs 
      Height          =   840
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   72
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command12 
      Caption         =   "selektierte löschen"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11520
      TabIndex        =   71
      Top             =   5340
      Width           =   1335
   End
   Begin VB.ListBox pgshowlist 
      Height          =   2205
      Left            =   0
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   70
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox neuerpreis 
      Height          =   285
      Left            =   12360
      TabIndex        =   68
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Preis"
      Enabled         =   0   'False
      Height          =   255
      Left            =   11520
      TabIndex        =   67
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Preisgrp setzen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   11520
      TabIndex        =   66
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   315
      Index           =   1
      Left            =   12120
      TabIndex        =   65
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox preislimit 
      Height          =   285
      Left            =   12480
      TabIndex        =   64
      Text            =   "0"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "<"
      Height          =   315
      Index           =   0
      Left            =   11880
      TabIndex        =   63
      Top             =   4320
      Width           =   255
   End
   Begin VB.ListBox selerg 
      Height          =   1620
      Left            =   9720
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   61
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ComboBox beglist 
      Height          =   315
      ItemData        =   "splan.frx":29CA
      Left            =   1800
      List            =   "splan.frx":29CC
      TabIndex        =   60
      Text            =   "20:00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text5 
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
      Left            =   120
      TabIndex        =   58
      Top             =   4320
      Width           =   1575
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
      ItemData        =   "splan.frx":29CE
      Left            =   11520
      List            =   "splan.frx":29D8
      TabIndex        =   57
      Text            =   "Combo2"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "select"
      Height          =   255
      Left            =   9720
      TabIndex        =   56
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox selstr_rowlist 
      Height          =   285
      Left            =   11280
      MultiLine       =   -1  'True
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox selstr_pglist 
      Height          =   285
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   54
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox p0 
      Height          =   960
      Left            =   7320
      ScaleHeight     =   900
      ScaleWidth      =   540
      TabIndex        =   53
      Top             =   6840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.ListBox rowlist 
      Height          =   1230
      Left            =   11520
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   52
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox pglist 
      Height          =   1230
      Left            =   9720
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   51
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   47
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   46
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   45
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox Check3 
      Caption         =   "ReihenNummer"
      Height          =   255
      Left            =   4440
      TabIndex        =   44
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "PlatzNummer"
      Height          =   255
      Left            =   3000
      TabIndex        =   43
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Halle"
      Height          =   255
      Left            =   3000
      TabIndex        =   42
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "kopiere nach"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9720
      TabIndex        =   39
      Top             =   5760
      Width           =   3135
   End
   Begin VB.ComboBox hid1 
      Height          =   315
      Left            =   10320
      TabIndex        =   36
      Top             =   6120
      Width           =   2535
   End
   Begin VB.ComboBox pgid1 
      Height          =   315
      Left            =   10320
      TabIndex        =   35
      Top             =   6840
      Width           =   2535
   End
   Begin VB.ListBox platzliste 
      Height          =   2205
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10920
      TabIndex        =   32
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   3
      Left            =   10920
      TabIndex        =   31
      Text            =   "1"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CheckBox rcntmode 
      Height          =   255
      Left            =   11400
      TabIndex        =   28
      Top             =   1200
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.TextBox nrows 
      Height          =   285
      Left            =   10560
      TabIndex        =   27
      Text            =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   2
      Left            =   10920
      TabIndex        =   25
      Text            =   "1"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CheckBox cntmode 
      Height          =   255
      Left            =   11400
      TabIndex        =   24
      Top             =   1560
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   0
      Left            =   11520
      TabIndex        =   18
      Text            =   "110"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox bbreit 
      Height          =   285
      Index           =   1
      Left            =   12120
      TabIndex        =   17
      Text            =   "120"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   11040
      TabIndex        =   16
      Text            =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   15
      Top             =   840
      Width           =   255
   End
   Begin VB.ComboBox pgid 
      Height          =   315
      Left            =   7920
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin VB.ComboBox hid 
      Height          =   315
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<=="
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "neu zeichnen"
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   11760
      TabIndex        =   6
      Text            =   "20"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11760
      TabIndex        =   5
      Text            =   "40"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox p1 
      Height          =   7080
      Left            =   3000
      ScaleHeight     =   7020
      ScaleWidth      =   6540
      TabIndex        =   4
      Top             =   360
      Width           =   6600
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   12240
      MaskColor       =   &H00000000&
      Picture         =   "splan.frx":29E5
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Speichern"
      Top             =   0
      Width           =   615
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
      Height          =   495
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   7200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Picture         =   "splan.frx":2D8C
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Raum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   94
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Raum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   91
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "% MwSt"
      Height          =   255
      Left            =   11760
      TabIndex        =   89
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginn"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   84
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   1800
      Picture         =   "splan.frx":2FDC
      ToolTipText     =   "Kontakt löschen verboten"
      Top             =   7200
      Width           =   315
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Preisgruppe"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   69
      Top             =   360
      Width           =   975
   End
   Begin VB.Label l5 
      Caption         =   "Preis"
      Height          =   255
      Index           =   9
      Left            =   11520
      TabIndex        =   62
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   59
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Preis"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   50
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Preisgrp"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   49
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   48
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Reihe/Platz"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   41
      Top             =   360
      Width           =   975
   End
   Begin VB.Label npl 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9720
      X2              =   12840
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Halle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   38
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Block:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   37
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label l5 
      Caption         =   "Preisgruppe"
      Height          =   255
      Index           =   8
      Left            =   9960
      TabIndex        =   33
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   9720
      X2              =   12720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label l5 
      Caption         =   "untere Reihen-Nr"
      Height          =   255
      Index           =   7
      Left            =   9600
      TabIndex        =   30
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label l5 
      Caption         =   "aufsteigend"
      Height          =   255
      Index           =   6
      Left            =   11640
      TabIndex        =   29
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label l5 
      Caption         =   "aufsteigend"
      Height          =   255
      Index           =   4
      Left            =   11640
      TabIndex        =   26
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label l5 
      Caption         =   "Reihen:"
      Height          =   255
      Index           =   5
      Left            =   9960
      TabIndex        =   23
      Top             =   840
      Width           =   615
   End
   Begin VB.Label l5 
      Alignment       =   1  'Rechts
      Caption         =   "linke Platz-Nr"
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label l5 
      Caption         =   "Tiefe"
      Height          =   255
      Index           =   2
      Left            =   12120
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.Label l5 
      Caption         =   "Breite"
      Height          =   255
      Index           =   1
      Left            =   11520
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.Label l5 
      Caption         =   "Plätze"
      Height          =   255
      Index           =   0
      Left            =   11040
      TabIndex        =   19
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Block:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiefe:"
      Height          =   255
      Left            =   11280
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Breite:"
      Height          =   255
      Left            =   11280
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Menu menu_bearb 
      Caption         =   "Platz bearbeiten"
      Visible         =   0   'False
      Begin VB.Menu menu_unselect 
         Caption         =   "Auswahl aufheben"
      End
      Begin VB.Menu menu_select_row 
         Caption         =   "Reihe markieren"
      End
      Begin VB.Menu ruler4 
         Caption         =   "----------------"
      End
      Begin VB.Menu menu_vk 
         Caption         =   "Verkauf"
      End
      Begin VB.Menu menu_retour 
         Caption         =   "Rücknahme"
      End
      Begin VB.Menu menu_kbuch 
         Caption         =   "Kassenbuch"
      End
      Begin VB.Menu ruler1 
         Caption         =   "----------------"
      End
      Begin VB.Menu menu_2abo 
         Caption         =   "Gewähltem Abo zuweisen"
      End
      Begin VB.Menu menu_abol_del 
         Caption         =   "Abo aufheben"
      End
      Begin VB.Menu menu_ehre 
         Caption         =   "Ehrenplatz"
      End
      Begin VB.Menu menu_bestell 
         Caption         =   "Vorbestellung"
      End
      Begin VB.Menu menu_bestell_del 
         Caption         =   "Vorbestellungen löschen"
      End
      Begin VB.Menu ruler2 
         Caption         =   "----------------"
      End
      Begin VB.Menu menu_bestell_kunde 
         Caption         =   "Kundendaten"
      End
      Begin VB.Menu menu_platz_kunde 
         Caption         =   "Zeige Plätze des Kunden"
      End
   End
End
Attribute VB_Name = "splan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nupd%
Dim posv As Double, posb As Double
Dim co%, px%, py%, chm%, save_in_progress%, m_pressx As Single, m_pressy As Single
Dim menu_bestell_override_vorgang$
Public noupd%, termlist_upd%, termlist_drw%

Private Function PlatzselektionVorhanden() As Integer
Dim i%

'd2infile = "splan": d2insub = "PlatzselektionVorhanden"
PlatzselektionVorhanden = 0
For i% = 0 To selerg.ListCount - 1
    If selerg.Selected(i%) = True Then
      PlatzselektionVorhanden = 1
      Exit Function
    End If
Next i%

End Function
Sub drw(o%, X, Y, col As Long, obb, obt, dsc$)
Dim i%, xo%, yoff%

'd2infile = "splan": d2insub = "drw"
Select Case o%
  Case 0: xo% = Val(nrows.text)
          For i% = 1 To xo%
            yoff% = ((i% - 1) * obt)
            P1.Line (X - obb / 2, Y - obt / 2 - yoff%)-(X + obb / 2, Y + obt / 2 - yoff%), col, B
            If dsc$ <> "" Then
              P1.Line (X + obb / 2, Y + obt / 2 - yoff%)-(X - obb / 2, Y - obt / 2 - yoff%), RGB(255, 255, 255)
              If col = 0 Then
                P1.Print dsc$
              Else
                P1.Print Space$(Len(dsc$))
              End If
            End If
          Next i%
  Case Else:
End Select

End Sub
Sub drwlegend()
Dim X As Integer, Y As Integer, dh As Integer

'd2infile = "splan": d2insub = "drwlegend"
P1.Line (0, 0)-(P1.ScaleWidth, 0)
P1.Line (10, 10)-(30, 0)
P1.Line (10, 10)-(30, 20)
P1.Line (P1.ScaleWidth - 10, 10)-(P1.ScaleWidth - 30, 0)
P1.Line (P1.ScaleWidth - 10, 10)-(P1.ScaleWidth - 30, 20)
P1.Line (P1.ScaleWidth / 2, 15)-(P1.ScaleWidth / 2, 15)
P1.Print trm(Int(P1.ScaleWidth) / 100) & "m"

P1.Line (10, 10)-(10, P1.ScaleHeight - 10)
P1.Line (0, 0)-(0, P1.ScaleHeight)
P1.Line (10, 10)-(20, 30)
P1.Line (10, 10)-(0, 30)
P1.Line (10, P1.ScaleHeight - 10)-(20, P1.ScaleHeight - 30)
P1.Line (10, P1.ScaleHeight - 10)-(0, P1.ScaleHeight - 30)
P1.Line (15, P1.ScaleHeight / 2)-(15, P1.ScaleHeight / 2)
P1.Print trm(Int(P1.ScaleHeight) / 100) & "m"

For X = 0 To P1.ScaleWidth Step 500
  dh = P1.ScaleWidth / 80: If (X Mod 100) = 0 Then dh = dh * 2
  P1.Line (X, P1.ScaleHeight)-(X, P1.ScaleHeight - dh)
  If dh <> P1.ScaleWidth / 100 Then
    P1.Line (X - 50, P1.ScaleHeight - dh * 2)-(X - 50, P1.ScaleHeight - dh * 2)
    P1.Print X / 100
  End If
Next X
For Y = 0 To P1.ScaleHeight Step 500
  dh = P1.ScaleHeight / 200: If (X Mod 100) = 0 Then dh = dh * 2
  P1.Line (P1.ScaleWidth, Y)-(P1.ScaleWidth - dh, Y)
  If dh <> P1.ScaleWidth / 200 Then
    P1.Line (P1.ScaleWidth - 2 * dh, Y)-(P1.ScaleWidth - 2 * dh, Y)
    P1.Print Y / 100
  End If
Next Y
End Sub

Private Sub aboliste_Click()
Dim rrr
Dim i%, id$, j%, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "aboliste_Click"
abotermine.Clear
i% = aboliste.ListIndex
If i% < 0 Then Exit Sub
id$ = aboliste.List(i%)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
id$ = "select * from hbabotermine where aboid='" & id$ & "' order by dtg"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, id$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  id$ = r!dtg & " " & r!adrid & " | " & r!pid & Space$(80) & "(ID:" & r!id
  abotermine.AddItem id$
  r.MoveNext
Wend


End Sub


Private Sub bbreit_Change(Index As Integer)
Dim rrr

'd2infile = "splan": d2insub = "bbreit_Change"
If Val(bbreit(0).text) > 0 Then
  rrr = Int(P1.ScaleWidth / Val(bbreit(0).text)) - 1
  Text3.text = trm(rrr)
End If

End Sub


Public Sub beglist_Change()
Dim rrr
Dim h$, p$, r As ADODB.Recordset, dtg$, c$
Dim col As Long, offy%, dy As Single, p1x As Single, p1y As Single

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "beglist_Change"
If noupd% = 1 Then Exit Sub

splan.Caption = "Saalplan"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ <> "" And trm(p$) <> "" Then
  dtg$ = trm(Text5.text)
  If dtg$ <> "" Then
    dtg$ = trm(dtg$ & " " & beglist.text)
    'alle platzdaten für den eingetragenen tag/beginn finden
    c$ = "SELECT hbplist.hid as hid, hbplist.pgid, hbplist.px, hbplist.py, hbplist.obb, hbplist.obt, hbplist.pid as hbpid, hbplist.platz as pnr, hbplist.reihe as rnr, hbplist.preisgruppe as pgruppe, hbplist.preis , hbpstatus.pstatus as hbpstat " + _
          "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
          "WHERE ((((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "')) and (hbpstatus.dtg='" & dtg$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    While Not r.EOF
      dy = r!obt / 4
      col = RGB(255, 0, 0)
      Select Case LCase(r!hbpstat)
        Case "bestellung": col = RGB(0, 255, 255)
                           offy% = 1
        Case "verkauft": col = RGB(0, 0, 128)
                           offy% = 1
        Case "ehrenplatz": col = RGB(255, 0, 0)
                           offy% = 0
        Case "abo": col = RGB(0, 255, 0)
                           offy% = 2
        Case "abo verkauft": col = RGB(0, 0, 0)
                           offy% = 2
        Case Else:  col = RGB(128, 128, 128)
                    offy% = 3
      End Select
      p1x = r!px - r!obb / 2
      p1y = r!py - r!obt / 2
      P1.Line (p1x + r!obb, p1y + (offy% + 1) * dy)-(p1x, p1y + offy% * dy), col, BF
      DoEvents
      r.MoveNext
    Wend
    Call abochk(h$, p$, dtg$)
  End If
End If

End Sub

Private Sub beglist_Click()

'd2infile = "splan": d2insub = "beglist_Click"
Call Command2_Click
Call beglist_Change
End Sub

Private Sub Check1_Click(Index As Integer)
Dim aw%, i%

'd2infile = "splan": d2insub = "Check1_Click"
If Index < 0 Then Exit Sub
If nupd% = 0 Then
  nupd% = 1
  aw% = Check1(Index).value
  co% = -1: If aw% = 1 Then co% = Index
  For i% = 0 To chm%
    If i% <> Index Then Check1(i%) = 0
  Next i%
End If
nupd% = 0

End Sub

Private Sub Check2_Click()
'd2infile = "splan": d2insub = "Check2_Click"
Call form1.setmylastFormVar(Me.name, "shw_pn", trm(Check2.value))
Call Command2_Click
Call beglist_Change
End Sub

Private Sub Check3_Click()
'd2infile = "splan": d2insub = "Check3_Click"
Call form1.setmylastFormVar(Me.name, "shw_rn", trm(Check3.value))
Call Command2_Click
Call beglist_Change
End Sub

Private Sub Check4_Click()

'd2infile = "splan": d2insub = "Check4_Click"
If Check4.value = 1 Then
  Command14.Enabled = False
Else
  Command14.Enabled = True
End If

End Sub

Private Sub Combo1_Click()
'd2infile = "splan": d2insub = "Combo1_Click"
If selerg.ListCount > 0 And Combo1.text <> "" Then Command9.Enabled = True
End Sub

Private Sub Combo1_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Combo1_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT id,preis,waehrung FROM preisgruppen"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
Combo1.Clear
While Not r.EOF
  Combo1.AddItem r!id & " - (" & trm(fixeur(r!preis)) & " " & r!waehrung & ")"
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Sub Command1_Click()
'd2infile = "splan": d2insub = "Command1_Click"
Call savecheck
Unload kvk
Unload Me

End Sub


Private Sub Command10_Click(Index As Integer)
Dim rrr
Dim c$, h$, p$, r As ADODB.Recordset, i%, verg$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command10_Click"
verg$ = ">"
If Index = 0 Then verg$ = "<"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
For i% = 0 To rowlist.ListCount - 1
  rowlist.Selected(i%) = False
Next i%
For i% = 0 To pglist.ListCount - 1
  pglist.Selected(i%) = False
Next i%
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
If exist(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif") = 0 Then Call Command2_Click
P1.Picture = LoadPicture(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif")
c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and preis" & verg$ & d2db(preislimit.text)
selerg.Clear
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  selerg.AddItem r!platzname & Space$(80) & "(ID:" & r!id & Space$(8) & "(px:" & r!px & Space$(8) & "(py:" & r!py & Space$(8) & "(obb:" & r!obb & Space$(8) & "(obt:" & r!obt
  r.MoveNext
Wend
For i% = 0 To selerg.ListCount - 1
  selerg.Selected(i%) = True
Next i%

End Sub

Private Sub Command11_Click()
Dim rrr
Dim i%, c$, r As ADODB.Recordset, l$, id$, j%, waehr$, preis As Double, pg$, h$, p$, pi%, dtg$, dtg0$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command11_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If

On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0
dtg$ = trm(Text5.text)
If dtg$ <> "" Then
  dtg0$ = trm(dtg$ & " " & beglist.text)
  dtg$ = " and dtg = '" & dtg0$ & "'"
End If
MousePointer = 11
DoEvents
co% = -2
If selerg.ListCount > 0 Then
  knownwbegs.Clear
  For i% = 0 To selerg.ListCount - 1
    If selerg.Selected(i%) = True Then
      l$ = selerg.List(i%)
      j% = InStr(l$, "(ID:"): id$ = word1(Mid$(l$, j% + 4))
      c$ = "select * from hbpstatus where hbpid='" & id$ & "'" & dtg$
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        c$ = "update hbpstatus set preis=" & d2db(neuerpreis.text) & " where hbpid='" & id$ & "'" & dtg$
        Call form1.sqlqry(c$)
        If r!pstatus = "Abo" Then
          l$ = "update hbpstatus set preis=" & d2db(neuerpreis.text) & " where pstatus2='" & r!pstatus2 & "'"
          For pi% = 0 To knownwbegs.ListCount - 1
            If knownwbegs.List(pi%) = l$ Then pi% = knownwbegs.ListCount + 10
          Next pi%
          If pi% < knownwbegs.ListCount + 5 Then knownwbegs.AddItem l$
        End If
      Else
        c$ = "insert into hbpstatus (id,hbpid,preis,dtg) values('" & form1.newid("hbpstatus", "id", 20) & "','" & id$ & "','" & d2db(neuerpreis.text) & "','" & dtg0$ & "')"
        Call form1.sqlqry(c$)
      End If
    End If
  Next i%
  For i% = 0 To knownwbegs.ListCount - 1
    Call form1.sqlqry(knownwbegs.List(i%))
  Next i%
  knownwbegs.Clear
Else
  Command11.Enabled = False
End If
MousePointer = 0
co% = -1
End Sub

Private Sub Command12_Click()
Dim rrr
Dim i%, c$, r As ADODB.Recordset, l$, id$, j%, waehr$, preis As Double, pg$, h$, p$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command12_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If

On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0

MousePointer = 11
DoEvents

If selerg.ListCount > 0 Then
  For i% = 0 To selerg.ListCount - 1
    If selerg.Selected(i%) = True Then
      l$ = selerg.List(i%)
      j% = InStr(l$, "(ID:"): id$ = word1(Mid$(l$, j% + 4))
      c$ = "select id from hbpstatus where hbpid='" & id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If r.EOF Then
        c$ = "delete from hbplist where id='" & id$ & "'": Call form1.sqlqry(c$)
        c$ = "delete from hbpstatus where hbpid='" & id$ & "'": Call form1.sqlqry(c$)
      Else
        j% = MsgBox("Ex existieren Platzdaten für " & trm(Left$(selerg.List(i%), 20)) & vbCrLf & "Dieser Platz wird nicht gelöscht.", vbOKCancel)
        If j% = 2 Then i% = selerg.ListCount + 10
      End If
    End If
  Next i%
Else
  Command11.Enabled = False
End If
MousePointer = 0
Call Command2_Click

End Sub

Private Sub Command13_Click()

'd2infile = "splan": d2insub = "Command13_Click"
  selerg.Clear
  Call Command2_Click
  Call beglist_Change
  Text6.text = ""
End Sub

Private Sub Command14_Click()
Dim rrr
Dim h$, p$, r As ADODB.Recordset, dtg$, c$
Dim col As Long, offy%, dy As Single, p1x As Single, p1y As Single

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command14_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ <> "" And trm(p$) <> "" Then
  dtg$ = trm(Text5.text)
  If dtg$ <> "" Then
    dtg$ = trm(dtg$ & " " & beglist.text)
    MousePointer = 11: DoEvents
    'alle platzdaten für den eingetragenen tag/beginn löschen
    c$ = "SELECT hbplist.hid as hid, hbplist.pgid, hbplist.px, hbplist.py, hbplist.obb, hbplist.obt, hbplist.pid as hbpid, hbplist.platz as pnr, hbplist.reihe as rnr, hbplist.preisgruppe as pgruppe, hbplist.preis , hbpstatus.id as hbpsid " + _
          "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
          "WHERE ((((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "')) and (hbpstatus.dtg='" & dtg$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    While Not r.EOF
      c$ = "delete from hbpstatus where id='" & r!hbpsid & "'": Call form1.sqlqry(c$)
      r.MoveNext
    Wend
    MousePointer = 0
    Call Command2_Click
    Call beglist_Change
  End If
End If
Check4.value = 1
End Sub

Private Sub Command15_Click()
'd2infile = "splan": d2insub = "Command15_Click"
Me.Width = delme.Left + delme.Width + 200
Command15.Enabled = False
rid.Enabled = True

End Sub

Private Sub Command16_Click()
Dim s0$, neuwert$, ks$, s$

'd2infile = "splan": d2insub = "Command16_Click"
Load adrselect
s0$ = Text6.text
If Len(s0$) = 0 Then s0$ = ""
Call adrselect.sel_init(s0$, s$)
Call adrselect.SetFocus
Do
  DoEvents
Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
If adrselect.sel_brk() = 0 Then
  neuwert$ = adrselect.sel_getselected()
  ks$ = adrselect.get_kontsel()
  If ks$ <> "" Then neuwert$ = neuwert$ & "|" & ks$
  Unload adrselect
  Text6.text = neuwert$
End If
End Sub

Private Sub Command17_Click()
Dim rrr
Dim s$, p%, c$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command17_Click"
s$ = trm(Text6.text)
p% = InStr(s$, "|")
If p% > 0 Then s$ = Left$(s$, p% - 1)
If s$ <> "" Then
    c$ = "select id from adresse where id='" & s$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    If r.EOF Then
      p% = MsgBox("Der Kunde existiert nicht, anlegen?", vbYesNo)
      If p% = vbYes Then
        c$ = "insert into adresse (id,name) values('" & s$ & "','" & s$ & "')"
        form1.sqlqry (c$)
        form1.sqlqry ("insert into adresstyp (id,vid,typ,wert,kid) values('" + _
                   form1.newid("adresstyp", "id", 20) & "','" + _
                   s$ & "','Kunde',NULL,'-1')")
      End If
    End If
    Load shwAdrDetail
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(s$, "")
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
Else
  Call Command16_Click
End If

End Sub

Private Sub Command18a(k$)
Dim rrr
Dim c$, r As ADODB.Recordset, h$, p$, dtg$, l$, i%

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command18a"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Or p$ = "" Then Exit Sub
dtg$ = trm(Text5.text)
If dtg$ = "" Then Exit Sub
l$ = trm(beglist.text)
If l$ = "" Then Exit Sub
MousePointer = 11: DoEvents
dtg$ = trm(dtg$ & " " & l$)
dtg$ = Left$(dtg$, 20)
c$ = "SELECT hbplist.hid, hbplist.pgid, hbplist.id as hbplid, hbpstatus.dtg, hbpstatus.adrid, hbplist.platzname, " + _
     "hbplist.px, hbplist.py, hbplist.obt, hbplist.obb " + _
     "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
     "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') AND ((hbpstatus.dtg)='" & dtg$ & "') AND ((hbpstatus.adrid)='" & k$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  l$ = r!platzname & Space$(80) & "(ID:" & r!hbplid & Space$(8) & "(px:" & r!px & Space$(8) & "(py:" & r!py & Space$(8) & "(obb:" & r!obb & Space$(8) & "(obt:" & r!obt
  For i% = 0 To selerg.ListCount - 1
    If l$ = selerg.List(i%) Then i% = selerg.ListCount + 10
  Next i%
  If i% < selerg.ListCount + 5 Then
    selerg.AddItem l$
    For i% = 0 To selerg.ListCount - 1
      If selerg.Selected(i) = False Then
        selerg.Selected(i) = True
        DoEvents
      End If
    Next i%
  End If
  r.MoveNext
Wend
MousePointer = 0

End Sub
Private Sub Command18_Click()
Dim k$

'd2infile = "splan": d2insub = "Command18_Click"
k$ = trm(Text7.text)
If k$ <> "" Then Call Command18a(k$)


End Sub

Private Sub Command19_Click()
'd2infile = "splan": d2insub = "Command19_Click"
Call form1.handbuchcall("index.html")

End Sub

Public Sub Command2_Click()
Dim rrr
Dim h$, p$, r As ADODB.Recordset, c$, o%, l$, co%, obb As Single, obt As Single, obl$, p1x As Single, p1y As Single
Dim reihe, platz, i%, gx%, pg$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command2_Click"
selerg.Clear
pgshowlist.Clear
P1.Cls
P1.Picture = p0.Picture
DoEvents
platzliste.Clear

DoEvents
If Command13.Caption = "Auswahl aufheben" Then Command13.Enabled = False

rowlist.Clear
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
c$ = "select * from hblist where hid='" & h$ & "' and pgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Text1.text = r!breite
  Text2.text = r!tiefe
  rid.text = trm(r!raum)
  DoEvents
End If
c$ = form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif"
gx% = 0: If exist(c$) <> 0 Then gx% = 1
c$ = form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
If exist(c$) <> 0 Then
  o% = FreeFile
  Open c$ For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    platzliste.AddItem l$
    co% = InStr(l$, "|pg=")
    pg$ = Mid$(l$, co% + 4)
    pg$ = trm(Left$(pg$, InStr(pg$, "|") - 1))
    For i% = 0 To pgshowlist.ListCount - 1
      If pgshowlist.List(i%) = pg$ Then
        i% = pgshowlist.ListCount + 10
      End If
    Next i%
    If i% < pgshowlist.ListCount + 5 Then pgshowlist.AddItem pg$
    If gx% = 0 Then
      reihe = Val(l$)
      obl$ = Mid$(l$, InStr(l$, "/") + 1)
      platz = Val(Left$(obl$, InStr(l$, " ")))
      co% = InStr(l$, Space$(10))
      l$ = trm(Mid$(l$, co%))
      co% = InStr(l$, "|ob")
      obl$ = Mid$(l$, co% + 5)
      l$ = trm(Left$(l$, co% - 1))
      obb = Val(obl$)
      co% = InStr(obl$, "|obt=")
      obt = Val(Mid$(obl$, co% + 5))
      p1x = Val(l$)
      co% = InStr(l$, "/")
      p1y = Val(trm(Mid$(l$, co% + 1)))
      P1.Line (p1x + obb / 2, p1y + obt / 2)-(p1x - obb / 2, p1y - obt / 2), RGB(0, 0, 0), B
      If Check2.value = 1 Or Check3.value = 1 Then
        P1.Line (p1x + obb / 2, p1y + obt / 2)-(p1x - obb / 2, p1y - obt / 2), RGB(255, 255, 255)
        If Check3.value = 1 Then
          P1.Print reihe;
          If Check2.value = 1 Then P1.Print "/";
        End If
        If Check2.value = 1 Then P1.Print platz
      End If
    End If
    DoEvents
  Wend
  Close #o%
  If gx% = 0 Then
    SavePicture P1.Image, form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif"
  Else
    P1.Picture = LoadPicture(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif")
  End If
Else
  On Error Resume Next
  MkDir form1.s0dir() + "\" + form1.medien() + "\"
  MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(h$)
  MkDir form1.s0dir() + "\" + form1.medien() + "\" + form1.medienname(h$) & "\saalplan\"
  On Error GoTo 0
  o% = FreeFile
  Open c$ For Output As #o%
  c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    P1.Line (r!px + r!obb / 2, r!py + r!obt / 2)-(r!px - r!obb / 2, r!py - r!obt / 2), RGB(0, 0, 0), B
    If Check2.value = 1 Or Check3.value = 1 Then
      P1.Line (r!px + r!obb / 2, r!py + r!obt / 2)-(r!px - r!obb / 2, r!py - r!obt / 2), RGB(255, 255, 255)
      If Check3.value = 1 Then
        P1.Print r!reihe;
        If Check2.value = 1 Then P1.Print "/";
      End If
      If Check2.value = 1 Then P1.Print r!platz
    End If
    l$ = Format$(r!reihe, "0###") & "/" & Format$(r!platz, "0###") & Space$(20) & r!px & "/" & r!py & "|obb=" & r!obb & "|obt=" & r!obt & "|pg=" & r!preisgruppe & " |"
    platzliste.AddItem l$
    For i% = 0 To pgshowlist.ListCount - 1
      If pgshowlist.List(i%) = r!preisgruppe Then
        i% = pgshowlist.ListCount + 10
      End If
    Next i%
    If i% < pgshowlist.ListCount + 5 Then pgshowlist.AddItem r!preisgruppe
    Print #o%, l$
    DoEvents
    r.MoveNext
  Wend
  Close #o%
  SavePicture P1.Image, form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif"
End If
l$ = ""
For i% = 0 To platzliste.ListCount - 1
  If l$ <> Left$(platzliste.List(i%), 4) Then
    l$ = Left$(platzliste.List(i%), 4)
    rowlist.AddItem "Reihe " & Val(l$)
  End If
Next i%
Call drwlegend
BackColor = form1.cleancolor()
Call setminrow
npl.Caption = platzliste.ListCount & " Plätze"
If termlist_upd% = 1 Then
  l$ = ""
  termlist.Clear
  c$ = "SELECT hbpstatus.dtg " + _
   "FROM (hblist INNER JOIN hbplist ON (hblist.pgid = hbplist.pgid) AND (hblist.hid = hbplist.hid)) INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
   "Where (((hblist.hid) = '" & h$ & "') And ((hblist.pgid) = '" & p$ & "')) " + _
   "ORDER BY hbpstatus.dtg;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    If l$ <> r!dtg Then
      termlist.AddItem r!dtg
      l$ = r!dtg
    End If
    r.MoveNext
  Wend
End If

possabo.Clear: l$ = ""
c$ = "SELECT hbabos.Name as aname, hbabos.abosproraum as apr, hbabos.id as aid " + _
   "FROM hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid " + _
   "WHERE (((hbabotermine.adrid)='" & h$ & "') AND ((hbabotermine.pid)='" & p$ & "'))" + _
   "ORDER BY hbabos.Name;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  c$ = r!aname & " (" & r!apr & " Plätze)" & Space$(80) & "(ID:" & r!aid
  If l$ <> c$ Then
    possabo.AddItem c$
    l$ = c$
  End If
  r.MoveNext
Wend

End Sub


Private Sub Command20_Click()
Dim rrr
Dim i%, l$, r1%, r2%, c$, r As ADODB.Recordset, pnr0%, h$, p$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command20_Click"
If Combo3.text = "Reihen renum." Then
  For i% = 0 To rowlist.ListCount - 1
    h$ = "Reihe" & str$(i% + 1)
    If Left(rowlist.List(i%), Len(h$)) <> h$ Then
      h$ = trm(hid.text)
      p$ = trm(pgid.text)
      c$ = "update hbplist set reihe=" & trm(i% + 1) & " where hid='" & h$ & "' and pgid='" & p$ & "' and reihe=" & r1%
    End If
  Next i%
End If
If Combo3.text = "Reihen zus." Then

r1% = -1: r2% = -1
For i% = 0 To rowlist.ListCount - 1
  If rowlist.Selected(i%) = True Then
    If r1% < 0 Then r1% = i%
    r2% = i%
  End If
Next i%
If r1% >= 0 And r2% >= 0 And r1% <> r2% Then
  l$ = rowlist.List(r1%): r1% = Val(trm(Mid$(l$, InStr(l$, " "))))
  l$ = rowlist.List(r2%): r2% = Val(trm(Mid$(l$, InStr(l$, " "))))
  h$ = trm(hid.text)
  p$ = trm(pgid.text)
  c$ = "select max(platz) as pmin from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and reihe=" & r1%
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF Then
    pnr0% = r!pmin + 1
  End If
  c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and reihe=" & r2% & " order by platz"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    c$ = "Reihe" & " " & r1% & " Platz " & pnr0%
    c$ = "update hbplist set reihe=" & trm(r1%) & ",platz=" & trm(pnr0%) & ",platzname='" & c$ & "' where id='" & r!id & "'"
    Call form1.sqlqry(c$)
    pnr0% = pnr0% + 1
    r.MoveNext
  Wend
End If
Call Command2_Click

End If
End Sub

Private Sub Command21_Click()
Dim h$, p$, r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command21_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0
Call Command2_Click

End Sub

Private Sub Command22_Click()
Dim dtg$, c$, r As ADODB.Recordset
Dim i%, l$, j%
Dim h$, p$, id$, aid$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command22_Click"
i% = abos.aboliste.ListIndex
Command13.Enabled = False
If i% < 0 Then Exit Sub
aid$ = abos.aboliste.List(i%)
j% = InStr(aid$, "(ID:") + 4
aid$ = Mid$(aid$, j%)

h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Or trm(p$) = "" Then Exit Sub
dtg$ = trm(Text5.text)
If Len(dtg$) = 10 Then
  l$ = trm(beglist.text)
  If l$ <> "" Then
    dtg$ = trm(dtg$ & " " & l$)
    dtg$ = Left$(dtg$, 20)
    id$ = form1.newid("hbabotermine", "id", 25)
    c$ = "insert into hbabotermine (id,aboid,dtg,adrid,pid) values('" + _
         id$ & "','" + _
         aid$ & "','" + _
         dtg$ & "','" + _
         h$ & "','" + _
         p$ & "')"
    Call form1.sqlqry(c$)
    Call abos.aboliste_Click
  End If
End If

End Sub

Private Sub Command3_Click()
'd2infile = "splan": d2insub = "Command3_Click"
Load abos
    On Error Resume Next
    Call abos.SetFocus
    On Error GoTo 0

End Sub

Private Sub Command31_Click()
Dim h$, p$, tpid$, c$, r As ADODB.Recordset, tg$, rrr

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command31_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
c$ = "select * from hblist where hid='" & h$ & "' and pgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If r.EOF Then Exit Sub
tpid$ = r!id
MousePointer = 11: DoEvents
On Error Resume Next
Kill form1.mydatadir() & "\*.sql"
On Error GoTo 0
Call form1.sqlex_adresse("hblist", "id", tpid$)
Load smtp
On Error Resume Next
Call smtp.SetFocus
On Error GoTo 0
smtp.txtMessageSubject = "Agencyprof Datenpakete Saalplan " & h$ & " " & p$
smtp.txtMessageText = "Speichern Sie das Attachment in Ihrem Agencyprof-Verzeichnis"
tg$ = Dir(form1.mydatadir() & "\*.sql")
While tg$ <> ""
  Call smtp.attachfile(form1.mydatadir() & "\" & tg$)
  tg$ = Dir
Wend
MousePointer = 0

End Sub

Private Sub Command4_Click()
Dim rrr
Dim h$, p$, r As ADODB.Recordset, c$, raum$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command4_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
raum$ = trm(rid.text)
If raum$ = "" Then raum$ = p$
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0

co% = -1
c$ = "select * from hblist where hid='" & h$ & "' and pgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  c$ = "update hblist set breite=" & trm(Text1.text) & " where hid='" & h$ & "' and pgid='" & p$ & "'": Call form1.sqlqry(c$)
  c$ = "update hblist set tiefe=" & trm(Text2.text) & " where hid='" & h$ & "' and pgid='" & p$ & "'": Call form1.sqlqry(c$)
  c$ = "update hblist set raum='" & raum$ & "' where hid='" & h$ & "' and pgid='" & p$ & "'": Call form1.sqlqry(c$)
Else
  c$ = "insert into hblist (id,hid,raum,pgid,tiefe,breite) values('" + _
        form1.newid("hblist", "id", 20) & "','" & h$ & "','" & raum$ & "','" & p$ & "'," + _
        trm(Text2.text) & "," & trm(Text1.text) & ")"
  Call form1.sqlqry(c$)
End If
BackColor = form1.cleancolor()
Call Command2_Click

End Sub

Private Sub Command5_Click()
Dim i%

'd2infile = "splan": d2insub = "Command5_Click"
nupd% = 1
For i% = 0 To chm%
 Check1(i%) = 0
Next i%
Call savecheck
nupd% = 0
Call Command2_Click
Me.Width = P1.Left + P1.Width + 200
Command15.Enabled = True
Call beglist_Change
rid.Enabled = False
End Sub

Private Sub Command6_Click()
Dim rrr
Dim srch$, srcpg$, c$, p$
Dim j%, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command6_Click"
srcpg$ = trm(pgid.text)
srch$ = trm(hid.text)
If srch$ = "" Then Exit Sub
If pgid.text = pgid1.text And hid.text = hid1.text Then Exit Sub

If trm(srcpg$) = "" Then
  srcpg$ = "_leer"
  pgid1.text = srcpg$
End If

MousePointer = 11
DoEvents
pgid.text = trm(" " & pgid1.text)
p$ = pgid.text
If p$ = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
hid.text = hid1.text
Call Command4_Click
c$ = "delete from hblist where hid='" & hid1.text & "' and pgid='" & pgid1.text & "'": Call form1.sqlqry(c$)
c$ = "insert into hblist (id,hid,raum,pgid,tiefe,breite) values('" + _
        form1.newid("hblist", "id", 20) & "','" & hid1.text & "','" & rid1.text & "','" & pgid1.text & "'," + _
        trm(Text2.text) & "," & trm(Text1.text) & ")"
Call form1.sqlqry(c$)
c$ = "delete from hbplist where hid='" & hid1.text & "' and pgid='" & pgid1.text & "'": Call form1.sqlqry(c$)

c$ = "select * from hbplist where hid='" & srch$ & "' and pgid='" & srcpg$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  c$ = "insert into hbplist (id,hid,pgid,pid,platzname,preisgruppe,preis,reihe,platz,px,py,obb,obt,waehrung) values('" + _
                    form1.newid("hbplist", "id", 30) & "','" + _
                    hid1.text & "','" + _
                    p$ & "','" + _
                    r!pid & "','" + _
                    r!platzname & "','" + _
                    r!preisgruppe & "'," + _
                    d2db(r!preis) & "," + _
                    r!reihe & "," + _
                    r!platz & "," + _
                    r!px & "," + _
                    r!py & "," + _
                    r!obb & "," + _
                    r!obt & ",'" + _
                    r!waehrung & "')"
  Call form1.sqlqry(c$)
  r.MoveNext
Wend
On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(hid1.text) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(hid1.text) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0
MousePointer = 0
Call Command2_Click

End Sub

Private Sub Command7_Click()
Dim sid$

'd2infile = "splan": d2insub = "Command7_Click"
sid$ = hid.text
If Len(sid$) > 0 Then
    Load shwAdrDetail
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(sid$, "")
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
End If

End Sub

Private Sub Command8_Click()
Dim rrr
Dim c$, h$, p$, r As ADODB.Recordset, i%

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command8_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
If exist(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif") = 0 Then Call Command2_Click
'p1.Picture = LoadPicture(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif")
c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' "
If selstr_pglist.text <> "" Then
  c$ = c$ & " and (" & selstr_pglist
  If selstr_rowlist.text <> "" Then c$ = c$ & " " & Combo2.text & " " & selstr_rowlist
  c$ = c$ & ")"
Else
  If selstr_rowlist.text <> "" Then c$ = c$ & " and (" & selstr_rowlist & ")"
End If
'MsgBox c$
selerg.Clear
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  selerg.AddItem r!platzname & Space$(80) & "(ID:" & r!id & Space$(8) & "(px:" & r!px & Space$(8) & "(py:" & r!py & Space$(8) & "(obb:" & r!obb & Space$(8) & "(obt:" & r!obt
  r.MoveNext
Wend
For i% = 0 To selerg.ListCount - 1
  selerg.Selected(i%) = True
Next i%
Command9.Enabled = False
Command11.Enabled = False
If selerg.ListCount > 0 Then
  If neuerpreis.text <> "" Then Command11.Enabled = True
  If Combo1.text <> "" Then Command9.Enabled = True
End If
End Sub

Private Sub Command9_Click()
Dim rrr
Dim i%, c$, r As ADODB.Recordset, l$, id$, j%, waehr$, preis As Double, pg$, h$, p$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Command9_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If

On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0

pg$ = ""
i% = InStr(Combo1.text, " - (")
If i% = 0 Then
  Combo1.text = ""
  Command9.Enabled = False
  Exit Sub
End If
MousePointer = 11
DoEvents
If i% > 0 Then pg$ = Left$(Combo1.text, i% - 1)
preis = 0
waehr$ = ""
c$ = "SELECT id,preis,waehrung FROM preisgruppen where id='" & pg$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  waehr$ = r!waehrung
  preis = r!preis
End If
If selerg.ListCount > 0 And Combo1.text <> "" Then
  For i% = 0 To selerg.ListCount - 1
    If selerg.Selected(i%) = True Then
      l$ = selerg.List(i%)
      j% = InStr(l$, "(ID:"): id$ = word1(Mid$(l$, j% + 4))
      c$ = "update hbplist set preisgruppe='" & pg$ & "', preis=" & d2db(preis) & ", waehrung='" & waehr$ & "' where id='" & id$ & "'"
      Call form1.sqlqry(c$)
    End If
  Next i%
Else
  Command9.Enabled = False
End If
MousePointer = 0
End Sub

Private Sub delme_Click()
Dim rrr
Dim c$, h$, p$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "delme_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If
MousePointer = 11
DoEvents
On Error Resume Next
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
On Error GoTo 0
c$ = "SELECT hbpstatus.id, hbplist.hid, hbplist.pgid " + _
     "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
     "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  MsgBox "Es existieren noch Platzdaten. - Löschen nicht möglich"
Else
  c$ = "delete from hbplist where hid='" & h$ & "' and pgid='" & p$ & "'"
  Call form1.sqlqry(c$)
  c$ = "delete from hblist where hid='" & h$ & "' and pgid='" & p$ & "'"
  Call form1.sqlqry(c$)
  Call Command5_Click
End If
MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i%, rrr, klrv%, r As ADODB.Recordset, c$
Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "Form_Load"
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
Call form1.formpos(Me)
save_in_progress% = 0
menu_bestell_override_vorgang$ = ""
chm% = 0
nupd% = 0
px% = -1
co% = -1
noupd% = 0
termlist_upd% = 1
termlist_drw% = 1
P1.BackColor = RGB(255, 255, 255)
P1.AutoRedraw = True
nupd% = 0
BackColor = form1.cleancolor()
Check2.BackColor = form1.cleancolor()
Check3.BackColor = form1.cleancolor()
Text1.text = 40
Text2.text = 20
platzliste.Clear
P1.Cls
DoEvents

Call drwlegend
Me.Width = P1.Left + P1.Width + 200
Label3.Caption = ""
i% = 0
rrr = 0
While rrr = 0
  On Error Resume Next
  l5(i%).BackColor = form1.cleancolor()
  rrr = Err
  On Error GoTo 0
  i% = i% + 1
Wend
klrv% = Val(form1.mylastFormVar(Me.name, "shw_pn", "0"))
If klrv% <> 0 Then klrv% = 1
Check2.value = klrv%
klrv% = Val(form1.mylastFormVar(Me.name, "shw_rn", "0"))
If klrv% <> 0 Then klrv% = 1
Check3.value = klrv%
mwstvk.text = Val(form1.mylastFormVar(Me.name, "mwst_vk", "0"))
c$ = "SELECT id,preis,waehrung FROM preisgruppen"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
pglist.Clear
While Not r.EOF
  pglist.AddItem r!id & " - (" & trm(fixeur(r!preis)) & " " & r!waehrung & ")"
  r.MoveNext
Wend
pglist.AddItem "keine - (0,00 " + transe("€") + ")"
Combo2.text = "OR"
Combo3.text = ""
Combo3.AddItem "Reihen zus."
Combo3.AddItem "Reihen renum."
beglist.Clear
For i% = 0 To 23: beglist.AddItem Format$(i%, "0#") & ":00": Next i%
beglist.text = "20:00"
Show
BackColor = form1.cleancolor()
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "splan": d2insub = "Form_Unload"
Unload abos
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)

exuld:
On Error GoTo 0


End Sub

Private Sub hid_Change()
Dim rrr
Dim h$, r As ADODB.Recordset, c$
Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "hid_Change"
If noupd% = 1 Then Exit Sub

Call savecheck
h$ = trm(hid.text)
pgid.text = ""
c$ = "SELECT ID FROM adresse WHERE ID='" & h$ & "';"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then Command15.Enabled = True
Call Command2_Click
Call beglist_Change
End Sub

Private Sub hid_Click()
'd2infile = "splan": d2insub = "hid_Click"
Call hid_Change
End Sub

Private Sub hid_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "hid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT adresse.ID, adresstyp.typ FROM adresstyp INNER JOIN adresse ON adresstyp.vid = adresse.ID WHERE (((adresstyp.typ)='Halle')) OR (((adresstyp.typ)='Theater')) ORDER BY adresse.ID;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
hid.Clear
While Not r.EOF
  hid.AddItem r!id
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Sub hid1_Click()
'd2infile = "splan": d2insub = "hid1_Click"
Command6.Enabled = True
End Sub

Private Sub hid1_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "hid1_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT adresse.ID, adresstyp.typ FROM adresstyp INNER JOIN adresse ON adresstyp.vid = adresse.ID WHERE (((adresstyp.typ)='Halle')) OR (((adresstyp.typ)='Theater')) ORDER BY adresse.ID;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
hid1.Clear
While Not r.EOF
  hid1.AddItem r!id
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Function get_hbpstatusid(usage$, mess%) As String
Dim rrr
Dim dtg$, c$, r As ADODB.Recordset
Dim i%, rc$, l$
Dim h$, p$, id$, reihe, platz

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "get_hbpstatusid"
get_hbpstatusid = ""
If platzliste.ListIndex < 0 Then Exit Function
rc$ = ""
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ <> "" And trm(p$) <> "" Then
  dtg$ = trm(Text5.text)
  If dtg$ <> "" Then
    'platzdaten für den eingetragenen tag/beginn finden
    l$ = platzliste.List(platzliste.ListIndex)
    reihe = Val(l$)
    l$ = Mid$(l$, InStr(l$, "/") + 1)
    platz = Val(Left$(l$, InStr(l$, " ")))
    c$ = "select id from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and platz=" & platz & " and reihe=" & reihe
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    If Not r.EOF Then
      id$ = r!id
      l$ = trm(beglist.text)
      If l$ <> "" Then
        dtg$ = trm(dtg$ & " " & l$)
        dtg$ = Left$(dtg$, 20)
        c$ = "select * from hbpstatus  where hbpid='" & id$ & "' and dtg='" & dtg$ & "'"
        If usage$ = "Retour-VK" Then c$ = "select * from hbpstatus  where hbpid='" & id$ & "' and dtg='" & dtg$ & "' and pstatus='Verkauf'"
        If usage$ = "Retour-Bestellung" Then c$ = "select * from hbpstatus  where hbpid='" & id$ & "' and dtg='" & dtg$ & "' and ((pstatus='Bestellung') or (pstatus='Ehrenplatz'))"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          rc$ = r!id
        End If
      End If
    End If
  End If
End If
Select Case LCase(usage$)
  Case "bestellung":
                     If Len(dtg$) < 10 Then
                       If mess% = 1 Then MsgBox "Der Beginn ist nicht ausreichend ausgefüllt." & vbCrLf & usage$ & " ist nicht möglich"
                       rc$ = ""
                     Else
                       rc$ = form1.newid("hbpstatus", "id", 26)
                       c$ = "insert into hbpstatus (id,hbpid,dtg,pstatus,dat_creat,dat_upd) values('" + _
                           rc$ & "','" + _
                           id$ & "','" + _
                           dtg$ & "','" + _
                           usage$ & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "')"
                        Call form1.sqlqry(c$)
                     End If
  Case "ehrenplatz":
                     If Len(dtg$) < 10 Then
                       If mess% = 1 Then MsgBox "Der Beginn ist nicht ausreichend ausgefüllt." & vbCrLf & usage$ & " ist nicht möglich"
                       rc$ = ""
                     Else
                       rc$ = form1.newid("hbpstatus", "id", 26)
                       c$ = "insert into hbpstatus (id,preis,hbpid,dtg,pstatus,dat_creat,dat_upd) values('" + _
                           rc$ & "',0,'" + _
                           id$ & "','" + _
                           dtg$ & "','" + _
                           usage$ & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "')"
                        Call form1.sqlqry(c$)
                     End If
  Case "retour-vk":
  Case "retour-bestellung":
  Case Else:  rc$ = ""
End Select

get_hbpstatusid = rc$

End Function


Private Sub setaboplaetze(adresse$, saal$, platz%, reihe%, abo_id$, abo_name$)
Dim rrr
Dim c$, r As ADODB.Recordset, hbpid$, l$, abpid%, raum$, apreis As Double, ap As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "setaboplaetze"
c$ = "select raum from hblist where hid='" & adresse$ & "' and pgid='" & saal$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then raum$ = trm(r!raum)
If raum$ = "" Then raum$ = "_leer"

c$ = "select id from hbplist where hid='" & adresse$ & "' and pgid='" & saal$ & "' and platz=" & platz% & " and reihe=" & reihe%
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  hbpid$ = r!id
  abpid% = 1
  apreis = 0
  c$ = "select preis from hbabos where id='" & abo_id$ & "'"
Set ap = New ADODB.Recordset
ap.CursorLocation = adUseServer
rrr = form1.adoopen(ap, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not ap.EOF Then apreis = ap!preis
  c$ = "SELECT Max(hbpstatus.aboplatzid) AS maxid " + _
   "FROM hbpstatus INNER JOIN (hbplist INNER JOIN hblist ON (hbplist.pgid = hblist.pgid) AND (hbplist.hid = hblist.hid)) ON hbpstatus.hbpid = hbplist.id " + _
   "WHERE (((hbplist.hid)='" & adresse$ & "') AND ((hblist.raum)='" & raum$ & "') AND ((hbpstatus.pstatus2)='" & abo_id$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  If Not r.EOF And Not IsNull(r!maxid) Then abpid% = r!maxid + 1

  c$ = "select dtg from hbabotermine WHERE ((aboid='" & abo_id$ & "') AND (adrid='" & adresse$ & "') AND (pid='" & saal$ & "'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    l$ = form1.newid("hbpstatus", "id", 26)
    c$ = "insert into hbpstatus (id,hbpid,aboplatzid,adrid,dtg,pstatus,pstatus2,preis,dat_creat,dat_upd) values('" + _
                           l$ & "','" + _
                           hbpid$ & "'," + _
                           abpid% & ",'" + _
                           abo_name$ & "','" + _
                           r!dtg & "','" + _
                           "Abo" & "','" + _
                           abo_id$ & "'," + _
                           d2db(apreis) & ",'" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "')"
    Call form1.sqlqry(c$)
    r.MoveNext
  Wend
End If


End Sub




Private Sub menu_2abo_Click()
Dim rrr
Dim c$, h$, l$, p$, rnr%, pnr%, r As ADODB.Recordset, aboid$, i%, hbpid$, aboname$, j%, l1%, apreis As Double

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "menu_2abo_Click"
MousePointer = 11: DoEvents
i% = possabo.ListIndex
If i% < 0 Then GoTo exme
aboid$ = possabo.List(i%)
i% = InStr(aboid$, "(ID:")
aboname$ = trm(Left$(aboid$, i% - 1))
aboid$ = Mid$(aboid$, i% + 4)
i% = InStr(aboname$, " (")
aboname$ = trm(Left$(aboname$, i% - 1))
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then GoTo exme
If p$ = "" Then GoTo exme

If PlatzselektionVorhanden() = 0 Then
  If platzliste.ListIndex < 0 Then GoTo exme
  'platzdaten für den eingetragenen tag/beginn finden
  l$ = platzliste.List(platzliste.ListIndex)
  rnr% = Val(l$)
  l$ = Mid$(l$, InStr(l$, "/") + 1)
  pnr% = Val(Left$(l$, InStr(l$, " ")))

  'test ob der platz irgendwo an einem abotermin blockiert ist
  '"FROM ((hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid) INNER JOIN hbplist ON (hbabotermine.pid = hbplist.pgid) AND (hbabotermine.adrid = hbplist.hid)) INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid "
  c$ = "SELECT hbabos.Name, hbabos.id, hbabotermine.adrid, hbabotermine.pid, hbabotermine.dtg, hbpstatus.pstatus, hbpstatus.pstatus2, hbpstatus.dtg as statdtg, hbpstatus.adrid, hbplist.platzname, hbplist.reihe, hbplist.platz " + _
       "FROM ((hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid) INNER JOIN hbplist ON (hbabotermine.adrid = hbplist.hid) AND (hbabotermine.pid = hbplist.pgid)) INNER JOIN hbpstatus ON (hbabotermine.dtg = hbpstatus.dtg) AND (hbplist.id = hbpstatus.hbpid) " + _
       "WHERE (((hbabos.id)='" & aboid$ & "') AND ((hbabotermine.adrid)='" & h$ & "') " + _
         "AND ((hbabotermine.pid)='" & p$ & "') AND ((hbplist.reihe)=" & rnr% & ") " + _
         "AND ((hbplist.platz)=" & pnr% & "));"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    If trm(r!pstatus2) = aboid$ Or trm(r!pstatus2) = "" Then
      MsgBox "Dieser Platz geht nicht: " & r!statdtg
      GoTo exme
    End If
    r.MoveNext
  Wend
  Call setaboplaetze(h$, p$, pnr%, rnr%, aboid$, aboname$)
  Call beglist_Change
Else
  For i% = 0 To selerg.ListCount - 1
    If selerg.Selected(i%) = True Then
      l$ = selerg.List(i%)
      j% = InStr(l$, "Reihe "): rnr% = Val(Mid$(l$, j% + 6))
      j% = InStr(l$, "Platz "): pnr% = Val(Mid$(l$, j% + 6))
      l$ = Format$(rnr%, "0###") & "/" & Format$(pnr%, "0###") & "  "
      l1% = Len(l$)
      'test ob der platz irgendwo an einem abotermin blockiert ist
      c$ = "SELECT hbabos.Name, hbabos.id, hbabotermine.adrid, hbabotermine.pid, hbabotermine.dtg, hbpstatus.pstatus, hbpstatus.pstatus2, hbpstatus.dtg as statdtg, hbpstatus.adrid, hbplist.platzname, hbplist.reihe, hbplist.platz " + _
        "FROM ((hbabos INNER JOIN hbabotermine ON hbabos.id = hbabotermine.aboid) INNER JOIN hbplist ON (hbabotermine.adrid = hbplist.hid) AND (hbabotermine.pid = hbplist.pgid)) INNER JOIN hbpstatus ON (hbabotermine.dtg = hbpstatus.dtg) AND (hbplist.id = hbpstatus.hbpid) " + _
        "WHERE (((hbabos.id)='" & aboid$ & "') AND ((hbabotermine.adrid)='" & h$ & "') " + _
            "AND ((hbabotermine.pid)='" & p$ & "') AND ((hbplist.reihe)=" & rnr% & ") " + _
            "AND ((hbplist.platz)=" & pnr% & "));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      While Not r.EOF
        If trm(r!pstatus2) = aboid$ Or trm(r!pstatus2) = "" Then
          'MsgBox "Dieser Platz geht nicht: " & r!statdtg
          GoTo notme_2abo
        End If
        r.MoveNext
      Wend
      Call setaboplaetze(h$, p$, pnr%, rnr%, aboid$, aboname$)
notme_2abo:
    End If
  Next i%
  Call Command13_Click
End If

exme:
MousePointer = 0
co% = -1
Call beglist_Change

End Sub

Private Sub menu_abol_del_Click()
Dim rrr
Dim hbpsid$, i%, brk%, rnr%, pnr%, j%, l1%, l$, c$, h$, p$, aboid$, dtg$
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "menu_abol_del_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ <> "" And p$ <> "" Then

If co% < 0 Then
  dtg$ = trm(Text5.text)
  l$ = trm(beglist.text)
  If l$ <> "" Then
    dtg$ = trm(dtg$ & " " & l$)
    dtg$ = Left$(dtg$, 20)
  End If
  If PlatzselektionVorhanden() = 0 Then
    MousePointer = 11: DoEvents
    If Len(dtg$) = 16 Then
      'platzdaten für den eingetragenen tag/beginn finden
      l$ = platzliste.List(platzliste.ListIndex)
      rnr% = Val(l$)
      l$ = Mid$(l$, InStr(l$, "/") + 1)
      pnr% = Val(Left$(l$, InStr(l$, " ")))
      c$ = "SELECT hbpstatus.pstatus2 as aboid,hbpstatus.hbpid as hbpid " + _
           "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
           "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') " + _
                "AND ((hbpstatus.pstatus)='Abo') AND ((hbpstatus.dtg)='" & dtg$ & "') " + _
                "AND ((hbplist.reihe)=" & rnr% & ") AND ((hbplist.platz)=" & pnr% & "));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        c$ = "delete from hbpstatus where pstatus2='" & r!aboid & "' and hbpid='" & r!hbpid & "'"
        Call form1.sqlqry(c$)
      End If
    End If
    MousePointer = 0
    Call beglist_Change
    Call Command13_Click
  Else
    MousePointer = 11: DoEvents
    For i% = 0 To selerg.ListCount - 1
      If selerg.Selected(i%) = True Then
        l$ = selerg.List(i%)
        j% = InStr(l$, "Reihe "): rnr% = Val(Mid$(l$, j% + 6))
        j% = InStr(l$, "Platz "): pnr% = Val(Mid$(l$, j% + 6))
        c$ = "SELECT hbpstatus.pstatus2 as aboid,hbpstatus.hbpid as hbpid " + _
           "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
           "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') " + _
                "AND ((hbpstatus.pstatus)='Abo') AND ((hbpstatus.dtg)='" & dtg$ & "') " + _
                "AND ((hbplist.reihe)=" & rnr% & ") AND ((hbplist.platz)=" & pnr% & "));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          c$ = "delete from hbpstatus where pstatus2='" & r!aboid & "' and hbpid='" & r!hbpid & "'"
          Call form1.sqlqry(c$)
        End If
      End If
    Next i%
    MousePointer = 0
    Call Command13_Click
  End If
End If
End If
co% = -1


End Sub

Private Sub menu_bestell_Click()
Dim rrr
Dim hbpsid$, i%, brk%, rnr%, pnr%, j%, l1%, l$, c$, vorgang$, aid$, kid$
Dim r As ADODB.Recordset, ra As ADODB.Recordset, rc$, rb As ADODB.Recordset, reihe, platz
Dim cl$, px, py, obb%, obt%, j1%, col, offy%, dy, h$, p$, dtg$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "menu_bestell_Click"
If co% < 0 Then
  vorgang$ = "Bestellung"
  h$ = trm(hid.text)
  p$ = trm(pgid.text)
  dtg$ = trm(Text5.text)
  If dtg$ <> "" Then dtg$ = dtg$ & " " & beglist.text
  If menu_bestell_override_vorgang$ <> "" Then
    vorgang$ = menu_bestell_override_vorgang$
    menu_bestell_override_vorgang$ = ""
  End If
  If trm(Text6.text) = "" Then
    i% = MsgBox("Sie haben keinen Kunden angegeben, ok?", vbOKCancel)
    If i% = 2 Then
      co% = -1
      Exit Sub
    End If
    Text6.text = "unbekannt"
    DoEvents
  End If
  If PlatzselektionVorhanden() = 0 Then
    If h$ <> "" And trm(p$) <> "" Then
      'platzdaten für den eingetragenen tag/beginn finden
      l$ = platzliste.List(platzliste.ListIndex)
      reihe = Val(l$)
      l$ = Mid$(l$, InStr(l$, "/") + 1)
      platz = Val(Left$(l$, InStr(l$, " ")))
      If platzistverkauft(h$, p$, dtg$, platz, reihe) = 0 Then
        MousePointer = 11: DoEvents
        hbpsid$ = get_hbpstatusid(vorgang$, 1)
        If hbpsid$ <> "" Then
          aid$ = trm(Text6.text)
          If aid$ <> "" Then
            i% = InStr(aid$, "|")
            If i% = 0 Then
              c$ = "update hbpstatus set adrid='" & aid$ & "' where id='" & hbpsid & "'"
            Else
              kid$ = Mid$(aid$, i% + 1)
              aid$ = Left$(aid$, i% - 1)
              c$ = "update hbpstatus set adrid='" & aid$ & "' kontakt='" & kid$ & "' where id='" & hbpsid & "'"
            End If
            Call form1.sqlqry(c$)
            'aboplätze markieren
            c$ = "SELECT hbpstatus_1.aboplatzid as abpid,hbpstatus_1.pstatus2 as aboid " + _
               "FROM hbpstatus AS hbpstatus_1 INNER JOIN (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbpstatus_1.hbpid = hbplist.id) " + _
               "WHERE (((hbpstatus.id)='" & hbpsid$ & "') AND ((hbpstatus_1.pstatus)='Abo'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
            If Not r.EOF Then
              c$ = "SELECT hbpstatus.dtg as tdtg, hbplist.id as hlid, * FROM (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) INNER JOIN hbpstatus AS hbpstatus_1 ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbplist.id = hbpstatus_1.hbpid) " + _
                 "WHERE (((hbpstatus.aboplatzid)=" & r!abpid & ") AND ((hbpstatus.pstatus2)='" & r!aboid & "')  AND ((hbpstatus.pstatus)='Abo') AND ((hbpstatus_1.pstatus)='Abo'));"
Set ra = New ADODB.Recordset
ra.CursorLocation = adUseServer
rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
              If Not ra.EOF Then
                l$ = "delete from hbpstatus where id='" & hbpsid$ & "'"
                Call form1.sqlqry(l$)
              End If
              While Not ra.EOF
                rc$ = form1.newid("hbpstatus", "id", 26)
                c$ = "insert into hbpstatus (id,hbpid,adrid,dtg,pstatus,dat_creat,dat_upd) values('" + _
                           rc$ & "','" + _
                           ra!hlid & "','" + _
                           aid$ & "','" + _
                           ra!tdtg & "','" + _
                           "Bestellung" & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "')"
                Call form1.sqlqry(c$)
                If kid$ <> "" Then
                  c$ = "update hbpstatus set kontakt='" & kid$ & "' where id='" & rc$ & "'"
                  Call form1.sqlqry(c$)
                End If
                ra.MoveNext
              Wend
            End If
          End If
          Call beglist_Change
        End If
        MousePointer = 0
      End If
    End If
  Else
    MousePointer = 11: DoEvents
    For i% = 0 To selerg.ListCount - 1
      If selerg.Selected(i%) = True Then
        l$ = selerg.List(i%)
        j% = InStr(l$, "Reihe "): rnr% = Val(Mid$(l$, j% + 6))
        j% = InStr(l$, "Platz "): pnr% = Val(Mid$(l$, j% + 6))
        l$ = Format$(rnr%, "0###") & "/" & Format$(pnr%, "0###") & "  "
        l1% = Len(l$)
        For j% = 0 To platzliste.ListCount - 1
          If Left(platzliste.List(j%), l1%) = l$ Then
            platzliste.ListIndex = j%
            l$ = platzliste.List(platzliste.ListIndex)
            reihe = Val(l$)
            l$ = Mid$(l$, InStr(l$, "/") + 1)
            platz = Val(Left$(l$, InStr(l$, " ")))
            If platzistverkauft(h$, p$, dtg$, platz, reihe) = 0 Then

            hbpsid$ = get_hbpstatusid(vorgang$, 0)
            If hbpsid$ <> "" And trm(Text6.text) <> "" Then
              c$ = "update hbpstatus set adrid='" & Text6.text & "' where id='" & hbpsid & "'"
              Call form1.sqlqry(c$)
              cl$ = platzliste.List(j%)
              c$ = word2(cl$)
              px = Val(c$)
              py = Val(Mid$(c$, InStr(c$, "/") + 1))
              j1% = InStr(cl$, "|obt="): obt% = Val(Mid$(cl$, j1% + 5))
              j1% = InStr(cl$, "|obb="): obb% = Val(Mid$(cl$, j1% + 5))
              dy = obt% / 4
              Select Case LCase(vorgang$)
                Case "bestellung": col = RGB(0, 255, 255)
                           offy% = 1
                Case "verkauft": col = RGB(0, 0, 128)
                           offy% = 1
                Case "ehrenplatz": col = RGB(255, 0, 0)
                           offy% = 0
                Case "abo": col = RGB(0, 255, 0)
                           offy% = 2
                Case "abo verkauft": col = RGB(0, 0, 0)
                           offy% = 2
                Case Else:  col = RGB(0, 0, 0)
                            offy% = 3
              End Select
              P1.Line (px - obb% / 2, py - obt% / 2 + (offy% + 1) * dy)-(px + obb% / 2, py - obt% / 2 + offy% * dy), col, BF
DoEvents
          'aboplätze markieren
        c$ = "SELECT hbpstatus_1.aboplatzid as abpid,hbpstatus_1.pstatus2 as aboid " + _
           "FROM hbpstatus AS hbpstatus_1 INNER JOIN (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbpstatus_1.hbpid = hbplist.id) " + _
           "WHERE (((hbpstatus.id)='" & hbpsid$ & "') AND ((hbpstatus_1.pstatus)='Abo'));"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          c$ = "SELECT hbpstatus.dtg as tdtg, hbplist.id as hlid, * FROM (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) INNER JOIN hbpstatus AS hbpstatus_1 ON (hbpstatus.dtg = hbpstatus_1.dtg) AND (hbplist.id = hbpstatus_1.hbpid) " + _
             "WHERE (((hbpstatus.aboplatzid)=" & r!abpid & ") AND ((hbpstatus.pstatus2)='" & r!aboid & "')  AND ((hbpstatus.pstatus)='Abo') AND ((hbpstatus_1.pstatus)='Abo'));"
Set ra = New ADODB.Recordset
ra.CursorLocation = adUseServer
rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
          If Not ra.EOF Then
            l$ = "delete from hbpstatus where id='" & hbpsid$ & "'"
            Call form1.sqlqry(l$)
          End If
          While Not ra.EOF
            rc$ = form1.newid("hbpstatus", "id", 26)
            c$ = "insert into hbpstatus (id,hbpid,adrid,dtg,pstatus,dat_creat,dat_upd) values('" + _
                           rc$ & "','" + _
                           ra!hlid & "','" + _
                           Text6.text & "','" + _
                           ra!tdtg & "','" + _
                           "Bestellung" & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "','" + _
                           datum2sql(Date) & " " & Left(Time, 5) & "')"
            Call form1.sqlqry(c$)
            If kid$ <> "" Then
              c$ = "update hbpstatus set kontakt='" & kid$ & "' where id='" & rc$ & "'"
              Call form1.sqlqry(c$)
            End If
            ra.MoveNext
          Wend
        End If


            End If
            Exit For
          End If
          End If
        Next j%
      End If
    Next i%
    MousePointer = 0
    'Call Command13_Click
  End If
End If
Text6.text = ""
While selerg.ListCount > 0: selerg.RemoveItem 0: Wend
For i% = 0 To rowlist.ListCount - 1: rowlist.Selected(i%) = False: Next i%
co% = -1

End Sub

Public Sub menu_bestell_del_Click()
Dim hbpsid$, i%, brk%, rnr%, pnr%, j%, l1%, l$, c$, vorgang$

'd2infile = "splan": d2insub = "menu_bestell_del_Click"
If co% < 0 Then
  vorgang$ = "Retour-Bestellung"
  If PlatzselektionVorhanden() = 0 Then
    MousePointer = 11: DoEvents
    Do
      hbpsid$ = get_hbpstatusid(vorgang$, 1)
      If hbpsid$ <> "" Then
        c$ = "delete from hbpstatus where id='" & hbpsid & "'"
        Call form1.sqlqry(c$)
      End If
    Loop Until hbpsid$ = ""
    Call beglist_Change
    MousePointer = 0
    Call Command13_Click
  Else
    MousePointer = 11: DoEvents
    For i% = 0 To selerg.ListCount - 1
      If selerg.Selected(i%) = True Then
        l$ = selerg.List(i%)
        j% = InStr(l$, "Reihe "): rnr% = Val(Mid$(l$, j% + 6))
        j% = InStr(l$, "Platz "): pnr% = Val(Mid$(l$, j% + 6))
        l$ = Format$(rnr%, "0###") & "/" & Format$(pnr%, "0###") & "  "
        l1% = Len(l$)
        For j% = 0 To platzliste.ListCount - 1
          If Left(platzliste.List(j%), l1%) = l$ Then
            platzliste.ListIndex = j%
            Do
              hbpsid$ = get_hbpstatusid(vorgang$, 0)
              If hbpsid$ <> "" Then
                c$ = "delete from hbpstatus where id='" & hbpsid & "'"
                Call form1.sqlqry(c$)
              End If
            Loop Until hbpsid$ = ""
            Exit For
          End If
        Next j%
      End If
    Next i%
    MousePointer = 0
    Call Command13_Click
  End If
End If
co% = -1


End Sub

Private Sub menu_bestell_kunde_Click()
Dim i%, l$

'd2infile = "splan": d2insub = "menu_bestell_kunde_Click"
If knownwbegs.ListCount = 0 Then
  MsgBox "Für diesen Eintrag sind keine Kundendaten bekannt"
Else
  i% = knownwbegs.ListIndex
  If i% < 0 Then i% = 0
  l$ = trm(Mid$(knownwbegs.List(i%), 8))
  i% = InStr(l$, "|")
  If i% > 0 Then l$ = Left$(l$, i% - 1)
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(l$, "")
  On Error Resume Next
  Call shwAdrDetail.SetFocus
  On Error GoTo 0

End If

co% = -1
End Sub

Private Sub menu_ehre_Click()
'd2infile = "splan": d2insub = "menu_ehre_Click"
menu_bestell_override_vorgang$ = "Ehrenplatz"
Call menu_bestell_Click
End Sub

Private Sub menu_kbuch_Click()
'd2infile = "splan": d2insub = "menu_kbuch_Click"
Load kbuch
    On Error Resume Next
    Call kbuch.SetFocus
    On Error GoTo 0

End Sub

Public Sub menu_platz_kunde_Click()
'd2infile = "splan": d2insub = "menu_platz_kunde_Click"
Call Command18_Click
co% = -1
End Sub

Private Sub menu_retour_Click()
Dim hbpsid$, i%, brk%, rnr%, pnr%, j%, l1%, l$, c$, vorgang$

'd2infile = "splan": d2insub = "menu_retour_Click"
If co% < 0 Then
  vorgang$ = "Retour-VK"
  If PlatzselektionVorhanden() = 0 Then
    MousePointer = 11: DoEvents
    hbpsid$ = get_hbpstatusid(vorgang$, 1)
    If hbpsid$ <> "" Then
        c$ = "update hbpstatus set adrid='" & Text6.text & "' where id='" & hbpsid & "'"
        'Call form1.sqlqry(c$)
      Call beglist_Change
    End If
    MousePointer = 0
  Else
    MousePointer = 11: DoEvents
    For i% = 0 To selerg.ListCount - 1
      If selerg.Selected(i%) = True Then
        l$ = selerg.List(i%)
        j% = InStr(l$, "Reihe "): rnr% = Val(Mid$(l$, j% + 6))
        j% = InStr(l$, "Platz "): pnr% = Val(Mid$(l$, j% + 6))
        l$ = Format$(rnr%, "0###") & "/" & Format$(pnr%, "0###") & "  "
        l1% = Len(l$)
        For j% = 0 To platzliste.ListCount - 1
          If Left(platzliste.List(j%), l1%) = l$ Then
            platzliste.ListIndex = j%
            hbpsid$ = get_hbpstatusid(vorgang$, 0)
            If hbpsid$ <> "" And trm(Text6.text) <> "" Then
              c$ = "update hbpstatus set adrid='" & Text6.text & "' where id='" & hbpsid & "'"
              'Call form1.sqlqry(c$)
            End If
            Exit For
          End If
        Next j%
      End If
    Next i%
    MousePointer = 0
    Call Command13_Click
  End If
End If
co% = -1

End Sub

Private Sub menu_select_row_Click()
Dim i%, l$

'd2infile = "splan": d2insub = "menu_select_row_Click"
i% = platzliste.ListIndex
If i% < 0 Then Exit Sub
l$ = "Reihe " & Val(platzliste.List(i%))
For i% = 0 To rowlist.ListCount - 1
  If rowlist.List(i%) = l$ Then
    rowlist.Selected(i%) = True
    Call Command8_Click
    Exit For
  End If
Next i%
'Call beglist_Change
co% = -1

End Sub

Private Sub menu_unselect_Click()
'd2infile = "splan": d2insub = "menu_unselect_Click"
'Call Command13_Click
Dim i%, col As Long, c$, r As ADODB.Recordset, l$, px, py, obb%, obt%, j%

For i% = 0 To selerg.ListCount - 1

l$ = selerg.List(i%)
j% = InStr(l$, "(px:"): px = Val(Mid$(l$, j% + 4))
j% = InStr(l$, "(py:"): py = Val(Mid$(l$, j% + 4))
j% = InStr(l$, "(obt:"): obt% = Val(Mid$(l$, j% + 5))
j% = InStr(l$, "(obb:"): obb% = Val(Mid$(l$, j% + 5))
If selerg.Selected(i%) = True Then
  selerg.Selected(i%) = False
  col = RGB(255, 255, 255)
  P1.Circle (px, py), imin(obb%, obt%) / 4, col
End If


Next i%
While selerg.ListCount > 0
  selerg.RemoveItem 0
Wend
For i% = 0 To rowlist.ListCount - 1
  If rowlist.Selected(i%) = True Then rowlist.Selected(i%) = False
Next i%
For i% = 0 To pglist.ListCount - 1
  If pglist.Selected(i%) = True Then pglist.Selected(i%) = False
Next i%

co% = -1

End Sub

Private Sub menu_vk_Click()
Dim rrr
Dim c$, r As ADODB.Recordset, h$, p$, dtg$, l$, i%, k$, lvitem
Dim kbdtg$, vg$, vgl$, epreisnetto As Double, mwst As Double, vdtg$, j%, rc$
Dim epn As Double, epm As Double, epb As Double, dtmp As Double
Dim ra As ADODB.Recordset, rb As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "menu_vk_Click"
kbdtg$ = datum2sql(Date) & " " & Time
mwst = d2db(mwstvk.text)
k$ = trm(Text7.text)
If k$ = "" Then k$ = trm(Text6.text)
If k$ = "" Then
  k$ = "Barverkauf " & form1.getuserid() & " " & Date & " " & Time
  Text6.text = k$
End If
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Or p$ = "" Then GoTo ex_menu_vk_Click
dtg$ = trm(Text5.text)
If dtg$ = "" Then GoTo ex_menu_vk_Click
l$ = trm(beglist.text)
If l$ = "" Then GoTo ex_menu_vk_Click
Load kvk
On Error Resume Next
Call kvk.SetFocus
On Error GoTo 0
Call kvk.gd1_clear
kvk.Caption = "Kartenverkauf"
MousePointer = 11: DoEvents
dtg$ = trm(dtg$ & " " & l$)
dtg$ = Left$(dtg$, 20)
c$ = "SELECT hbplist.hid, hbplist.pgid, hbplist.id as hbplid, hbplist.preis as gpreis, hbplist.reihe as preihe, hbplist.platz as pplatz, hbpstatus.dtg as vdtg, hbpstatus.preis as spreis, hbpstatus.pstatus2 as aboid, hbpstatus.pstatus, hbpstatus.aboplatzid as abpid, hbplist.platzname " + _
     "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
     "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') AND ((hbpstatus.dtg)='" & dtg$ & "') AND ((hbpstatus.adrid)='" & k$ & "')) " + _
     "order by hbplist.reihe, hbplist.platz;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
kvk.Combo1(0).text = k$
epn = 0: epm = 0: epb = 0: vg$ = ""
If r.EOF Then       'Karten f. Kunde nicht erkennbar, selektierte probieren
  If PlatzselektionVorhanden() = 0 Then
    co% = -1
    Call p1_MouseDown(1, 0, m_pressx, m_pressy)
    co% = -2
  End If
  Call menu_bestell_Click
  c$ = "SELECT hbplist.hid, hbplist.pgid, hbplist.id as hbplid, hbplist.preis as gpreis, hbplist.reihe as preihe, hbplist.platz as pplatz, hbpstatus.dtg as vdtg, hbpstatus.preis as spreis, hbpstatus.pstatus, hbpstatus.pstatus2 as aboid, hbpstatus.aboplatzid as abpid, hbplist.platzname " + _
     "FROM hbplist INNER JOIN hbpstatus ON hbplist.id = hbpstatus.hbpid " + _
     "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') AND ((hbpstatus.dtg)='" & dtg$ & "') AND ((hbpstatus.adrid)='" & k$ & "')) " + _
     "order by hbplist.reihe, hbplist.platz;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
End If
While Not r.EOF
  If LCase(r!pstatus) <> "verkauft" Then
    l$ = r!platzname
    vdtg$ = datfromsql(word1(r!vdtg)) & Mid$(r!vdtg, InStr(r!vdtg, " "))
    If vg$ = "" Then
      vg$ = vdtg$ & " " & h$ & " " & p$ & " " & " " & r!platzname
      kvk.vg.Caption = vg$
    End If
'    c$ = "SELECT hbpstatus.aboplatzid,hbpstatus.adrid,hbpstatus.pstatus2 " + _
'         "FROM hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id " + _
'         "WHERE (((hbpstatus.hbpid)='" & r!hbplid & "') and ((hbpstatus.pstatus)='Abo'));"
    c$ = "SELECT hbpstatus.aboplatzid, hbpstatus.adrid, hbpstatus.preis, hbpstatus.pstatus2 " + _
         "FROM (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) INNER JOIN hbabotermine ON (hbabotermine.pid = hbplist.pgid) AND (hbabotermine.adrid = hbplist.hid) AND (hbabotermine.dtg = hbpstatus.dtg) AND (hbpstatus.pstatus2 = hbabotermine.aboid) " + _
         "WHERE (((hbpstatus.hbpid)='" & r!hbplid & "') and ((hbpstatus.pstatus)='Abo') AND ((hbabotermine.dtg)='" & r!vdtg & "'));"
Set ra = New ADODB.Recordset
ra.CursorLocation = adUseServer
rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
    If Not ra.EOF Then
      'abo
      epreisnetto = 0
      If trm(ra!preis) <> "" Then epreisnetto = ra!preis
      vgl$ = ra!adrid & " Nr. " & ra!aboplatzid & " bestehend aus:"
      c$ = "SELECT hbabos.Name, hbabotermine.adrid, hbabotermine.pid, hbabotermine.dtg as adtg, hbplist.platzname, hbabos.preis, hbpstatus.aboplatzid, hbpstatus.dtg as tdtg, hbabotermine.aboid " + _
         "FROM (hbabotermine INNER JOIN hbabos ON hbabotermine.aboid = hbabos.id) INNER JOIN (hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id) ON (hbabotermine.dtg = hbpstatus.dtg) AND (hbabotermine.pid = hbplist.pgid) AND (hbabotermine.adrid = hbplist.hid) " + _
         "Where (((hbpstatus.aboplatzid) =" & ra!aboplatzid & ") And ((hbabotermine.aboid) ='" & ra!pstatus2 & "')) " + _
         "ORDER BY hbabos.Name, hbabotermine.dtg, hbplist.reihe, hbplist.platz;"
Set ra = New ADODB.Recordset
ra.CursorLocation = adUseServer
rrr = form1.adoopen(ra, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      kvk.ralist.Clear
      While Not ra.EOF
        kvk.ralist.AddItem datfromsql(word1(ra!tdtg)) & Mid$(ra!tdtg, InStr(ra!tdtg, " ")) & " " & ra!adrid & " " & ra!pid & " " & " " & ra!platzname
        ra.MoveNext
      Wend
    Else
      'kein abo
      vgl$ = vdtg$ & " " & h$ & " " & p$ & " " & " " & r!platzname
      epreisnetto = r!gpreis
      If trm(r!spreis) <> "" Then epreisnetto = r!spreis
      c$ = "select * from hbpstatus where hbpid='" & r!hbplid & "' and dtg='" & dtg$ & "'"
Set rb = New ADODB.Recordset
rb.CursorLocation = adUseServer
rrr = form1.adoopen(rb, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
      If Not rb.EOF Then
        If Not IsNull(rb!preis) Then
          If trm(rb!preis) <> "" Then epreisnetto = rb!preis
        End If
      End If
    End If
    epreisnetto = (epreisnetto / (100 + mwst)) * 100
    epn = epn + epreisnetto
    dtmp = mwst / 100# * epreisnetto
    epm = epm + dtmp
    epb = epb + dtmp + epreisnetto
    Set lvitem = kvk.gd1.ListItems.add(, , "1")
    lvitem.SubItems(1) = vgl$
    lvitem.SubItems(2) = fixeur(epreisnetto)
    lvitem.SubItems(3) = fixeur(mwst)
    lvitem.SubItems(4) = fixeur(dtmp)
    lvitem.SubItems(5) = fixeur(dtmp + epreisnetto)
    kvk.gd1ids.AddItem r!hbplid
    While kvk.ralist.ListCount > 0
      Set lvitem = kvk.gd1.ListItems.add(, , " ")
      lvitem.SubItems(1) = kvk.ralist.List(0)
      kvk.gd1ids.AddItem "NULL"
      kvk.ralist.RemoveItem 0
    Wend
  End If
  r.MoveNext
Wend

kvk.epn.Caption = fixeur(epn)
kvk.epm.Caption = fixeur(epm)
kvk.epb.Caption = "Endbetrag: " & fixeur(epb)
kvk.epb.Enabled = True

MousePointer = 0

ex_menu_vk_Click:
co% = -1
End Sub

Private Sub mwstvk_Change()
'd2infile = "splan": d2insub = "mwstvk_Change"
Call form1.setmylastFormVar(Me.name, "mwst_vk", trm(mwstvk.text))
End Sub

Private Sub neuerpreis_Change()
'd2infile = "splan": d2insub = "neuerpreis_Change"
If selerg.ListCount > 0 And neuerpreis.text <> "" Then Command11.Enabled = True

End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rrr
Dim i%, xo%, yoff%, obb As Long, obt As Long, obn As Long
Dim pnr%, px%, py%, j%, rnr%, r As ADODB.Recordset, neuid$, dtg$, l$
Dim pname$, kpname$, h$, p$, c$, pg$, preis As Double, waehr$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "p1_MouseDown"
If save_in_progress% = 1 Then Exit Sub
If Button = 2 Then
  co% = -2
  m_pressx = X
  m_pressy = Y
  PopupMenu menu_bearb
  Exit Sub
End If

Select Case co%
  Case -2: co% = -1
  Case -1: 'selecting a seat
            h$ = trm(hid.text)
            p$ = trm(pgid.text)
            If h$ <> "" And trm(p$) <> "" Then
              dtg$ = trm(Text5.text)
              If Len(dtg$) = 10 Then
                i% = platzliste.ListIndex
                If i% >= 0 Then
                  l$ = platzliste.List(i%)
                  rnr% = Val(l$)
                  j% = InStr(l$, "/"): l$ = word1(Mid$(l$, j% + 1)): pnr% = Val(l$)
                  c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and platz=" & pnr% & " and reihe=" & rnr%
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
                  If Not r.EOF Then
                    Command13.Enabled = True
                    l$ = r!platzname & Space$(80) & "(ID:" & r!id & Space$(8) & "(px:" & r!px & Space$(8) & "(py:" & r!py & Space$(8) & "(obb:" & r!obb & Space$(8) & "(obt:" & r!obt
                    For i% = 0 To selerg.ListCount - 1
                      If l$ = selerg.List(i%) Then i% = selerg.ListCount + 10
                    Next i%
                    If i% < selerg.ListCount + 5 Then
                      selerg.AddItem l$
                      For i% = 0 To selerg.ListCount - 1
                        If selerg.Selected(i) = False Then
                          selerg.Selected(i) = True
                          DoEvents
                        End If
                      Next i%
                    End If
                  End If
                End If
              End If
            End If
  Case 0: 'placing a row
           co% = -1
          xo% = Val(nrows.text)
          obn = Val(Text3.text)
          obb = Val(bbreit(0).text)
          obt = Val(bbreit(1).text)
          rnr% = Val(bbreit(3).text)
          h$ = trm(hid.text)
          p$ = trm(pgid.text)
          If h$ = "" Then Exit Sub
          save_in_progress% = 1
          On Error Resume Next
          Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & ".pln"
          Kill form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & "*.gif"
          On Error GoTo 0
          MousePointer = 11
          DoEvents
          pg$ = Combo1.text
          i% = InStr(pg$, " - (")
          If i% > 0 Then pg$ = Left$(pg$, i% - 1)
          preis = 0
          waehr$ = ""
          c$ = "SELECT id,preis,waehrung FROM preisgruppen where id='" & pg$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
          If Not r.EOF Then
            waehr$ = r!waehrung
            preis = r!preis
          End If
          For i% = 1 To xo%
            yoff% = ((i% - 1) * obt)
            'py% = y - obt / 2 - yoff%
            py% = Y - yoff%
            px% = X - (obb * obn) / 2 - obb / 2
            pnr% = Val(bbreit(2).text)
            For j% = 1 To obn
              px% = px% + obb
              kpname$ = rnr% & "/" & trm(pnr%)
'              pname$ = trm(pgid.Text & " Reihe") & " " & rnr% & " Platz " & pnr%
              pname$ = "Reihe" & " " & rnr% & " Platz " & pnr%
              c$ = "select id from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and pid='" & kpname$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
              If Not r.EOF Then
                neuid$ = r!id
              Else
                neuid$ = form1.newid("hbplist", "id", 30)
              End If
              c$ = "delete from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and pid='" & kpname$ & "'"
              Call form1.sqlqry(c$)
              c$ = "insert into hbplist (id,hid,pgid,pid,platzname,preisgruppe,preis,reihe,platz,px,py,obb,obt,waehrung) values('" + _
                    neuid$ & "','" + _
                    h$ & "','" + _
                    p$ & "','" + _
                    kpname$ & "','" + _
                    pname$ & "','" + _
                    pg$ & "'," + _
                    d2db(preis) & "," + _
                    rnr% & "," + _
                    pnr% & "," + _
                    px & "," + _
                    py & "," + _
                    obb & "," + _
                    obt & ",'" + _
                    waehr$ & "')"
              Call form1.sqlqry(c$)
              P1.Line (px% + obb / 2, py% + obt / 2)-(px% - obb / 2, py% - obt / 2), RGB(0, 0, 0), B
              P1.Line (px% + obb / 2, py% + obt / 2)-(px% - obb / 2, py% - obt / 2), RGB(255, 255, 255)
              P1.Print kpname$
              DoEvents
              If cntmode.value = 0 Then
                pnr% = pnr% - 1
              Else
                pnr% = pnr% + 1
              End If
            Next j%
            If rcntmode.value = 0 Then
              rnr% = rnr% - 1
            Else
              rnr% = rnr% + 1
            End If
          Next i%
          MousePointer = 0
          Call Command2_Click
          nupd% = 1
          For i% = 0 To chm%
           Check1(i%) = 0
          Next i%
          nupd% = 0
          Call setminrow
          save_in_progress% = 0
  Case Else:
End Select

End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rrr
Dim cmx As Double, cmy As Double, obb As Long, obt As Long, dtg$, c$, r As ADODB.Recordset
Dim i%, cl%, p1x As Long, p1y As Long, l$, dst As Long, dx As Long, dy As Long
Dim h$, p$, id$, reihe, platz, prs$, prsv As Double

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "p1_MouseMove"
cmy = P1.ScaleHeight * Y / P1.ScaleHeight
cmx = P1.ScaleWidth * X / P1.ScaleWidth
Label3.Caption = "x=" & trm(Int(cmx) / 100) & " m / y=" & trm(Int(cmy) / 100) & " m"
knownwbegs.Clear
If co% = -1 Then

  cl% = -1
  dst = -1
  For i% = 0 To platzliste.ListCount - 1
    l$ = platzliste.List(i%)
    co% = InStr(l$, Space$(10))
    l$ = trm(Mid$(l$, co%))
    p1x = Val(l$)
    co% = InStr(l$, "/")
    p1y = Val(trm(Mid$(l$, co% + 1)))
    dx = p1x - X
    dy = p1y - Y
    dx = dx * dx: dy = dy * dy
    If dst < 0 Or dst > dx + dy Then
      dst = dx + dy
      cl% = i%
    End If
  Next i%
  co% = -1
  If cl% >= 0 Then
    Text7.text = ""
    platzliste.ListIndex = cl%
    h$ = trm(hid.text)
    p$ = trm(pgid.text)
    If h$ <> "" And trm(p$) <> "" Then
      dtg$ = trm(Text5.text)
      If dtg$ <> "" Then
        'platzdaten für den eingetragenen tag/beginn finden
        l$ = platzliste.List(platzliste.ListIndex)
        reihe = Val(l$)
        l$ = Mid$(l$, InStr(l$, "/") + 1)
        platz = Val(Left$(l$, InStr(l$, " ")))
        c$ = "select id,preis,waehrung from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and platz=" & platz & " and reihe=" & reihe
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
        If Not r.EOF Then
          Text4(3).text = fixeur(r!preis) & " " & r!waehrung
          prs$ = r!waehrung
          prsv = r!preis
          id$ = r!id
          c$ = "select * from hbpstatus  where hbpid='" & id$ & "' and instr(dtg,'" & dtg$ & "')=1"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
          dtg$ = trm(dtg$ & " " & trm(beglist.text))
          dtg$ = Left$(dtg$, 20)
          While Not r.EOF
            l$ = trm(Mid(r!dtg, InStr(r!dtg, " "))) & " " & Left(r!pstatus, 1) & " " & r!adrid
            If Left(r!pstatus, 1) = "A" Then l$ = l$ & " (" & r!aboplatzid & ")"
            knownwbegs.AddItem l$
            If dtg$ = r!dtg Then
              If LCase(Left(r!adrid, 3)) <> "abo" Then
                Text7.text = trm(r!adrid)
                If InStr(Text7.text, "|") > 0 Then Text7.text = Left(Text7.text, InStr(Text7.text, "|") - 1)
              End If
            End If
            If Not IsNull(r!waehrung) Then
              If trm(r!waehrung) <> "" Then prs$ = r!waehrung
            End If
            If Not IsNull(r!preis) Then
              Text4(3).text = fixeur(r!preis) & " " & prs$
            End If
            r.MoveNext
          Wend
          For i% = 0 To knownwbegs.ListCount - 1
            If InStr(knownwbegs.List(i%), beglist.text) > 0 Then
              knownwbegs.ListIndex = i%
              Exit For
            End If
          Next i%
        End If
      End If
    End If
  End If

Else  'cl%>=0


    obb = Val(Text3.text) * Val(bbreit(0).text)
    obt = Val(bbreit(1).text)
    If px% > 0 Then Call drw(co%, px%, py%, RGB(255, 255, 255), obb, obt, "")
    Call drw(co%, X, Y, RGB(0, 0, 0), obb, obt, "")
    px% = X: py% = Y


End If



End Sub

Private Sub pgid_Click()
'd2infile = "splan": d2insub = "pgid_Click"
If noupd% = 1 Then Exit Sub

Call savecheck
Call Command2_Click
Call beglist_Change

End Sub

Private Sub pgid_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "pgid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT * FROM hblist where hid='" & hid.text & "' order by pgid;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
pgid.Clear
While Not r.EOF
  pgid.AddItem r!pgid
  r.MoveNext
Wend
MousePointer = 0


End Sub

Private Sub pgid1_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "pgid1_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT * FROM hblist where hid='" & hid1.text & "' order by pgid;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
pgid1.Clear
While Not r.EOF
  pgid1.AddItem r!pgid
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Sub pglist_Click()

'd2infile = "splan": d2insub = "pglist_Click"
selstr_pglist.text = pglist_getsel

End Sub

Private Sub pgshowlist_Click()
Dim rrr
Dim sel_str$, h$, p$, c$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "pgshowlist_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then Exit Sub
P1.Cls
P1.Picture = p0.Picture
DoEvents
If exist(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif") = 0 Then Call Command2_Click
P1.Picture = LoadPicture(form1.s0dir() & "\" + form1.medien() & "\" & form1.medienname(h$) & "\saalplan\" & p$ & trm(Check2.value) & trm(Check3.value) & ".gif")
sel_str$ = pgshowlist_getsel()
If sel_str$ <> "" Then
  If Me.Width > P1.Left + P1.Width + 200 Then
    Call Command5_Click
  End If
  c$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' "
  If sel_str$ <> "" Then c$ = c$ & " AND (" & sel_str$ & ")"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
  While Not r.EOF
    P1.Circle (r!px, r!py), imin(r!obb, r!obt) / 4, RGB(0, 0, 0)
    DoEvents
    r.MoveNext
  Wend
End If
Call beglist_Change

End Sub

Private Sub platzliste_Click()
Dim rrr
Dim i%, reihe, platz, l$, r As ADODB.Recordset, h$, p$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "platzliste_Click"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If

i% = platzliste.ListIndex
If i% < 0 Then Exit Sub
l$ = platzliste.List(i%)
reihe = Val(l$)
l$ = Mid$(l$, InStr(l$, "/") + 1)
platz = Val(Left$(l$, InStr(l$, " ")))
l$ = "select * from hbplist where hid='" & h$ & "' and pgid='" & p$ & "' and reihe=" & reihe & " and platz=" & platz
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, l$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  Text4(0).text = r!platzname
  Text4(1).text = r!preisgruppe
  Text4(2).text = fixeur(r!preis) & " " & r!waehrung
End If

End Sub

Private Sub possabo_DblClick()
Dim rrr
Dim i%, id$, j%, r As ADODB.Recordset, rc$, h$, p$, slx$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "possabo_DblClick"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Or p$ = "" Then Exit Sub

i% = possabo.ListIndex
If i% < 0 Then Exit Sub
id$ = possabo.List(i%)
j% = InStr(id$, "(ID:") + 4
id$ = Mid$(id$, j%)
slx$ = id$
rc$ = "select * from hbabotermine where aboid='" & id$ & "' and adrid='" & h$ & "' and pid='" & p$ & "' order by dtg"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, rc$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
rc$ = ""
While Not r.EOF
  rc$ = rc$ & r!dtg & vbCrLf
  r.MoveNext
Wend
MsgBox rc$
Load abos
On Error Resume Next
Call abos.SetFocus
On Error GoTo 0
For i% = 0 To abos.aboliste.ListCount - 1
  If InStr(abos.aboliste.List(i%), slx$) > 0 Then
    abos.aboliste.ListIndex = i%
    Exit For
  End If
Next i%

End Sub


Private Sub rid_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$, l$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "rid_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT * FROM hblist where hid='" & hid.text & "' order by pgid;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
rid.Clear
l$ = "----"
While Not r.EOF
  If l$ <> r!raum Then
    rid.AddItem r!raum
    l$ = r!raum
  End If
  r.MoveNext
Wend
MousePointer = 0


End Sub

Private Sub rid1_DropDown()
Dim rrr
Dim r As ADODB.Recordset, c$, l$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "rid1_DropDown"
MousePointer = 11: DoEvents
c$ = "SELECT * FROM hblist where hid='" & hid.text & "' order by pgid;"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
rid1.Clear
l$ = "----"
While Not r.EOF
  If l$ <> r!raum Then
    rid1.AddItem r!raum
    l$ = r!raum
  End If
  r.MoveNext
Wend
MousePointer = 0

End Sub

Private Sub rowlist_Click()
'd2infile = "splan": d2insub = "rowlist_Click"
selstr_rowlist.text = rowlist_getsel
End Sub

Private Sub selerg_Click()
Dim i%, col As Long, c$, r As ADODB.Recordset, l$, px, py, obb%, obt%, j%

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "selerg_Click"
i% = selerg.ListIndex
If i% < 0 Then Exit Sub

l$ = selerg.List(i%)
j% = InStr(l$, "(px:"): px = Val(Mid$(l$, j% + 4))
j% = InStr(l$, "(py:"): py = Val(Mid$(l$, j% + 4))
j% = InStr(l$, "(obt:"): obt% = Val(Mid$(l$, j% + 5))
j% = InStr(l$, "(obb:"): obb% = Val(Mid$(l$, j% + 5))
If selerg.Selected(i%) = True Then
  col = RGB(0, 0, 0)
Else
  col = RGB(255, 255, 255)
End If
P1.Circle (px, py), imin(obb%, obt%) / 4, col
Command13.Enabled = True

End Sub

Private Sub termlist_Click()
Dim c$, i%

'd2infile = "splan": d2insub = "termlist_Click"
If termlist_drw% = 0 Then Exit Sub

i% = termlist.ListIndex
If i% < 0 Then Exit Sub

c$ = termlist.List(i)
Text5.text = ""
beglist.text = trm(Mid$(c$, InStr(c$, " ")))
termlist_upd% = 0
Text5.text = word1(c$)
termlist_upd% = 1

End Sub

Private Sub Text1_Change()
Dim rrr

'd2infile = "splan": d2insub = "Text1_Change"
On Error Resume Next
P1.ScaleHeight = Val(Text1.text) * 100
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
BackColor = form1.dirtycolor()
End Sub

Private Sub Text2_Change()
Dim rrr

'd2infile = "splan": d2insub = "Text2_Change"
On Error Resume Next
P1.ScaleWidth = Val(Text2.text) * 100
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
rrr = Int(P1.ScaleWidth / Val(bbreit(0).text)) - 1
Text3.text = trm(rrr)
BackColor = form1.dirtycolor()

End Sub
Sub savecheck()
'd2infile = "splan": d2insub = "savecheck"
If BackColor = form1.dirtycolor() Then
Dim antw As Integer

  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command4_Click
  End If
End If
BackColor = form1.cleancolor()
End Sub

Sub setminrow()
Dim rrr
Dim r As ADODB.Recordset, h$, p$, c$

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "setminrow"
bbreit(3).text = "1"
h$ = trm(hid.text)
p$ = trm(pgid.text)
If h$ = "" Then Exit Sub
If trm(p$) = "" Then
  p$ = "_leer"
  pgid.text = p$
End If

c$ = "SELECT max(reihe) as cnt FROM hbplist where hid='" & h$ & "' and pgid='" & p$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
If Not r.EOF Then
  bbreit(3).text = Val("0" & r!cnt) + 1
End If

End Sub
Private Function pglist_getsel() As String
Dim s$, i%, pgl$

'd2infile = "splan": d2insub = "pglist_getsel"
s$ = ""
For i% = 0 To pglist.ListCount - 1
  If pglist.Selected(i%) = True Then
    pgl$ = pglist.List(i%)
    pgl$ = trm(Left$(pgl$, InStr(pgl$, " - ")))
    If Len(s$) = 0 Then
      s$ = "((preisgruppe='" + pgl$ + "') "
    Else
      s$ = s$ + "or (preisgruppe='" + pgl$ + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"
pglist_getsel = s$

End Function
Private Function pgshowlist_getsel() As String
Dim s$, i%, pgl$

'd2infile = "splan": d2insub = "pgshowlist_getsel"
s$ = ""
For i% = 0 To pgshowlist.ListCount - 1
  If pgshowlist.Selected(i%) = True Then
    pgl$ = pgshowlist.List(i%)
    'pgl$ = trm(Left$(pgl$, InStr(pgl$, " - ")))
    If Len(s$) = 0 Then
      s$ = "((preisgruppe='" + pgl$ + "') "
    Else
      s$ = s$ + "or (preisgruppe='" + pgl$ + "') "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"
pgshowlist_getsel = s$

End Function

Private Function rowlist_getsel() As String
Dim s$, i%, pgl$

'd2infile = "splan": d2insub = "rowlist_getsel"
s$ = ""
For i% = 0 To rowlist.ListCount - 1
  If rowlist.Selected(i%) = True Then
    pgl$ = rowlist.List(i%)
    pgl$ = trm(Mid$(pgl$, InStr(pgl$, " ")))
    If Len(s$) = 0 Then
      s$ = "((reihe=" + pgl$ + ") "
    Else
      s$ = s$ + "or (reihe=" + pgl$ + ") "
    End If
  End If
Next i%
If Len(s$) > 0 Then s$ = s$ + ")"
rowlist_getsel = s$

End Function

Private Sub Text5_Change()
'd2infile = "splan": d2insub = "Text5_Change"
If noupd% = 1 Then Exit Sub
If Len(trm(Text5.text)) = 10 Then
  Call Command2_Click
  Call beglist_Change
End If

End Sub

Private Sub Text5_DblClick()

'd2infile = "splan": d2insub = "Text5_DblClick"
  With frmCalendar
    .init Text5, Text5.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text5.text = Format(.SelectedDate, "yyyy-mm-dd")
    End If
  End With
  Unload frmCalendar
End Sub

Private Sub Text6_Change()

'd2infile = "splan": d2insub = "Text6_Change"
form1.Combo1.text = Text6.text

End Sub
Sub abochk(halle$, block$, dtg$)
Dim rrr
Dim c$, r As ADODB.Recordset, raum$
Dim air%, aib%

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "abochk"
c$ = "select raum from hblist where hid='" & halle$ & "' and pgid='" & block$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
On Error GoTo exme_abochk
raum$ = r!raum
On Error GoTo 0
air% = form1.abosimraum(halle, raum$, dtg$)
aib% = form1.abosimblock(halle$, block$, dtg$)
c$ = "Aboplätze im Raum: " & air% & ", davon hier: " & aib%
splan.Caption = c$

exme_abochk:
On Error GoTo 0

End Sub
Public Function platzistverkauft(h$, p$, dtg$, platz, reihe)
Dim rrr
Dim c$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "splan": d2insub = "platzistverkauft"
platzistverkauft = 0
c$ = "SELECT hbpstatus.id, hbplist.hid, hbplist.pgid, hbplist.platzname, hbpstatus.dtg, hbpstatus.pstatus " + _
   "FROM hbpstatus INNER JOIN hbplist ON hbpstatus.hbpid = hbplist.id " + _
   "WHERE (((hbplist.hid)='" & h$ & "') AND ((hbplist.pgid)='" & p$ & "') " + _
          "AND ((hbpstatus.dtg)='" & dtg$ & "') " + _
          "AND ((hbplist.platz)=" & platz & ") " + _
          "AND ((hbplist.reihe)=" & reihe & ") " + _
          ");"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, dbOpenDynaset, dbReadOnly, d2infile, d2insub)
While Not r.EOF
  If InStr(LCase(r!pstatus), "verkauft") > 0 Then
    platzistverkauft = 1
    Exit Function
  End If
  r.MoveNext
Wend

End Function
