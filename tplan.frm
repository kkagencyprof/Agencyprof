VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form tplan 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Projekte - AgencyProf"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   FillColor       =   &H00C0C0C0&
   Icon            =   "tplan.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView gd1 
      Height          =   4095
      Left            =   2760
      TabIndex        =   60
      Top             =   3600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bleistift löschen"
      Height          =   255
      Left            =   2400
      TabIndex        =   97
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton btnTopic 
      Caption         =   "Topic"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   96
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "P"
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   95
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10800
      TabIndex        =   94
      ToolTipText     =   "Zoom"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Index           =   2
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Text            =   "tplan.frx":000C
      Top             =   960
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   360
      TabIndex        =   93
      ToolTipText     =   "Zum Löschen deaktivieren"
      Top             =   3840
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tplan.frx":0012
      Style           =   1  'Grafisch
      TabIndex        =   92
      ToolTipText     =   "Löschen gesamten Datensatz"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
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
      Left            =   6600
      TabIndex        =   91
      ToolTipText     =   "Priorität, Änderungen werden sofort gespeichert."
      Top             =   300
      Width           =   375
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6480
      TabIndex        =   88
      ToolTipText     =   "Neue Kalkulation hinzufügen"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton delme 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1680
      Picture         =   "tplan.frx":0502
      Style           =   1  'Grafisch
      TabIndex        =   79
      ToolTipText     =   "Liste der aufgerufenen Projekte löschen. (Löscht NICHT das Projekt)"
      Top             =   540
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   77
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   3480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "tplan.frx":17D8
      Style           =   1  'Grafisch
      TabIndex        =   75
      ToolTipText     =   "Projektverzeichnis im Explorer öffnen"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox eintermin 
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
      Left            =   3120
      TabIndex        =   73
      ToolTipText     =   "erstellt nur einen Termin (falls ein Datum eingetragen ist)"
      Top             =   6120
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   480
   End
   Begin VB.ComboBox Text2 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "tplan.frx":1E02
      Left            =   480
      List            =   "tplan.frx":1E04
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
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
      Left            =   2400
      Picture         =   "tplan.frx":1E06
      Style           =   1  'Grafisch
      TabIndex        =   69
      ToolTipText     =   "Leere Termine erstellen"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton Command25 
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
      Left            =   120
      Picture         =   "tplan.frx":2198
      Style           =   1  'Grafisch
      TabIndex        =   68
      ToolTipText     =   "Neues Projekt anlegen"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tplan.frx":252A
      Style           =   1  'Grafisch
      TabIndex        =   67
      ToolTipText     =   "per Email an Agencyprof"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton Command23 
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
      Height          =   255
      Left            =   120
      TabIndex        =   66
      ToolTipText     =   "Hilfeseite öfnen"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
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
      Left            =   6840
      TabIndex        =   65
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6480
      TabIndex        =   64
      ToolTipText     =   "Neuen Beteiligten hinzufügen"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton fshow 
      BackColor       =   &H00C0C0C0&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   59
      Top             =   6840
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2880
      Picture         =   "tplan.frx":25DC
      Style           =   1  'Grafisch
      TabIndex        =   58
      ToolTipText     =   "Zeige mehr Details"
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton alarm 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tplan.frx":3756
      Style           =   1  'Grafisch
      TabIndex        =   57
      ToolTipText     =   "Benachrichtigung bei Änderung der Daten"
      Top             =   4800
      Width           =   375
   End
   Begin VB.ListBox List4 
      Height          =   1755
      IntegralHeight  =   0   'False
      Left            =   720
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   56
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   55
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command20 
      Caption         =   "&drucken"
      Height          =   255
      Left            =   2400
      TabIndex        =   54
      Top             =   7440
      Width           =   1335
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   0
      Top             =   5640
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Programme drucken"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   53
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      Picture         =   "tplan.frx":3B98
      Style           =   1  'Grafisch
      TabIndex        =   52
      ToolTipText     =   "Wiedervorlage"
      Top             =   4080
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "tplan.frx":3F17
      Left            =   3600
      List            =   "tplan.frx":3F19
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alle anzeigen"
      Height          =   255
      Left            =   720
      TabIndex        =   51
      Top             =   7440
      Width           =   1575
   End
   Begin VB.ListBox chgs 
      Height          =   255
      Left            =   720
      TabIndex        =   50
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tplan.frx":3F1B
      Style           =   1  'Grafisch
      TabIndex        =   49
      ToolTipText     =   "Speichern"
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   3480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   5400
      TabIndex        =   46
      Top             =   8040
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leere löschen"
      Height          =   255
      Left            =   2400
      TabIndex        =   45
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Neuer Auftritt"
      Height          =   255
      Left            =   7320
      TabIndex        =   44
      Top             =   8040
      Width           =   1455
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   43
      Top             =   5640
      Width           =   7335
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   5040
      TabIndex        =   42
      Top             =   3600
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   12
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "tplan.frx":3F7B
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "löschen"
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3480
      TabIndex        =   39
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Programme"
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   36
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   35
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   34
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
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
      Left            =   6000
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "löschen"
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Neu"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   1485
      Index           =   11
      Left            =   720
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   2  'Automatisch
      TabIndex        =   5
      Text            =   "tplan.frx":3F81
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   8025
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   3480
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3480
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   915
      Index           =   6
      Left            =   8760
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "tplan.frx":3F87
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   9840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Picture         =   "tplan.frx":3F8D
      Style           =   1  'Grafisch
      TabIndex        =   15
      Top             =   7320
      Width           =   375
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
   Begin MSComctlLib.ListView gd2 
      Height          =   1215
      Left            =   6480
      TabIndex        =   62
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command27 
      Height          =   375
      Left            =   2760
      Picture         =   "tplan.frx":41DD
      Style           =   1  'Grafisch
      TabIndex        =   70
      ToolTipText     =   "Künstlerauftritte erstellen"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton Command28 
      Height          =   375
      Left            =   3120
      Picture         =   "tplan.frx":459A
      Style           =   1  'Grafisch
      TabIndex        =   71
      ToolTipText     =   "Orchesterauftritte erstellen"
      Top             =   5760
      Width           =   375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   80
      Text            =   "Combo3"
      Top             =   480
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   81
      Text            =   "Combo3"
      Top             =   840
      Width           =   2550
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   82
      Text            =   "Combo3"
      Top             =   1200
      Width           =   2550
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   3
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   83
      Text            =   "Combo3"
      Top             =   1560
      Width           =   2550
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   4
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   84
      Text            =   "Combo3"
      Top             =   1920
      Width           =   2550
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   5
      IntegralHeight  =   0   'False
      Left            =   3480
      TabIndex        =   85
      Text            =   "Combo3"
      Top             =   2280
      Width           =   2550
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8280
      Picture         =   "tplan.frx":460B
      Style           =   1  'Grafisch
      TabIndex        =   89
      ToolTipText     =   "Kalkulationen neu berechnen"
      Top             =   2880
      Width           =   375
   End
   Begin VB.ListBox kalklist 
      Height          =   885
      IntegralHeight  =   0   'False
      Left            =   6480
      TabIndex        =   86
      ToolTipText     =   "Kalkulation öffnen oder neu aus einer Vorlage erstellen"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CheckBox sh_wd 
      Height          =   255
      Left            =   5640
      TabIndex        =   98
      Top             =   5400
      Width           =   135
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10920
      TabIndex        =   99
      ToolTipText     =   "zoom"
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
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
      Left            =   10680
      TabIndex        =   100
      ToolTipText     =   "zoom"
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox sh_lnk 
      Height          =   255
      Left            =   9360
      TabIndex        =   101
      Top             =   5400
      Value           =   1  'Aktiviert
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "linked events"
      Height          =   255
      Left            =   9600
      TabIndex        =   102
      Top             =   5460
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   120
      Picture         =   "tplan.frx":4C7D
      ToolTipText     =   "Kontakt löschen verboten"
      Top             =   3600
      Width           =   315
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prio.:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   6600
      TabIndex        =   90
      Top             =   120
      Width           =   495
   End
   Begin VB.Label gd1stat 
      Alignment       =   1  'Rechts
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
      Height          =   255
      Left            =   2760
      TabIndex        =   61
      Top             =   3360
      Width           =   8415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Kalkulationen"
      Height          =   255
      Left            =   7440
      TabIndex        =   87
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   2160
      TabIndex        =   78
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   2280
      TabIndex        =   76
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Erstellen:"
      Height          =   255
      Left            =   2400
      TabIndex        =   72
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "weitere Beteiligte"
      Height          =   255
      Left            =   7320
      TabIndex        =   63
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   120
      Picture         =   "tplan.frx":51A1
      ToolTipText     =   "Doppelklick zum Umbenennen"
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   13
      Left            =   2160
      TabIndex        =   47
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   12
      Left            =   2160
      TabIndex        =   41
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   960
      TabIndex        =   30
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   28
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   27
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   24
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   8760
      TabIndex        =   23
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   22
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6480
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3015
      Left            =   2040
      Shape           =   4  'Gerundetes Rechteck
      Top             =   360
      Width           =   4335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   1935
      Left            =   600
      Shape           =   4  'Gerundetes Rechteck
      Top             =   3480
      Width           =   10815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ein Term."
      Height          =   255
      Left            =   2400
      TabIndex        =   74
      ToolTipText     =   "Doppelclick=heute"
      Top             =   6120
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   2295
      Left            =   600
      Shape           =   4  'Gerundetes Rechteck
      Top             =   5520
      Width           =   10815
   End
End
Attribute VB_Name = "tplan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nflds, prv$, neuprogid$
Dim tptyp$, prvc$, break%, l6h
Dim svc As Boolean
Dim memox0, memoy0, memow0, memoh0, c26nocreate As Boolean

Private Sub alarm_Click()
'd2infile = "tplan": d2insub = "alarm_Click"
Load alarmlist
Call alarmlist.settab("tplan")
alarmlist.Caption = transe("Projekt-ID:") + Text1(0).text

End Sub

Private Sub btnTopic_Click()
Dim tpid$

If form1.isfieldmissing("opt_topics", "id") Then Exit Sub
tpid$ = Text1(0).text
Load dochist2
DoEvents
dochist2.topics.Clear
dochist2.topics.AddItem tpid$
dochist2.topics.Selected(0) = True
DoEvents
'Call dochist2.topics_Click

End Sub

Private Sub Check1_Click()
'd2infile = "tplan": d2insub = "Check1_Click"
If Check1.value = 1 Then
  Command4.Enabled = True
Else
  Command4.Enabled = False
End If

End Sub


Private Sub Check2_Click()
If Check2.value = 1 Then
  Command32.Visible = False
Else
  Command32.Visible = True
End If

End Sub

Private Sub Combo1_Click()
'd2infile = "tplan": d2insub = "Combo1_Click"
If prvc$ <> transo(Combo1.text) Then
  BackColor = form1.dirtycolor()
  Command17.Enabled = True
  Me.Caption = transe(Combo1.text + " - Projekte")
End If


End Sub

Private Sub Combo1_GotFocus()
'd2infile = "tplan": d2insub = "Combo1_GotFocus"
prvc$ = transo(Combo1.text)
End Sub

Public Sub combo1_LostFocus()
Dim tpid$, hp$, rtmp As ADODB.Recordset
Dim nid$, own$, wert$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "combo1_LostFocus"
Call setcaption(trm(Combo1.text) + " - " + transe("Projekt"))
tpid$ = Text1(0).text
If tpid$ = "" Then Exit Sub
hp$ = transo(trm(Combo1.text))
If trm(hp$) = "" Then
  hp$ = transe("Künstler")
  Combo1.text = hp$
  If prvc$ <> "" And hp$ <> "" Then
    Combo1.text = prvc$
    hp$ = transo(prvc$)
  End If
End If
If LCase(hp$) <> "künstler" And LCase(hp$) <> "orchester" And LCase(hp$) <> "kammermusik" And LCase(hp$) <> "crossover" Then
    c$ = "SELECT * FROM sysvars where instr(owner,'sysvar_system_projekttyp')=1 and LCASE(wert)='" + LCase(hp$) + "';"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If rtmp.EOF Then
        ask% = MsgBox(transe("Neuen Projekttyp anlegen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Neuer Projekttyp?"))
        If ask% <> vbYes Then
          Call Combo1.SetFocus
          Exit Sub
        End If
        nid$ = form1.newid("sysvars", "id", 38)
        own$ = "sysvar_system_projekttyp_auto_" + hp$
        wert$ = hp$
        c$ = "insert into sysvars (id,owner,wert) values('" + _
           nid$ + "','" + _
           own$ + "','" + _
           wert$ + "');"
        Call form1.sqlqry(c$)
    End If
End If
If prvc$ <> transo(Combo1.text) Then
  BackColor = form1.dirtycolor()
  Command17.Enabled = True
  s$ = "update tplan set Hauptperson='" & hp$ & "' where id='" + tpid$ & "'"
  chgs.AddItem s$
End If

End Sub

Private Sub Combo3_Change(Index As Integer)
Dim trgt%, trgf$

'd2infile = "tplan": d2insub = "Combo3_Change"
Select Case Index
    Case 0:
        trgt% = 14: trgf$ = "Projektbetreuer"
    Case 1:
        trgt% = 3: trgf$ = "Orchester"
    Case 2:
        trgt% = 13: trgf$ = "Solist"
    Case 3:
        trgt% = 9: trgf$ = "Dirigent"
    Case 4:
        trgt% = 8: trgf$ = "Veranstalter"
    Case 5:
        trgt% = 1: trgf$ = "Tourneeleitung"
    Case Else
        Exit Sub
End Select
Call Text1_GotFocus(trgt%)
Text1(trgt%).text = Combo3(Index).text: DoEvents
Call Text1_LostFocus(trgt%)
End Sub

Private Sub Combo3_Click(Index As Integer)

'd2infile = "tplan": d2insub = "Combo3_Click"
Call Combo3_Change(Index)

End Sub

Private Sub Combo3_DropDown(Index As Integer)
Dim trgt%, trgf$, cmd$, r As ADODB.Recordset, perg$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Combo3_DropDown"
Select Case Index
    Case 0:
        trgt% = 14: trgf$ = "Projektbetreuer"
    Case 1:
        trgt% = 3: trgf$ = "Orchester"
    Case 2:
        trgt% = 13: trgf$ = "Solist"
    Case 3:
        trgt% = 9: trgf$ = "Dirigent"
    Case 4:
        trgt% = 8: trgf$ = "Veranstalter"
    Case 5:
        trgt% = 1: trgf$ = "Tourneeleitung"
    Case Else
        Exit Sub
End Select
Combo3(Index).Clear
hp$ = "": If trm(transo(Combo1.text)) <> "" Then hp$ = "and Hauptperson='" & transo(Combo1.text) & "'"
cmd$ = "select " & trgf$ & " as ergwert from tplan where " & trgf$ & "<>'' and not isnull(" & trgf$ & ") " & hp$ & " order by " & trgf$ & ";"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
perg$ = ""
While Not r.EOF
  If LCase(perg$) <> LCase(trm(r!ergwert)) Then
    perg$ = trm(r!ergwert)
    Combo3(Index).AddItem perg$
  End If
  r.MoveNext
Wend

End Sub

Private Sub Command1_Click()


'd2infile = "tplan": d2insub = "Command1_Click"
Hide
Unload tabkalk: DoEvents
Unload tplan
Unload fdet

End Sub

Private Sub Command10_Click()
Dim cmd$, rtmp As QueryDef, up$, tpid$, ProgID$

'd2infile = "tplan": d2insub = "Command10_Click"
tpid$ = Text1(0).text
If tpid$ = "" Then Exit Sub
ProgID$ = ""
Load prog
Call prog.SetFocus
Call prog.callbackinit("tplan", "")
While neuprogid$ = "": DoEvents: Wend

If neuprogid$ = "" Or neuprogid$ = "_LOGOUT_" Then Exit Sub
List3.AddItem neuprogid$
cmd$ = "insert into tpprogli (id,tpid,prgid,_desc) values('" + _
                  form1.newid("tpprogli", "id", 18) & "','" + _
                  tpid$ & "','" + _
                  neuprogid$ & "'," & _
                  trm(10 * (chgs.ListCount + 1)) & ")"
Call form1.sqlqry(cmd$)
neuprogid$ = ""
Call nulldsp
Call showrec(tpid$)


End Sub

Private Sub Command11_Click()

'd2infile = "tplan": d2insub = "Command11_Click"
If List3.ListIndex < 0 Then Exit Sub
id$ = List3.List(List3.ListIndex)
If InStr(id$, "(ID:") = 0 Then Exit Sub

List3.RemoveItem List3.ListIndex
id$ = Mid$(id$, InStr(id$, "(ID:") + 4)
Call form1.sqlqry("delete from tpprogli where id='" & id$ & "'")
'Call List1_Click

End Sub


Private Sub Command12_Click()
'd2infile = "tplan": d2insub = "Command12_Click"
If trm(Text1(14).text) = "" Then
  Call Label1_DblClick(14)
Else
  Call openadr(Text1(14).text)
End If

End Sub

Private Sub Command13_Click()
Dim nid$
'd2infile = "tplan": d2insub = "Command13_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub

nid$ = form1.newid("auftritt", "id", 20)
form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung) VALUES ('" + nid$ & "','" + tpid$ & "','Neuer Auftritt','" + tpid$ & "')")
Call rlist6(tpid$)

End Sub

Private Sub Command14_Click()
Dim r As ADODB.Recordset
Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command14_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
MousePointer = 11
DoEvents
c$ = "SELECT id FROM auftritt where TourneeplanID ='" + tpid$ & "' and Auftrittstyp='Neuer Auftritt'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  form1.sqlqry ("delete from finanzen where id='" & r!id & "'")
  r.MoveNext
Wend

form1.sqlqry ("delete from auftritt where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Neuer Auftritt'")
Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub Command15_Click()
'd2infile = "tplan": d2insub = "Command15_Click"
If trm(Text1(13).text) = "" Then
  Call Label1_DblClick(13)
Else
  Call openadr(Text1(13).text)
End If
End Sub

Public Sub Command16_Click()
'd2infile = "tplan": d2insub = "Command16_Click"
Load tpzoom
Call tpzoom.setmode(1)
On Error Resume Next
tpzoom.SetFocus
On Error GoTo 0

End Sub

Public Sub Command17_Click()
'd2infile = "tplan": d2insub = "Command17_Click"
If trm(Text1(5).text) = "" And trm(Text1(4).text) <> "" Then
  Call Text1_GotFocus(5)
  DoEvents
  Text1(5).text = Text1(4).text
  Call Text1_LostFocus(5)
  DoEvents
End If
DoEvents
If trm(Text1(4).text) <> "" And trm(Text1(5).text) <> "" Then Call checktournee
For i% = 0 To chgs.ListCount - 1
  form1.sqlqry (chgs.List(i%))
Next i%
chgs.Clear
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
Call tpltchk(tpid$, "Tourneeleitung", Text1(1).text)
Call tpltchk(tpid$, "Projektbetreuer", Text1(14).text)
Call tpltchk(tpid$, "Orchester", Text1(3).text)
Call tpltchk(tpid$, "Solist", Text1(13).text)
Call tpltchk(tpid$, "Dirigent", Text1(9).text)
Call tpltchk(tpid$, "Veranstalter", Text1(8).text)
Call tpltchk(tpid$, "mehr_Solisten", Text1(12).text)
BackColor = form1.cleancolor()
Command17.Enabled = False
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub Command18_Click()

'd2infile = "tplan": d2insub = "Command18_Click"
Load create2do
Call create2do.initmsg(form1.getuserid(), form1.getuserid(), Text1(0).text & " [Wiedervorlage] Projekt:" + _
               Text1(0).text, "", Date, Left(Time, 5))
Call create2do.SetFocus
create2do.Text1(1).Enabled = False
create2do.Text1(3).Enabled = False

End Sub

Private Sub Command2_Click()
'd2infile = "tplan": d2insub = "Command2_Click"
For i% = 0 To List4.ListCount - 1
  List4.Selected(i%) = False
Next i%
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
  Call rlist6(tpid$)
  Call rgd1
End Sub

Private Sub Command20_Click()
Dim df$, tp0$

'd2infile = "tplan": d2insub = "Command20_Click"
Call savecheck
df$ = Combo2.text
If df$ = "" Then Exit Sub

tp0$ = trm(Text1(0).text): If tp0$ = "" Then Exit Sub

Call List6_DblClick
DoEvents
For i% = 0 To auftritt.List1.ListCount - 1
  If df$ = auftritt.List1.List(i%) Then
    auftritt.List1.ListIndex = i%
    DoEvents
    Call auftritt.List1_DblClick
    DoEvents
    Unload auftritt
    If List6.ListIndex >= 0 Then List6.ListIndex = 0
    Call gotorec(tp0$)
    Exit Sub
  End If
Next i%

End Sub


Private Sub Command21_Click()

'd2infile = "tplan": d2insub = "Command21_Click"
  tpid$ = Text1(0).text: If trm(tpid$) = "" Then Exit Sub
  Load adrselect
  Call adrselect.sel_init("", transe("Person"))
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then
    wer$ = adrselect.kontselid$
    If trm(wer$) = "" Or wer$ = "-1" Then wer$ = adrselect.sel_getselected()
    c$ = "insert into tpwernoch (id,tpid,kid) values('" & _
      form1.newid("tpwernoch", "id", 40) & "','" & _
      tpid$ & "','" & _
      wer$ & "')"
    Call form1.sqlqry(c$)
    Call rgd2
  End If
  Unload adrselect
End Sub

Private Sub Command22_Click()
Dim rrr
'd2infile = "tplan": d2insub = "Command22_Click"
  Set lvitem = gd2.SelectedItem
  On Error Resume Next
  id$ = lvitem.SubItems(2)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  
  c$ = "delete from tpwernoch where id='" & id$ & "'"
  Call form1.sqlqry(c$)
  Call rgd2

End Sub

Private Sub Command23_Click()
'd2infile = "tplan": d2insub = "Command23_Click"
Call form1.handbuchcall("12-Projekte.htm")

End Sub

Private Sub Command24_Click()
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, r1 As ADODB.Recordset, stmp As ADODB.Recordset, werk As ADODB.Recordset, komp As ADODB.Recordset
Dim tpid$, c$, tb$, o%, tg$, p%, tg0$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command24_Click"
Call Command17_Click
tpid$ = trm(Text1(0).text)
MousePointer = 11: DoEvents
c$ = "select * from tplan where id='" & tpid$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  On Error Resume Next
  Kill form1.mydatadir() & "\*.sql"
  On Error GoTo 0
  o% = FreeFile
  tg$ = form1.mydatadir() & "\" & strrepl(tpid$, " ", "_") & ".sql"
  tg0$ = tg$
  Open tg$ For Output As #o%
  c$ = "insert into tplan (id) values('" & r.Fields(0).value & "');"
  Print #o%, c$
  For i% = 1 To form1.sqla.TableDefs("tplan").Fields.Count - 1
    If trm(r.Fields(i%).value) <> "" Then
      c$ = form1.mkupdcmd("tplan", "id", tpid$, r.Fields(i%).name, r.Fields(i%).Type, r.Fields(i%).value) & ";"
      Print #o%, c$
    End If
  Next i%
  If trm(r!tourneeleitung) <> "" Then Call form1.sqlex_adresse("adresse", "id", r!tourneeleitung)
  If trm(r!orchester) <> "" Then Call form1.sqlex_adresse("adresse", "id", r!orchester)
  If trm(r!veranstalter) <> "" Then Call form1.sqlex_adresse("adresse", "id", r!veranstalter)
  If trm(r!dirigent) <> "" Then Call form1.sqlex_adresse("adresse", "id", r!dirigent)
  tb$ = "tpwernoch"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM " & tb$ & " where tpid ='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    c$ = "insert into " & tb$ & " (id) values('" & r.Fields(0).value & "');"
    Print #o%, c$
    For i% = 1 To form1.sqla.TableDefs(tb$).Fields.Count - 1
      If trm(r.Fields(i%).value) <> "" Then
        c$ = form1.mkupdcmd(tb$, "id", r.Fields(0).value, r.Fields(i%).name, r.Fields(i%).Type, r.Fields(i%).value) & ";"
        Print #o%, c$
      End If
    Next i%
    c$ = "select id from adresse where id='" & r!kid & "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r1.EOF Then
      Call form1.sqlex_adresse("adresse", "id", r!kid)
    End If
    r.MoveNext
  Wend
  tb$ = "tpprogli"
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM " & tb$ & " where tpid ='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not r.EOF
    c$ = "insert into " & tb$ & " (id) values('" & r.Fields(0).value & "');"
    Print #o%, c$
    For i% = 1 To r.Fields.Count - 1
      If trm(r.Fields(i%).value) <> "" Then
        c$ = form1.mkupdcmd(tb$, "id", r.Fields(0).value, r.Fields(i%).name, r.Fields(i%).Type, r.Fields(i%).value) & ";"
        Print #o%, c$
      End If
    Next i%
    If Not IsNull(r!prgid) Then
      Call form1.sqlex_adresse("programm", "programmid", r!prgid)
      Set stmp = New ADODB.Recordset
      stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT * FROM programmliste where programmid ='" + r!prgid & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not stmp.EOF
        c$ = "insert into programmliste (id) values('" & stmp!id & "');"
        Print #o%, c$
        For i% = 1 To stmp.Fields.Count - 1
          If trm(stmp.Fields(i%).value) <> "" Then
            c$ = form1.mkupdcmd("programmliste", "id", stmp.Fields(0).value, stmp.Fields(i%).name, stmp.Fields(i%).Type, stmp.Fields(i%).value) & ";"
            Print #o%, c$
          End If
        Next i%
        Set werk = New ADODB.Recordset
        werk.CursorLocation = adUseServer
rrr = form1.adoopen(werk, "SELECT * FROM w_loc where id ='" + stmp!werkid & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not werk.EOF Then
          c$ = "insert into w_loc (id) values('" & werk!id & "');"
          Print #o%, c$
          For i% = 1 To werk.Fields.Count - 1
            If trm(werk.Fields(i%).value) <> "" Then
              c$ = form1.mkupdcmd("w_loc", "id", werk.Fields(0).value, werk.Fields(i%).name, werk.Fields(i%).Type, werk.Fields(i%).value) & ";"
              Print #o%, c$
            End If
          Next i%
          Set komp = New ADODB.Recordset
          komp.CursorLocation = adUseServer
rrr = form1.adoopen(komp, "SELECT * FROM k_loc where id ='" + werk!KomponistenNummer & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If Not komp.EOF Then
            c$ = "insert into k_loc (id) values('" & komp!id & "');"
            Print #o%, c$
            For i% = 1 To komp.Fields.Count - 1
              If trm(komp.Fields(i%).value) <> "" Then
                c$ = form1.mkupdcmd("k_loc", "id", komp.Fields(0).value, komp.Fields(i%).name, komp.Fields(i%).Type, komp.Fields(i%).value) & ";"
                Print #o%, c$
              End If
            Next i%
          End If
          komp.MoveNext
        End If
        stmp.MoveNext
      Wend
    End If
    r.MoveNext
  Wend
  tb$ = "auftritt"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM " & tb$ & " where tourneeplanid ='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  While Not rtmp.EOF
    If trm(rtmp!auftrittstyp) <> "" Then
    If rtmp!auftrittstyp <> "Neuer Auftritt" Then
    tb$ = "auftritt"
    Call form1.sqlex_adresse("auftritt", "id", rtmp!id)
    Call form1.sqlex_adresse("finanzen", "id", rtmp!id)
    Call form1.sqlex_adresse("usr_" & utabn(rtmp!auftrittstyp), "id", rtmp!id)
    tb$ = "auftritthigru"
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT * FROM auftritthigru where auftrittsid ='" + rtmp.Fields(0).value & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      c$ = "insert into auftritthigru (id) values('" & stmp.Fields(0).value & "');"
      Print #o%, c$
      If Len(stmp!felddaten) < 80 Then
        c$ = "select id from adresse where id='" & stmp!felddaten & "'"
        Set r1 = New ADODB.Recordset
        r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not r1.EOF Then
          Call form1.sqlex_adresse("adresse", "id", stmp!felddaten)
        End If
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
           LCase(stmp.Fields(i%).name) <> "auftrittstyp" And _
           LCase(stmp.Fields(i%).name) <> "astatus" Then
          c$ = form1.mkupdcmd(tb$, "id", stmp.Fields(0).value, stmp.Fields(i%).name, stmp.Fields(i%).Type, stmp.Fields(i%).value) & ";"
          Print #o%, c$
        End If
        End If
      Next i%
      stmp.MoveNext
    Wend
    End If
    End If
    rtmp.MoveNext
  Wend
  Load smtp
  On Error Resume Next
  Call smtp.SetFocus
  On Error GoTo 0
  tg$ = Dir(form1.mydatadir() & "\*.sql")
  While tg$ <> ""
    If form1.mydatadir() & "\" & tg$ <> tg0$ Then
    On Error Resume Next
    p% = FreeFile
    Open form1.mydatadir() & "\" & tg$ For Input As #p%
    rrr = Err
    On Error GoTo 0
    If rrr = 0 Then
      While Not EOF(p%)
        Line Input #p%, l$
        Print #o%, l$
      Wend
      Close #p%
      On Error Resume Next
      Kill form1.mydatadir() & "\" & tg$
      On Error GoTo 0
    End If
    End If
    tg$ = Dir
  Wend
  Close #o%
  smtp.txtMessageSubject = "Agencyprof Datenpakete Tourneeplan " & Text1(0).text
  smtp.txtMessageText = "Speichern Sie die Attachments in Ihrem Agencyprof-Verzeichnis"
  tg$ = Dir(form1.mydatadir() & "\*.sql")
  While tg$ <> ""
    Call smtp.attachfile(form1.mydatadir() & "\" & tg$)
    tg$ = Dir
  Wend
End If
MousePointer = 0

End Sub

Private Sub Command25_Click()
Dim neuid$

'd2infile = "tplan": d2insub = "Command25_Click"
Call savecheck
Call nulldsp
Me.BackColor = form1.cleancolor()
neuid$ = trm(InputBox(transe("Neue Projekt-ID:"), transe("Neues Projekt anlegen"), ""))
neuid$ = strrepl(trm(neuid$), "/", "_")
neuid$ = strrepl(neuid$, "'", "´")
neuid$ = strrepl(neuid$, "&", "_")
If neuid$ <> "" Then
  Call Command3_Click
  Text1(0).text = neuid$
  Call Text1_LostFocus(0)
End If

End Sub

Public Sub Command26_Click()
Dim r As ADODB.Recordset, r1 As ADODB.Recordset
Dim nid$, immeranlegen As Boolean, anlegen As Boolean

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command26_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
If trm(Text1(4).text) = "" Then
  MsgBox transe("Anfangsdatum des Projekts fehlt.")
  Exit Sub
End If
MousePointer = 11
On Error Resume Next
d0 = CDate(datum2sql(Text1(4).text))
rrr = Err
On Error GoTo 0
If rrr <> 0 Then d0 = CDate(Date)
d1 = d0
If Text1(5).text <> "" Then
  On Error Resume Next
  d1 = CDate(Text1(5).text)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then d1 = CDate(Date)
Else
  Text1(5).text = Text1(4).text
End If
immeranlegen = True
If form1.getusersetting("ProjektLeereTermineImmer") <> "ja" Then immeranlegen = False
diff = d1 - d0
cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Tournee'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  nid$ = form1.newid("auftritt", "id", 20)
  Call form1.sqlqry("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Tournee','" + tpid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
  Call tpltchk(tpid$, "Tourneeleitung", Text1(1).text)
  Call tpltchk(tpid$, "Projektbetreuer", Text1(14).text)
  Call tpltchk(tpid$, "Orchester", Text1(3).text)
  Call tpltchk(tpid$, "Solist", Text1(13).text)
  Call tpltchk(tpid$, "Dirigent", Text1(9).text)
  Call tpltchk(tpid$, "Veranstalter", Text1(8).text)
  Call tpltchk(tpid$, "mehr_Solisten", Text1(12).text)
Else
  cmd$ = "update auftritt set datum='" + datum2sql(CDate(d0)) & "' where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Tournee'"
  Call form1.sqlqry(cmd$)
End If

Dt = d0
If eintermin.text <> "" Then
  d1 = CDate(datum2sql(eintermin.text))
  immeranlegen = True
  Dt = d1
End If
If Not c26nocreate Then
  While Dt <= d1
    anlegen = True
    If Not immeranlegen Then
      cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(Dt)) & "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then anlegen = False
    End If
    If anlegen Then
      nid$ = form1.newid("auftritt", "id", 20)
      form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Neuer Auftritt','" + tpid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
    End If
    Dt = CDate(Dt) + 1
  Wend
End If
Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub Command27_Click()
Dim r As ADODB.Recordset, r1 As ADODB.Recordset
Dim nid$, immeranlegen As Boolean, anlegen As Boolean

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command27_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
If trm(Text1(4).text) = "" Then Exit Sub

MousePointer = 11
d0 = CDate(datum2sql(Text1(4).text))
d1 = d0
If Text1(5).text <> "" Then
  d1 = CDate(datum2sql(Text1(5).text))
Else
  Text1(5).text = Text1(4).text
End If
diff = d1 - d0
immeranlegen = True
If form1.getusersetting("ProjektKünstlerTermineImmer") <> "ja" Then immeranlegen = False
cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(d0)) & "' and Auftrittstyp='Tournee'"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  nid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Tournee','" + tpid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
End If

Dt = d0
If eintermin.text <> "" Then
  d1 = CDate(datum2sql(eintermin.text))
  immeranlegen = True
  Dt = d1
End If
While Dt <= d1
  anlegen = True
  If Not immeranlegen Then
    cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(Dt)) & "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then anlegen = False
  End If
  If anlegen Then
    nid$ = form1.newid("auftritt", "id", 20)
    form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Künstlerauftritt','" + tpid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
  End If
  Dt = CDate(Dt) + 1
Wend

Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub Command28_Click()
Dim r As ADODB.Recordset, r1 As ADODB.Recordset
Dim nid$, immeranlegen As Boolean, anlegen As Boolean

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command28_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
If trm(Text1(4).text) = "" Then Exit Sub

MousePointer = 11
d0 = CDate(datum2sql(Text1(4).text))
d1 = d0
If Text1(5).text <> "" Then
  d1 = CDate(datum2sql(Text1(5).text))
Else
  Text1(5).text = Text1(4).text
End If
diff = d1 - d0
immeranlegen = True
If form1.getusersetting("ProjektOrchesterTermineImmer") <> "ja" Then immeranlegen = False
cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(d0)) & "' and Auftrittstyp='Tournee'"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  nid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Tournee','" + tpid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
End If

Dt = d0
If eintermin.text <> "" Then
  d1 = CDate(datum2sql(eintermin.text))
  immeranlegen = True
  Dt = d1
End If
While Dt <= d1
  anlegen = True
  If Not immeranlegen Then
    cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(Dt)) & "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then anlegen = False
  End If
  If anlegen Then
    nid$ = form1.newid("auftritt", "id", 20)
    form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Orchesterauftritt','" + tpid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
  End If
  Dt = CDate(Dt) + 1
Wend

Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub


Private Sub Command29_Click()
'd2infile = "tplan": d2insub = "Command29_Click"
On Error Resume Next
MkDir form1.s0dir() + "\" + form1.medien() + "\"
MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\"
MkDir form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + form1.medienname(Text1(0).text)
X = Shell("explorer.exe " + form1.s0dir() + "\" + form1.medien() + "\__PROJEKTE__\" + form1.medienname(Text1(0).text), vbNormalFocus)
On Error GoTo 0


End Sub

Public Sub Command3_Click()

'd2infile = "tplan": d2insub = "Command3_Click"
Call nulldsp
Text1(0).Enabled = True
Text1(0).text = ""
Text1(0).SetFocus

End Sub


Private Sub Command30_Click()
Dim neuid$

'd2infile = "tplan": d2insub = "Command30_Click"
neuid$ = trm(InputBox(transe("Bezeichnung:"), transe("Neue Kalkulation anlegen"), ""))
If neuid$ <> "" Then
  neuid$ = strrepl(neuid$, " ", "_")
  Unload tabkalk
  DoEvents
  Load tabkalk
  On Error Resume Next
  Call tabkalk.SetFocus
  On Error GoTo 0

  tabkalk.Label2.Caption = Text1(0).text
  tabkalk.Label3.Caption = neuid$
  tabkalk.Label1.Caption = transe("Projekt")
End If

End Sub

Private Sub Command31_Click()
Dim i%

'd2infile = "tplan": d2insub = "Command31_Click"
For i% = 0 To kalklist.ListCount - 1
  kalklist.ListIndex = i%: DoEvents
  Call kalklist_DblClick: DoEvents
  Call tabkalk.Command2_Click: DoEvents
  Call tabkalk.Command1_Click: DoEvents
Next i%
End Sub

Private Sub Command32_Click()
Dim id$
Command32.Visible = False
Check2.value = 1

id$ = Text1(0).text
Call delproj(id$)
End Sub

Private Sub Command33_Click()

If Command33.Caption = "+" Then
  Text1(2).Left = Shape2.Left
  Text1(2).Width = 9255
  Text1(2).Height = 4335
  Command33.Caption = "-"
Else
  Text1(2).Left = memox0
  Text1(2).Top = memoy0
  Text1(2).Width = memow0
  Text1(2).Height = memoh0
  Command33.Caption = "+"
End If
End Sub

Private Sub Command34_Click()
Dim r As ADODB.Recordset, r1 As ADODB.Recordset
Dim rtmp As ADODB.Recordset
Dim nid$, immeranlegen As Boolean, anlegen As Boolean, c$, crt$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command34_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
If trm(Text1(4).text) = "" Then Exit Sub

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM auftrittstypen where id='Zeitraum'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rtmp.EOF Then
  rtmp.Close
  rrr = form1.adoopen(rtmp, "SELECT id FROM auftrittstypen where id='Bleistift'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If Not rtmp.EOF Then
    rtmp.Close
    crt$ = "Bleistift"
  End If
Else
  crt$ = "Zeitraum"
End If
If crt$ = "" Then
  MsgBox transe("Das geht nicht. Weder der Typ Zeitraum, noch der Typ Bleistift sind vorhanden.")
  Exit Sub
End If
MousePointer = 11
On Error Resume Next
d0 = CDate(Text1(4).text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then d0 = CDate(datum2sql(Text1(4).text))
d1 = d0
If Text1(5).text <> "" Then
  On Error Resume Next
  d1 = CDate(Text1(5).text)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then d1 = CDate(datum2sql(Text1(5).text))
Else
  Text1(5).text = Text1(4).text
End If
diff = d1 - d0
immeranlegen = True
If form1.getusersetting("ProjektKünstlerTermineImmer") <> "ja" Then immeranlegen = False
cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(d0)) & "' and Auftrittstyp='Tournee'"

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  nid$ = form1.newid("auftritt", "id", 20)
  form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Tournee','" + tpid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
End If

Dt = d0
If eintermin.text <> "" Then
  On Error Resume Next
  d1 = CDate(eintermin.text)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then d1 = CDate(datum2sql(eintermin.text))
  immeranlegen = True
  Dt = d1
End If
While Dt <= d1
  anlegen = True
  If Not immeranlegen Then
    cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and datum='" + datum2sql(CDate(Dt)) & "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not r.EOF Then anlegen = False
  End If
  If anlegen Then
    nid$ = form1.newid("auftritt", "id", 20)
    form1.sqlqry ("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','" + crt$ + "','" + tpid$ & "','" + _
                 datum2sql(CDate(Dt)) & "')")
    If trm(Text1(13).text) <> "" Then
      c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) VALUES ('" + _
                form1.newid("auftritthigru", "id", 18) & "','" + nid$ + _
                "','" + crt$ + "','Künstler','" + trm(Text1(13).text) & "')"
      form1.sqlqry (c$)
      If LCase(crt$) = "zeitraum" Then
        c$ = "insert into usr_zeitraum (id,Künstler) values('" + nid$ + "','" + trm(Text1(13).text) + "')"
      Else
        c$ = "insert into usr_bleistift (id,Künstler) values('" + nid$ + "','" + trm(Text1(13).text) + "')"
      End If
      form1.sqlqry (c$)
    End If
  End If
  Dt = CDate(Dt) + 1
Wend

Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click

End Sub

Private Sub Command35_Click()
Dim r As ADODB.Recordset
Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command35_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
ask% = MsgBox(transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Alle Termine vom Typ Bleistift löschen?"))
If ask% <> vbYes Then Exit Sub
MousePointer = 11
DoEvents
c$ = "SELECT id FROM auftritt where TourneeplanID ='" + tpid$ & "' and Auftrittstyp='Bleistift'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  form1.sqlqry ("delete from finanzen where id='" & r!id & "'")
  form1.sqlqry ("delete from auftritthigru where auftrittsid='" & r!id & "'")
  r.MoveNext
Wend

form1.sqlqry ("delete from auftritt where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Bleistift'")
Call rlist6(tpid$)
MousePointer = 0
If form1.kalopen Then Call kc.Command1_Click
If form1.dayvopen Then Call dayvw.Command4_Click
If form1.priosopen Then Call prios.Command20_Click


End Sub

Private Sub Command36_Click()
List6.Font.Size = List6.Font.Size + 1
DoEvents
List6.Height = l6h
End Sub

Private Sub Command37_Click()
List6.Font.Size = List6.Font.Size - 1
DoEvents
List6.Height = l6h
End Sub

Private Sub Command4_Click()
Dim id$

'd2infile = "tplan": d2insub = "Command4_Click"
Command4.Enabled = False
Check1.value = 0
id$ = List1.List(List1.ListIndex)
If List1.ListIndex < 0 Then Exit Sub
If id$ = "" Then Exit Sub

antw = MsgBox("Tourneeplan " + id$ & " löschen?", vbYesNo + vbCritical + vbDefaultButton2, "Daten löschen?")
If antw = vbYes Then
  chgs.AddItem "delete from tplan where id='" + id$ & "'"
  BackColor = form1.dirtycolor()
  Command17.Enabled = True
End If
Call rlist1
End Sub


Private Sub Command5_Click()

'd2infile = "tplan": d2insub = "Command5_Click"
If trm(Text1(3).text) = "" Then
  Call Label1_DblClick(3)
Else
  Call openadr(Text1(3).text)
End If
End Sub

Private Sub Command6_Click()
'd2infile = "tplan": d2insub = "Command6_Click"
If trm(Text1(9).text) = "" Then
  Call Label1_DblClick(9)
Else
  Call openadr(Text1(9).text)
End If
End Sub

Private Sub Command7_Click()
'd2infile = "tplan": d2insub = "Command7_Click"
If trm(Text1(8).text) = "" Then
  Call Label1_DblClick(8)
Else
  Call openadr(Text1(8).text)
End If
End Sub

Private Sub Command8_Click()
'd2infile = "tplan": d2insub = "Command8_Click"
If trm(Text1(1).text) = "" Then
  Call Label1_DblClick(1)
Else
  Call openadr(Text1(1).text)
End If

End Sub

Private Sub Command9_Click()
'd2infile = "tplan": d2insub = "Command9_Click"
Load prog
prog.SetFocus

End Sub

Private Sub delme_Click()
'd2infile = "tplan": d2insub = "delme_Click"
On Error Resume Next
Kill form1.mydatadir() & "\" & form1.mkfn(tplan.Caption)
On Error GoTo 0
Text2.Clear

End Sub

Private Sub eintermin_DblClick()

'd2infile = "tplan": d2insub = "eintermin_DblClick"
  p$ = eintermin.text
  With frmCalendar
    .init eintermin, eintermin.text
    .Show vbModal, Me
    If (.SelectionOK) Then
      eintermin.text = Format(.SelectedDate, "dd.mm.yyyy")
    End If
  End With
  Unload frmCalendar

End Sub

Private Sub Form_Load()
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Form_Load"
axsResizer1.SaveControlPositions
svc = True
Call form1.dbg2f("tplan:load")
s% = form1.myfontsize()
List1.Font.Size = s%
List5.Font.Size = s%
List6.Font.Size = s%
List3.Font.Size = s%
c26nocreate = False
nflds = 15
For i% = 0 To nflds - 1: Text1(i%).Font.Size = s%: Next i%
gd1.View = lvwReport
Set colHeader = gd1.ColumnHeaders.add(, , transe("Datum"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Von"), 1200)
Set colHeader = gd1.ColumnHeaders.add(, , transe("An"), 1200)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Anzahl"), 800)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Netto"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Währ."), 600)
Set colHeader = gd1.ColumnHeaders.add(, , transe("ges. netto"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("MwSt"), 800)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Text"), 1600)
Set colHeader = gd1.ColumnHeaders.add(, , transe("Kurs"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("vom"), 1000)
Set colHeader = gd1.ColumnHeaders.add(, , transe("in €"), 1000)
fshow.Caption = transe("€")
gd1.Font.Size = s%
gd2.Font.Size = s%

List6.Visible = True
gd1.Visible = False
gd1stat.Visible = False
Label1(14).ForeColor = form1.lnkcolor
Label1(3).ForeColor = form1.lnkcolor
Label1(9).ForeColor = form1.lnkcolor
Label1(8).ForeColor = form1.lnkcolor
Label1(1).ForeColor = form1.lnkcolor
Label1(9).ForeColor = form1.lnkcolor
Label1(12).ForeColor = form1.lnkcolor
Label1(13).ForeColor = form1.lnkcolor

gd2.View = lvwReport
Set colHeader = gd2.ColumnHeaders.add(, , transe("Funktion"), 900)
Set colHeader = gd2.ColumnHeaders.add(, , transe("Kontakt"), 1200)
Set colHeader = gd2.ColumnHeaders.add(, , transe("ID"), 12)
Shape4.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Shape2.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Shape3.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Shape1.BackColor = form1.getusersetting("shapecolor", "12632256"): DoEvents
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
sh_wd = form1.getusersetting("tplan_showweekday", 1)

Call form1.formpos(Me)

tplan.Caption = transe("Projekte - AgencyProf")
Command30.ToolTipText = transe("Neue Kalkulation hinzufügen")
delme.ToolTipText = transe("Liste der aufgerufenen Projekte löschen. (Löscht NICHT das Projekt)")
Command29.ToolTipText = transe("Projektverzeichnis im Explorer öffnen")
eintermin.ToolTipText = transe("erstellt nur einen Termin (falls ein Datum eingetragen ist)")
Command26.ToolTipText = transe("Leere Termine erstellen")
Command25.ToolTipText = transe("Neues Projekt anlegen")
Command24.ToolTipText = transe("per Email an Agencyprof")
Command23.ToolTipText = transe("Hilfeseite öffnen")
Command21.ToolTipText = transe("Neuen Beteiligten hinzufügen")
fshow.Caption = transe("€")
Command16.ToolTipText = transe("Zeige mehr Details")
alarm.ToolTipText = transe("Benachrichtigung bei Änderung der Daten")
Command20.Caption = transe("&drucken")
Command19.Caption = transe("Programme drucken")
Command18.ToolTipText = transe("Wiedervorlage")
Command2.Caption = transe("Alle anzeigen")
Command17.ToolTipText = transe("Speichern")
Command14.Caption = transe("Leere löschen")
Command35.Caption = transe("Bleistift löschen")
Command13.Caption = transe("Neuer Auftritt")
Command11.Caption = transe("löschen")
Command9.Caption = transe("Programme")
Command4.Caption = transe("löschen")
Command3.Caption = transe("Neu")
Command27.ToolTipText = transe("Künstlerauftritte erstellen")
Command28.ToolTipText = transe("Orchesterauftritte erstellen")
Command31.ToolTipText = transe("Kalkulationen neu berechnen")
kalklist.ToolTipText = transe("Kalkulation öffnen oder neu aus einer Vorlage erstellen")
Label5.Caption = transe("Kalkulationen")
Label3.Caption = transe("Erstellen:")
Label2.Caption = transe("weitere Beteiligte")
Image1.ToolTipText = transe("Doppelklick zum Umbenennen")
Label4.Caption = transe("Ein Term.")
Label4.ToolTipText = transe("Doppelklick=heute")

Show

List4.Clear
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM auftrittstypen order by sortierung", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

Command34.Enabled = False
While (Not rtmp.EOF) And (Not Command34.Enabled)
  List4.AddItem transe(rtmp!id)
  If LCase(trm(rtmp!id)) = "zeitraum" Or LCase(trm(rtmp!id)) = "bleistift" Then
    Command34.Enabled = True
  End If
  rtmp.MoveNext
Wend
List4.ListIndex = -1

Command4.Enabled = False
Check1.value = 0
Call rlists
BackColor = form1.cleancolor()
Command17.Enabled = False
Call rcombo2
Call rkalklist
memox0 = Text1(2).Left
memoy0 = Text1(2).Top
memow0 = Text1(2).Width
memoh0 = Text1(2).Height
BackColor = form1.cleancolor()
l6h = List6.Height
End Sub
Public Sub rlist1()
Dim rtmp As ADODB.Recordset, i As Integer

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rlist1"
'On Error GoTo errhdl
Call nulldsp
BackColor = form1.cleancolor()
List1.Clear
List6.Clear

If trm(Text2.text) = "" Then Exit Sub
break% = 0
cmd$ = "SELECT id FROM tplan where "
cmd$ = cmd$ & "id like '%" & Text2.text & "%'"

'If tptyp$ = "künstler" Or tptyp$ = "orchester" Then cmd$ = cmd$ & " and hauptperson='" + tptyp$ & "'"
'If tptyp$ <> "" And tptyp$ <> "projekte" And tptyp <> "-" Then cmd$ = cmd$ & " and hauptperson='" + tptyp$ & "'"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While Not rtmp.EOF And break% = 0
  List1.AddItem rtmp!id
  rtmp.MoveNext
  If rtmp.EOF Then break% = 1
  DoEvents
Wend
On Error GoTo 0
If List1.ListCount > 0 Then List1.ListIndex = 0
BackColor = form1.cleancolor()

For i = 0 To List1.ListCount - 1
  If List1.List(i) = Text2.text Then
    List1.ListIndex = i
    Exit For
  End If
Next i
Exit Sub
errhdl:
  rrr = Err
  If rrr <> 0 Then
    If rrr <> 3420 Then MsgBox trm("Fehler #" & rrr & " " & Error$(rrr))
    On Error GoTo 0
    break% = 1
    Unload adrselect
    Exit Sub
  End If
  Resume Next
End Sub

Sub rlist6(tpid$)
Dim rtmp As ADODB.Recordset, stmp As ADODB.Recordset, taid$, zt$, wd$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rlist6"
List6.Clear

nosel = 1
For i% = 0 To List4.ListCount - 1
  If List4.Selected(i%) = True Then
    i% = List4.ListCount
    nosel = 0
  End If
Next i%
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT * FROM auftritt where tourneeplanid ='" + tpid$ & "' order by datum, zeit"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Exit Sub
End If
While Not rtmp.EOF
  rtmport = " "
  If Not IsNull(rtmp!ort) Then
    If trm(rtmp!ort) <> "" Then
      rtmport = " " & rtmp!ort & " "
    End If
  End If
  zt$ = onlynums(trm(rtmp!zeit)): If Len(zt$) > 4 Then zt$ = Left(zt$, 4)
  While Len(zt$) < 4: zt$ = "0" + zt$: Wend
  wd$ = " ": If sh_wd.value = 1 Then wd$ = " (" & form1.dayofweek(rtmp!datum) & ") "
  If nosel = 1 Then
'    If rtmp!auftrittstyp <> "Tourneetag" Then
      List6.AddItem rtmp!datum & " " & zt$ & wd$ & form1.get_atabkz(rtmp!auftrittstyp) & rtmport & "(" & rtmp!bezeichnung & ")" & Space$(60) & "(AID:" & rtmp!id
'    End If
  Else
    For i% = 0 To List4.ListCount - 1
      If rtmp!auftrittstyp = transo(List4.List(i%)) And List4.Selected(i%) = True Then
        List6.AddItem rtmp!datum & " " & zt$ & wd$ & transe(rtmp!auftrittstyp) & "(" & rtmp!bezeichnung & ")" & Space$(60) & "(AID:" & rtmp!id
        i% = List4.ListCount
      End If
    Next i%
  End If
  rtmp.MoveNext
Wend
If form1.isfieldmissing("opt_othertplans", "id") Then Exit Sub
If sh_lnk.value = 0 Then Exit Sub
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
c$ = "SELECT aid FROM opt_othertplans where tpid ='" + tpid$ & "'"
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Exit Sub
End If
While Not rtmp.EOF
  taid$ = trm(rtmp!aid)
  
Set stmp = New ADODB.Recordset
stmp.CursorLocation = adUseServer
c$ = "SELECT * FROM auftritt where id ='" + taid$ & "'"
rrr = form1.adoopen(stmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then
  Exit Sub
End If
While Not stmp.EOF
  rtmport = " "
  If Not IsNull(stmp!ort) Then
    If trm(stmp!ort) <> "" Then
      rtmport = " " & stmp!ort & " "
    End If
  End If
  zt$ = onlynums(trm(stmp!zeit)): If Len(zt$) > 4 Then zt$ = Left(zt$, 4)
  While Len(zt$) < 4: zt$ = "0" + zt$: Wend
  wd$ = " ": If sh_wd.value = 1 Then wd$ = " (" & form1.dayofweek(stmp!datum) & ") "
  If nosel = 1 Then
'    If rtmp!auftrittstyp <> "Tourneetag" Then
      List6.AddItem stmp!datum & " " & zt$ & wd$ & form1.get_atabkz(stmp!auftrittstyp) & rtmport & "(by Ref: " & stmp!bezeichnung & ") " & Space$(60) & "(AID:" & stmp!id
'    End If
  Else
    For i% = 0 To List4.ListCount - 1
      If stmp!auftrittstyp = transo(List4.List(i%)) And List4.Selected(i%) = True Then
        List6.AddItem stmp!datum & " " & zt$ & wd$ & transe(stmp!auftrittstyp) & "(by Ref: " & stmp!bezeichnung & ")" & Space$(60) & "(AID:" & stmp!id
        i% = List4.ListCount
      End If
    Next i%
  End If
  stmp.MoveNext
Wend
  
  rtmp.MoveNext
Wend
End Sub
Sub showrec(tpid$)
Dim rtmp As ADODB.Recordset, pr As ADODB.Recordset, stmp As ADODB.Recordset
Dim komp As ADODB.Recordset, werk As ADODB.Recordset, rvv$, wid$, sid$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "showrec"
Show

i% = 0

If tpid$ = "-1" Or tpid$ = "" Then Exit Sub
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM tplan where id ='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not rtmp.EOF Then
  For i% = 0 To nflds
    On Error Resume Next
    rvv$ = rtmp.Fields(i%).value
    rrr = Err
    On Error GoTo 0
    If rrr <> 0 Then rvv$ = ""
    If rvv$ <> "" Then
      If i% = 4 Or i% = 5 Then
        Text1(i%).text = datfromsql(rvv$)
      Else
        If i% = nflds Then
          Combo1.text = word1(transe(rvv$ + " - Projekte"))
        Else
          Text1(i%).text = transe(rvv$)
        End If
      End If
    End If
  Next i%
  If Not form1.isfieldmissing("opt_prios", "id") Then
    prio.Enabled = True
    prio.Visible = True
    cmd$ = "SELECT * FROM opt_prios where evnt= 'T:" & tpid$ & "' and userid='" + form1.getuserid() + "'"
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
    Label6.Visible = False
    prio.Visible = False
  End If
End If

If Not form1.isfieldmissing("opt_topics", "id") Then
  c$ = "select * from opt_topics where topicid='" + tpid$ + "'"
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
  If rrr = 0 Then
    If Not rtmp.EOF Then btnTopic.Enabled = True
  End If
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM tpprogli where tpid ='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not rtmp.EOF
  List3.AddItem rtmp!prgid + Space$(40) & "(ID:" & rtmp!id
  If Not IsNull(rtmp!prgid) Then
    List5.AddItem rtmp!prgid & ":"
    Set stmp = New ADODB.Recordset
    stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT * FROM programmliste where programmid ='" + rtmp!prgid & "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not stmp.EOF
      wid$ = trm(stmp!werkid): sid$ = ""
      If Left$(wid$, 4) = "SBZ:" Then
        sid$ = Mid$(wid$, 5)
        wid$ = form1.getsatzidbywerkid(sid$)
      End If
      Set werk = New ADODB.Recordset
      werk.CursorLocation = adUseServer
rrr = form1.adoopen(werk, "SELECT KomponistenNummer,name,dauer FROM w_loc where id ='" + wid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not werk.EOF Then
        Set komp = New ADODB.Recordset
        komp.CursorLocation = adUseServer
rrr = form1.adoopen(komp, "SELECT name,vornamen FROM k_loc where id ='" + werk!KomponistenNummer & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
        If Not komp.EOF Then
          If sid$ = "" Then
            List5.AddItem werk!Dauer & " Min., " & komp!name & ", " & komp!vornamen & ": " & werk!name
          Else
            List5.AddItem komp!name & ", " & komp!vornamen & ": " + form1.getsatznamebyid(sid$) + " " + transe("aus") + " " + werk!name
          End If
        End If
      End If
      stmp.MoveNext
    Wend
  End If
  rtmp.MoveNext
Wend
Call rlist6(tpid$)
Call rgd2
Call rkalklist
BackColor = form1.cleancolor()

End Sub

Sub nulldsp()
Dim i%, c$
Dim rtmp As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "nulldsp"
i% = 0

gd1.ListItems.Clear
gd1stat.Caption = ""
fshow.Caption = transe("€")
List6.Visible = True
gd1.Visible = False
btnTopic.Enabled = False
gd1stat.Visible = False
For i% = 0 To nflds
  On Error Resume Next
  Label1(i%).Caption = transe(form1.sqla.TableDefs("tplan").Fields(i%).name)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then Exit Sub
  If i% <> nflds Then
    Text1(i%).text = ""
  Else
    Combo1.text = ""
    Combo1.Clear
    Combo1.AddItem word1(transe("Orchester" + " - Projekte"))
    Combo1.AddItem word1(transe("Künstler" + " - Projekte"))
    Combo1.AddItem word1(transe("Kammermusik" + " - Projekte"))
    Combo1.AddItem word1(transe("Crossover" + " - Projekte"))
    c$ = "SELECT * FROM sysvars where instr(owner,'sysvar_system_projekttyp')=1;"
    Set rtmp = New ADODB.Recordset
    rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    While Not rtmp.EOF
      If trm(rtmp!wert) <> "" Then Combo1.AddItem transe(trm(rtmp!wert))
      rtmp.MoveNext
    Wend
  End If
Next i%
List3.Clear
If List4.ListCount > 0 Then List4.ListIndex = -1
List5.Clear
List6.Clear
'List1.ListIndex = -1
Command17.Enabled = False
Me.BackColor = form1.cleancolor

End Sub

Private Sub Form_Resize()
'd2infile = "tplan": d2insub = "Form_Resize"
axsResizer1.Resize
DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
'd2infile = "tplan": d2insub = "Form_Unload"
Unload fdet
Call savecheck
Unload tpzoom
Unload auftrittshintergrund
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setusersetting("tplan_showweekday", sh_wd.value)
exuld:
On Error GoTo 0
End Sub

Public Sub fshow_Click()

'd2infile = "tplan": d2insub = "fshow_Click"
If fshow.Caption = transe("€") Then
fshow.Caption = "<-"
List6.Visible = False
Command31.Visible = False
gd1.Visible = True
gd1stat.Visible = True
DoEvents
Call rgd1
Else
fshow.Caption = transe("€")
Unload fdet
List6.Visible = True
Command31.Visible = True
gd1.Visible = False
gd1stat.Visible = False
End If
End Sub


Private Sub gd1_DblClick()
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "gd1_DblClick"
id$ = gd1.SelectedItem
p% = InStr(id$, "(AID:"): If p% = 0 Then Exit Sub
id$ = Mid$(id$, p% + 5)
Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(id$, 0)
Load fdet
Call fdet.SetFocus
fdet.fid = id$

End Sub


Private Sub gd2_AfterLabelEdit(Cancel As Integer, NewString As String)

'd2infile = "tplan": d2insub = "gd2_AfterLabelEdit"
Set lvitem = gd2.SelectedItem
id$ = lvitem.SubItems(2)
c$ = "update tpwernoch set funktion='" & trm(NewString) & "' where id='" & id$ & "'"
Call form1.sqlqry(c$)
Call rgd2

End Sub

Private Sub gd2_Click()
Dim r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "gd2_Click"
Set lvitem = gd2.SelectedItem

On Error Resume Next
frm$ = lvitem.SubItems(1)
rrr = Err
On Error GoTo 0
If rrr = 0 Then form1.Combo1.text = transe(frm$)

End Sub

Private Sub gd2_DblClick()
Dim r As ADODB.Recordset
Dim id$, aid$, tid$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "gd2_DblClick"
Set lvitem = gd2.SelectedItem
On Error Resume Next
id$ = lvitem.SubItems(2)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub

'MsgBox "ID=" & id$
c$ = "select kid from tpwernoch where id='" & id$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If Not r.EOF Then
  tid$ = trm(r!kid)
  If tid$ <> "" Then
    aid = form1.getadridbykontaktid(tid$)
    If aid$ = "" Or aid$ = "-1" Then aid$ = tid$
    Call shwAdrDetail.savecheck
    Call shwAdrDetail.refreshadrdetail(aid$, "")
    shwAdrDetail.Combo3.text = aid$
    Call shwAdrDetail.SetFocus
    shwAdrDetail.srchit% = 1
  Else
    Load adrselect
    Call adrselect.sel_init("", transe("Person"))
    Call adrselect.SetFocus
    Do
      DoEvents
    Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
    If adrselect.sel_brk() = 0 Then
      wer$ = adrselect.kontselid$
      If trm(wer$) = "" Or wer$ = "-1" Then wer$ = adrselect.sel_getselected()
      c$ = "update tpwernoch set kid='" + wer$ + "' where id='" + id$ + "'"
      Call form1.sqlqry(c$)
      Call rgd2
    End If
    Unload adrselect
  End If
End If
End Sub

Private Sub Image1_DblClick()
Dim r As ADODB.Recordset, l$, c$, neuid As String, altid As String

'd2infile = "tplan": d2insub = "Image1_DblClick"
Index = 0
If trm(Text1(0).text) <> "" Then
  altid = trm(Text1(0).text)
  neuid = InputBox(transe(transe("Projekt umbenennen")), Text1(0).text, Text1(0).text)
  If trm(neuid) = "" Then Exit Sub
  form1.sqlqry ("update tplan set id='" + neuid & "' where id='" + Text1(0).text & "'")
  form1.sqlqry ("update auftritt set TourneeplanID='" + neuid & "' where TourneeplanID='" + Text1(0).text & "'")
  form1.sqlqry ("update tpprogli set TpID='" + neuid & "' where TpID='" + Text1(0).text & "'")
  form1.sqlqry ("update tpwernoch set TpID='" + neuid & "' where TpID='" + Text1(0).text & "'")
  
  If Not form1.isfieldmissing("opt_topics", "id") Then
    c$ = "update opt_topics set topicid='" & neuid & "' where topicid='" & altid & "'"
    Call form1.sqlqry(c$)
    c$ = "select id,owner,wert from sysvars where owner like 'sysvar_system_tlnk_" + altid + "_%'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    While Not r.EOF
      l$ = strrepl(trm(r!Owner), altid, neuid)
'Debug.Print r!Owner; " ("; r!wert; ")"; vbCrLf; l$
      c$ = "update sysvars set owner='" + l$ + "' where id='" + trm(r!id) + "'"
      Call form1.sqlqry(c$)
      r.MoveNext
    Wend
  End If
  
  Text2.text = neuid
End If

End Sub

Private Sub kalklist_DblClick()
Dim tb0$, i%, kn$

'd2infile = "tplan": d2insub = "kalklist_DblClick"
i% = kalklist.ListIndex
If i% < 0 Then Exit Sub

Unload tabkalk
DoEvents
Load tabkalk
On Error Resume Next
Call tabkalk.SetFocus
On Error GoTo 0
kn$ = kalklist.List(i%)
i% = InStr(kn$, "("): If i% > 0 Then kn$ = trm(Left$(kn$, i% - 1))
i% = InStr(kn$, "="): If i% > 0 Then kn$ = Left$(kn$, i% - 1)
tabkalk.Label2.Caption = Text1(0).text
tabkalk.Label3.Caption = kn$
tabkalk.Label1.Caption = transe("Projekt")

End Sub

Private Sub Label1_DblClick(Index As Integer)
Dim neuwert As String, neukwert As String, s$

'd2infile = "tplan": d2insub = "Label1_DblClick"
Select Case Index
  Case 1: s$ = transe("Tourneeleitung")
  Case 9: s$ = transe("Dirigent")
  Case 8: s$ = transe("Veranstalter")
  Case 3: s$ = transe("Orchester|orch")
  Case 12: s$ = transe("Künstler")
  Case 13: s$ = transe("Künstler")
  Case 14: s$ = transe("Person")
  Case Default: s$ = ""
End Select
If s$ <> "" Then
  Load adrselect
  Call adrselect.sel_init("", s$)
  Call adrselect.SetFocus
  Do
    DoEvents
  Loop Until adrselect.sel_valid() = 1 Or adrselect.sel_brk() = 1
  If adrselect.sel_brk() = 0 Then
    Call Text1(Index).SetFocus
    DoEvents
    If Index = 12 And Len(Text1(Index).text) > 0 Then
      Text1(Index).text = Text1(Index).text + Chr$(13) + Chr$(10) + adrselect.sel_getselected()
    Else
      neukwert = adrselect.get_kontsel()
      neuwert = adrselect.sel_getselected(): neuawert = neuwert
      If neukwert <> "" Then neuwert = neukwert & " {" & neuwert & "}"
      Text1(Index).text = neuwert
    End If
  End If
  Unload adrselect
End If

End Sub

Private Sub Label4_DblClick()
'd2infile = "tplan": d2insub = "Label4_DblClick"
eintermin.text = Date
End Sub

Private Sub Label7_Click()
If sh_lnk.value = 0 Then
  sh_lnk.value = 1
Else
  sh_lnk.value = 0
End If
Call rlist6(Text1(0).text)
End Sub

Private Sub List1_Click()

'd2infile = "tplan": d2insub = "List1_Click"
Call savecheck
svc = False
Command4.Enabled = False
Check1.value = 0

tplan.MousePointer = vbHourglass
tpid$ = List1.List(List1.ListIndex)
Call nulldsp
Call showrec(tpid$)
BackColor = form1.cleancolor()
Command17.Enabled = False
tplan.MousePointer = 0

If trm(Text1(4).text) <> "" And trm(Text1(5).text) <> "" Then Call checktournee
If List6.ListCount > 0 Then List6.ListIndex = 0
svc = True

End Sub

Private Sub List1_DblClick()

'd2infile = "tplan": d2insub = "List1_DblClick"
Call Command16_Click

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx%, id$, sq$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "List1_KeyDown"
If KeyCode = 8 Or KeyCode = 46 Then
  idx% = List1.ListIndex
  If idx% < 0 Then Exit Sub
  id$ = List1.List(idx%)
  Call delproj(id$)
End If

End Sub

Private Sub List3_Click()

'd2infile = "tplan": d2insub = "List3_Click"
id$ = List3.List(List3.ListIndex)
If InStr(id$, "(ID:") = 0 Then Exit Sub
id$ = trm(Left$(id$, InStr(id$, "(ID:") - 1))
For i% = 0 To List5.ListCount - 1
  If Left$(List5.List(i%), Len(id$)) = id$ Then
    If i% + 8 > List5.ListCount - 1 Then
      j% = List5.ListCount - 1
    Else
      j% = i% + 7
    End If
    List5.ListIndex = j%
    List5.ListIndex = i%
    i% = List5.ListCount
  End If
Next i%


End Sub

Private Sub List3_DblClick()

'd2infile = "tplan": d2insub = "List3_DblClick"
Load prog
Call prog.SetFocus
id$ = List3.List(List3.ListIndex)
If InStr(id$, "(ID:") = 0 Then Exit Sub

id$ = trm(Left$(id$, InStr(id$, "(ID:") - 1))

Call prog.selectone(id$)

End Sub

Private Sub List4_Click()
'd2infile = "tplan": d2insub = "List4_Click"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
Call rlist6(tpid$)
If gd1.Visible = True Then Call rgd1

End Sub



Private Sub List6_Click()
'd2infile = "tplan": d2insub = "List6_Click"
Call rcombo2
End Sub

Public Sub List6_DblClick()

'd2infile = "tplan": d2insub = "List6_DblClick"
id$ = List6.List(List6.ListIndex)
id$ = Mid$(id$, InStr(id$, "(AID:") + 5)
Unload auftritt
DoEvents
Load auftritt
Call auftritt.SetFocus
Call auftritt.showrec(id$, 0)

End Sub

Private Sub List6_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sq$, atyp$, aid$, i%

If KeyCode = 8 Or KeyCode <> 46 Then Exit Sub
If List6.ListIndex < 0 Then Exit Sub
aid$ = List6.List(List6.ListIndex)
If InStr(aid$, "(AID:") > 0 Then
  aid$ = trm(Mid$(aid$, InStr(aid$, "(AID:") + 5))
  atyp$ = LCase$(form1.auftrittstyp(aid$))
  If LCase(atyp$) <> "zeitraum" Then
    If form1.getusersetting("tplandelete_" + LCase(atyp$), "") <> "erlaubt" Then
      Exit Sub
    End If
  End If
  sq$ = "delete from auftritthigru where auftrittsid='" + trm(aid$) & "'": Call form1.sqlqry(sq$)
  sq$ = "delete from auftritt where id='" + trm(aid$) & "'": Call form1.sqlqry(sq$)
  i% = List6.ListIndex
  List6.RemoveItem List6.ListIndex
  If i% >= List6.ListCount Then i% = List6.ListCount - 1
  List6.ListIndex = i%
  If form1.kalopen Then Call kc.Command1_Click
  If form1.dayvopen Then Call dayvw.Command4_Click
  If form1.priosopen Then Call prios.Command20_Click
End If
End Sub

Private Sub prio_Change()
Dim c As String, id, p As String
'd2infile = "tplan": d2insub = "prio_Change"
p = UCase(prio.text)
If p < "A" And p <> "" Then p = "A"
If p > "Z" And p <> "" Then p = "Z"
prio.text = p
id = trmx1(Text1(0).text)
If id <> "" Then
  c = "delete from opt_prios where userid='" + form1.getuserid() + "' and evnt='T:" + id + "';"
  Call form1.sqlqry(c)
  If p <> "" Then
    nid = form1.newid("opt_prios", "id", 36)
    c = "insert into opt_prios (id,evnt,userid,prio) values('" + _
        nid + "','T:" + _
        id + "','" + _
        form1.getuserid() + "','" + _
         p + "');"
    Call form1.sqlqry(c)
  End If
  If form1.priosopen Then Call prios.Command20_Click
End If
End Sub

Private Sub sh_lnk_Click()
DoEvents
Call rlist6(Text1(0).text)
End Sub

Private Sub sh_wd_Click()
tpid$ = Text1(0).text
Call rlist6(tpid$)
End Sub

Private Sub Text1_Change(Index As Integer)
'd2infile = "tplan": d2insub = "Text1_Change"
BackColor = form1.dirtycolor()
Command17.Enabled = True
Timer1.Interval = 15000
Timer1.Enabled = True

End Sub

Private Sub Text1_DblClick(Index As Integer)
Dim neuid As String, p$

'd2infile = "tplan": d2insub = "Text1_DblClick"
If Index = 2 Or Index = 11 Or Index = 6 Then
  Load memoview
  Call memoview.settext(Text1(Index).text)
  Exit Sub
End If

If Index = 4 Or Index = 5 Then
  If Text1(5).text = "" And Index = 5 Then Text1(5).text = Text1(4).text
  p$ = Text1(Index).text
  With frmCalendar
    .init Text1(Index), Text1(Index).text
    .Show vbModal, Me
    If (.SelectionOK) Then
      Text1(Index).text = Format(.SelectedDate, "dd.mm.yyyy")
      Call Text1_LostFocus(Index)
    End If
  End With
  Unload frmCalendar
  prv$ = p$
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)

'd2infile = "tplan": d2insub = "Text1_GotFocus"
prv$ = Text1(Index).text

End Sub

Public Sub Text1_LostFocus(Index As Integer)
Dim s$, r As ADODB.Recordset, chg2$, c$, c1$, idw$, k$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Text1_LostFocus"
If Index = 0 Then
  If Text1(0).Enabled = True Then
    Call form1.dbg2f("Neues Projekt")
    Text1(0).Enabled = False
    neuid$ = Text1(0).text
    If trm(neuid$) <> "" Then
      For i% = 0 To List1.ListCount - 1
        If List1.List(i%) = neuid$ Then
          List1.ListIndex = i%
          Exit Sub
        End If
      Next i%
      Call nulldsp
      Text1(0).text = neuid$
      Text1(10).text = Left$(neuid$, 4)
      List1.AddItem neuid$
      s$ = "insert into tplan (id,kuerzel,hauptperson) values('" + neuid$ & "','" + Text1(10).text & "','" + tptyp$ & "')"
      chgs.AddItem s$
      BackColor = form1.dirtycolor()
      Command17.Enabled = True
    End If
  End If

  For i% = 0 To List1.ListCount - 1
    If List1.List(i%) = neuid$ Then
      List1.ListIndex = i%
      i% = List1.ListCount + 1
    End If
  Next i%
  DoEvents
Else

BackColor = form1.dirtycolor()
Command17.Enabled = True
chg2$ = ""
id$ = Text1(0).text
If id$ = "" Then
  Text1(Index).text = prv$
  BackColor = form1.cleancolor()
  Command17.Enabled = False
  Exit Sub
End If
nwert$ = trm(Text1(Index).text)
If nwert$ <> prv$ Then
  fld$ = transo(Label1(Index).Caption)
  If nwert$ = "" Then
    nwert$ = "NULL"
  Else
    If Index = 4 Or Index = 5 Then
      nwert$ = datum2sql(nwert$)
      nwert$ = "'" + nwert$ & "'"
      c$ = "select id from taliste where id='" + id$ & "'"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        ask% = vbYes
        If form1.getusersetting("AngebotProjektDatumPrüfen") = "nein" Then ask% = vbNo
        If ask% = vbYes Then
          chg2$ = "update taliste set " + fld$ & "=" + nwert$ & " where id='" + id$ & "'"
        End If
      End If
    Else
      nwert$ = "'" + nwert$ & "'"
    End If
  End If
  If Index = 2 Then
    Call form1.sqlqry("update tplan set " + fld$ & "=" + nwert$ & " where id='" + id$ & "'")
  Else
    chgs.AddItem "update tplan set " + fld$ & "=" + nwert$ & " where id='" + id$ & "'"
    If chg2$ <> "" Then chgs.AddItem chg2$
  End If
End If

End If 'index=0

End Sub
Public Sub gotorec(tpid$)

'd2infile = "tplan": d2insub = "gotorec"
Text2.text = tpid$
break% = 1
Call rlist1

End Sub
Public Sub callback(prgid$)

neuprogid$ = prgid$

End Sub
Public Sub setcaption(c$)

'd2infile = "tplan": d2insub = "setcaption"
Call t2liste_sv(tplan.Caption)
tplan.Caption = c$
Call t2liste_ld(tplan.Caption)
tptyp$ = LCase(word1(transo(c$)))

End Sub
Public Sub rlists()
'd2infile = "tplan": d2insub = "rlists"
tptyp$ = LCase(word1(tplan.Caption))

Call rlist1
BackColor = form1.cleancolor()
Command17.Enabled = False

End Sub

Sub savecheck()
'd2infile = "tplan": d2insub = "savecheck"
If Not svc Then Exit Sub

If BackColor = form1.dirtycolor() Then
  If form1.immerspeichern() = "ja" Then
    antw = vbYes
  Else
    antw = MsgBox(transe("Sie haben Daten geändert, möchten Sie speichern?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Änderungen speichern?"))
  End If
  If antw = vbYes Then
    Call Command17_Click
  End If
End If
BackColor = form1.cleancolor()
End Sub

Private Sub Text2_Change()
'd2infile = "tplan": d2insub = "Text2_Change"
Timer1.Enabled = False
break% = 1
Call rlist1
Timer1.Interval = 15000
Timer1.Enabled = True
End Sub
Sub checktournee()
Dim r As ADODB.Recordset


Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "checktournee"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
DoEvents
If List6.ListCount <= 0 Then
  c26nocreate = True
  Call Command26_Click
  Call Command14_Click
  c26nocreate = False
Else
  Call C26a
End If
End Sub


Private Sub Command19_Click()
Dim o%, p%, nam$, vorlage$, ort$, dat$, ueb$, tpid$
Dim rtmp As ADODB.Recordset, stmp As ADODB.Recordset, r As ADODB.Recordset, s As ADODB.Recordset
Dim bkmstart$, bkmend$, prver%, udat As ADODB.Recordset, rev$, ttest$, wid$, sid$

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "Command19_Click"
bkmstart$ = "{\*\bkmkstart "
bkmend$ = "{\*\bkmkend "
Select Case LCase(form1.getsystemsetting("AusdruckAlleProgramme"))
  Case "version2": prver% = 2
  Case Else: prver% = 1
End Select
'v$ = "prgdrucknohead.rtf"
If V$ <> "" Then vorlage$ = V$
If vorlage$ = "" Then vorlage$ = form1.meineprgdruckvorlage()
If exist(form1.s0dir() & "\" + form1.getdbname() & ".rtf\" & vorlage$) = 0 Then
  MsgBox transe("Vorlage unbekannt:") + " " & vorlage$
  Exit Sub
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM programm where programmid='" + id$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
o% = FreeFile
Open form1.s0dir() & "\" & form1.getdbname() & ".rtf\" & vorlage$ For Input As #o%
p% = FreeFile
fn$ = form1.myuniquedocname("")
If fn$ = "" Then
  Close #p%
  Exit Sub
End If
MousePointer = 11
DoEvents
Set udat = New ADODB.Recordset
udat.CursorLocation = adUseServer
rrr = form1.adoopen(udat, "SELECT * FROM benutzerdaten where id ='" & form1.getuserid() & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
Open fn$ For Output As #p%
While Not EOF(o%)
  Line Input #o%, l$
  q% = InStr(l$, "PRGDRUCK")
  If q% > 0 Then
    For licnt% = 0 To List3.ListCount - 1

      id$ = List3.List(licnt%)
      p0% = InStr(id$, "(ID:")
      If p0% > 0 Then id$ = trm(Left$(id$, p0% - 1))
      Print #p%, " \par " & " \par Programm: " & id$

      If form1.getusersetting("ProgMitAuftritt") = "ja" Then
      tpid$ = trm(Text1(0).text)
      If tpid$ = "" Then Exit Sub
      cmd$ = "SELECT auftritthigru.auftrittsid, auftritt.id, auftritt.TourneeplanID " & _
             "FROM auftritthigru INNER JOIN auftritt ON auftritthigru.auftrittsid = auftritt.id " & _
             "WHERE (((auftritthigru.FeldDaten)='" & id$ & "') AND ((auftritthigru.FeldName)='Programm') AND ((auftritt.TourneeplanID)='" & tpid$ & "')) order by auftritt.datum;"
      Set r = New ADODB.Recordset
      r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      If Not r.EOF Then
        While Not r.EOF
          cmd$ = "select ort,datum from auftritt where id='" & r!auftrittsid & "'"
          Set s = New ADODB.Recordset
          s.CursorLocation = adUseServer
rrr = form1.adoopen(s, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          If Not s.EOF Then
            Print #p%, ", " & s!ort & ", " & datfromsql(s!datum)
          End If
          r.MoveNext
        Wend
      End If
      End If

      Print #p%, "\par "
      Print #p%, "\par "
      Set rtmp = New ADODB.Recordset
      rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT werkid FROM programmliste where programmid='" + id$ & "' order by position", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
      While Not rtmp.EOF
        wid$ = trm(rtmp!werkid): sid$ = ""
        If Left$(wid$, 4) = "SBZ:" Then
          sid$ = Mid$(wid$, 5)
          wid$ = form1.getsatzidbywerkid(sid$)
        End If
        k$ = form1.getkompvornamenamebywerkid(wid$)
        d$ = "(" + form1.getkompdatesbywerkid(wid$) & ")"
        If Left$(LCase$(k$), 7) = "pause p" Or Left$(LCase$(k$), 7) = "oder od" Then
          k$ = ""
          d$ = ""
        End If
        If prver% = 1 Then
          Print #p%, "\trowd \trgaph70\trleft-70 \cellx2410\cellx9476 \pard \intbl "
          Print #p%, "{"; form1.repl1310rtf("" & k$ & ""); "\par "; d$; "\cell "
          If sid$ = "" Then
            Print #p%, form1.repl1310rtf(form1.getwerknamebyid(wid$)); " ";
            dau$ = form1.repl1310rtf(form1.getdauerbywerkid("" & rtmp!werkid & "")): If sid$ <> "" Then dau$ = ""
            If trm(dau$) <> "" Then Print #p%, "(" & dau$ & " Min.) "
            Set stmp = New ADODB.Recordset
            stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT satzbezeichnung FROM sbz_loc where wid='" + rtmp!werkid & "' order by satznummer", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
            While Not stmp.EOF
              If InStr(stmp!satzbezeichnung, "Noten:") <> 1 Then
                Print #p%, "\par "; form1.repl1310rtf("" & stmp!satzbezeichnung & "")
              End If
              stmp.MoveNext
            Wend
          Else
            Print #p%, form1.repl1310rtf(form1.getsatznamebyid(sid$) + " " + transe("aus") + " " + form1.getwerknamebyid(wid$)); " ";
          End If
          Print #p%, "\par \cell \pard \intbl \row }\pard"
        End If
        If prver% = 2 Then
          Print #p%, form1.repl1310rtf("" & k$ & ""); " "; d$; " \tab "
          Print #p%, form1.repl1310rtf(form1.getwerknamebyid("" & rtmp!werkid & "")); " ";
          dau$ = form1.repl1310rtf(form1.getdauerbywerkid("" & rtmp!werkid & ""))
          If trm(dau$) <> "" Then Print #p%, "(" & dau$ & " " + transe("Min.") + ") "
'          Print #p%, "\trowd \trgaph70\trleft-70 \cellx4010\cellx7876 \pard \intbl "
'          Print #p%, "{"; form1.repl1310rtf("" & k$ & ""); " "; d$; "\cell "
'          Print #p%, form1.repl1310rtf(form1.getwerknamebyid("" & rtmp!werkid & "")); " (";
'          Print #p%, form1.repl1310rtf(form1.getdauerbywerkid("" & rtmp!werkid & "")); " Min.)\par \par "
'          Print #p%, "\cell \pard \intbl \row }\pard"
          Set stmp = New ADODB.Recordset
          stmp.CursorLocation = adUseServer
rrr = form1.adoopen(stmp, "SELECT satzbezeichnung FROM sbz_loc where wid='" + rtmp!werkid & "' order by satznummer", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
          While Not stmp.EOF
            sbz$ = stmp!satzbezeichnung
            If InStr(sbz$, "Noten:") <> 1 Then
              While InStr(sbz$, "  ") > 0
                sbz$ = strrepl(sbz$, "  ", " ")
              Wend
              Print #p%, form1.repl1310rtf("" & sbz$ & ""); " "
            End If
            stmp.MoveNext
          Wend
          Print #p%, "\par "
         End If
        rtmp.MoveNext
      Wend
    Next licnt%
  Else
    While Len(l$) > 0
      q% = InStr(l$, bkmstart$)
      If q% > 0 Then
        t$ = Mid$(l$, q% + Len(bkmstart$))
        Print #p%, Left$(l$, q% - 1)
        t$ = LCase(Left$(t$, InStr(t$, "}") - 1))
        If InStr(t$, "__") > 0 Then
          rev$ = Mid$(t$, InStr(t$, "__") + 2)
          ttest$ = Left$(t$, InStr(t$, "__") - 1)
          If ttest$ = "user" Then
            If Not udat.EOF Then
              For i% = 0 To 21 ' see einstellungen.load
                If Len(udat.Fields(i%).name) = Len(rev$) - 1 Then   'aliase ermitteln
                  If isdigit(Right$(rev$, 1)) <> 0 Then rev$ = Left$(rev$, Len(rev$) - 1)
                End If
                If LCase(udat.Fields(i%).name) = LCase(rev$) Then
                  Print #p%, strrepl(udat.Fields(i%).value, "\", "\\");
                  i% = 21
                End If
              Next i%
            End If
          End If
          If ttest$ = "system" Then
            Select Case LCase(rev$)
              Case "datum": Print #p%, Date
              Case "zeit": Print #p%, Left(Time, 5)
              Case "mwst": Print #p%, fixeurnozerotail(form1.sys_mwst / 100)
              Case Else: Print #p%, form1.getsystemsetting(rev$)
            End Select
          End If
        Else

          Select Case t$
            Case "von": Print #p%, apday(Text1(4).text) & "." & apmonth(Text1(4).text) & "."
            Case "bis": Print #p%, Text1(5).text
            Case "dirigent": Print #p%, form1.getnamebyid(Text1(9).text)
            Case "solist": k$ = form1.getnamebyid(Text1(13).text)
                         Print #p%, k$;
                         k$ = form1.instrumentvon(k$)
                         If k$ <> "" Then Print #p%, ", "; k$;
                         Print #p%, " "
            Case "mehr_solisten": Print #p%, form1.repl1310rtf(Text1(12).text)
            Case "orchester": Print #p%, form1.getnamebyid(Text1(3).text)
            Case Else
          End Select

        End If
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

MousePointer = 0
Call form1.openthisdoc(fn$, "")

End Sub


Sub rcombo2()
Dim aid$, r As ADODB.Recordset

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rcombo2"
Combo2.Clear

If List6.ListIndex < 0 Then Exit Sub
aid$ = List6.List(List6.ListIndex)
If InStr(aid$, "(AID:") > 0 Then
  aid$ = trm(Mid$(aid$, InStr(aid$, "(AID:") + 5))
  atyp$ = LCase$(form1.auftrittstyp(aid$))
  If atyp$ = "" Then Exit Sub
  tr$ = Dir(form1.vorlagendir & "\" + atyp$ & "*.rtf")
  rrr = Err
  On Error GoTo 0
  While tr$ <> "" And rrr = 0
    If atyp$ = Left$(tr$, InStr(LCase(tr$), ".") - 1) Then
      Combo2.AddItem basename(atyp$, ".rtf")
    Else
      Combo2.AddItem basename(Mid$(tr$, InStr(tr$, "_") + 1), ".rtf")
    End If
    tr$ = Dir
  Wend
End If
End Sub

Sub rgd1()
Dim r As ADODB.Recordset, c$, lvlitem As ListItem
Dim s As ADODB.Recordset, danz As Double, dnet As Double, d1 As Double, mwst As Double, kurs As Double
Dim sn As Double, sm As Double, sm0 As Double, sb As Double

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rgd1"
sn = 0
sm = 0
sb = 0
tpid$ = Text1(0).text
gd1.ListItems.Clear
nosel = 1
For i% = 0 To List4.ListCount - 1
  If List4.Selected(i%) = True Then
    i% = List4.ListCount
    nosel = 0
  End If
Next i%
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM auftritt where tourneeplanid ='" + tpid$ & "' order by datum,zeit", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  shw% = nosel
  If shw% = 0 Then
    For i% = 0 To List4.ListCount - 1
      If r!auftrittstyp = transo(List4.List(i%)) And List4.Selected(i%) = True Then
        shw% = 1
        i% = List4.ListCount
      End If
    Next i%
  End If
  If shw% = 1 Then
    mwst = 16
    c$ = "select * from finanzen where id='" & r!id & "'"
    Set s = New ADODB.Recordset
    s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
    If Not s.EOF Then
      mwst = s!mwst / 100
      von$ = "": If Not IsNull(s!von) Then von$ = s!von
      an$ = "": If Not IsNull(s!an) Then an$ = s!an
      net$ = "": If Not IsNull(s!netto) Then net$ = s!netto
      tut$ = "": If Not IsNull(s!bezeichnung) Then tut$ = s!bezeichnung
      anz$ = "0": If Not IsNull(s!anz) Then anz$ = s!anz
      wae$ = "": If Not IsNull(s!waehrung) Then wae$ = s!waehrung
      If Len(von$ + an$ + net$) > 0 Then
        danz = var2dbl(word1(strrepl(anz$, ".", "")))
        dnet = var2dbl(word1(strrepl(net$, ".", "")))
        d1 = danz * dnet
        If wae$ = "" Then wae$ = transe("€")
        If wae$ <> transe("€") Then
          kurs = var2dbl(strrepl(form1.kursvom(wae$, "" & datfromsql(r!datum)), ".", ","))
          If kurs = 0 Then kurs = 1
          sn = sn + d1 / kurs
        Else
          kurs = 1
          sn = sn + d1
        End If
        bru$ = fixeur(d1)
        sm0 = d1 * mwst / 100
        sm = sm + sm0 / kurs
        sb = sn + sm
        mws$ = fixeur(sm0)
        Set lvitem = gd1.ListItems.add(, , datfromsql(r!datum) & Space$(120) & "(AID:" & r!id)
        gd1stat.Caption = transe("Netto:") + " " & fixeur(sn) & " " + transe("MwSt:") + " " & fixeur(sm) & " " + transe("Brutto:") + " " & fixeur(sb)
        DoEvents
        On Error Resume Next
        lvitem.SubItems(1) = von$
        rrr = Err
        lvitem.SubItems(2) = an$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(3) = anz$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(4) = net$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(5) = wae$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(6) = bru$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(7) = "(" & mwst & "%) " & mws$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(8) = tut$
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(9) = d2db(kurs)
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(10) = form1.kursdatum(wae$, datfromsql(r!datum))
        If rrr <> 0 Then rrr = Err
        lvitem.SubItems(11) = fixeur(d1 / kurs)
        If rrr <> 0 Then rrr = Err
        On Error GoTo 0
        If rrr <> 0 Then Exit Sub
      End If
    End If
  End If
  r.MoveNext
Wend
gd1stat.Caption = transe("Netto:") + " " & fixeur(sn) & " " + transe("MwSt:") + " " & fixeur(sm) & " " + transe("Brutto:") + " " & fixeur(sb)

End Sub
Sub rgd2()
Dim r As ADODB.Recordset, c$, lvlitem As ListItem, i%

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rgd2"
tpid$ = Text1(0).text
gd2.ListItems.Clear

Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT * FROM tpwernoch where tpid='" + tpid$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  func$ = trm(r!funktion)
  wer$ = form1.get_kontaktname_by_id(trm(r!kid))
  If wer$ = "" Then
    wer$ = trm(r!kid)
  Else
    wer = wer$ + " {" + form1.getadridbykontaktid(r!kid) + "}"
  End If
  Set lvitem = gd2.ListItems.add(, , func$)
  lvitem.SubItems(1) = wer$
  lvitem.SubItems(2) = r!id
  r.MoveNext
Wend
      
      s$ = trm(strrepl(form1.getusersetting("tplan_alwaysadd", ""), " ", ""))
      While s$ <> ""
        c$ = cut_d1(s$, ","): s$ = cut_d2bis(s$, ",")
        idw$ = form1.newid("tpwernoch", "id", 16)
        k$ = cut_d2bis(c$, "(")
        If k$ <> "" Then
          c$ = cut_d1(c$, "(")
          k$ = cut_d1(k$, ")")
        End If
        For i% = 1 To gd2.ListItems.Count
          If gd2.ListItems(i%) = c$ Then Exit For
        Next i%
        If i% > gd2.ListItems.Count Then
          c2$ = "insert into tpwernoch (id,tpid,funktion,kid) values('" + idw$ + "','" + tpid$ + "','" + c$ + "','" + k$ + "')"
          Call form1.sqlqry(c2$)
          Set lvitem = gd2.ListItems.add(, , c$)
          lvitem.SubItems(1) = k$
          lvitem.SubItems(2) = k$
        End If
      Wend

End Sub


Private Sub Text2_Click()
'd2infile = "tplan": d2insub = "Text2_Click"
Call Text2_Change
End Sub

Private Sub Timer1_Timer()
'd2infile = "tplan": d2insub = "Timer1_Timer"
Timer1.Enabled = False
Call form1.dbg2f("tplan Timer1 start")
For i% = 0 To Text2.ListCount - 1
  If Text2.List(i%) = Text1(0).text Then
    i% = Text2.ListCount + 10
  End If
Next i%
If i% < Text2.ListCount + 5 Then
  Text2.AddItem Text1(0).text
  Call t2liste_sv(Me.Caption)
End If
Call form1.dbg2f("tplan Timer1 exit")
End Sub
Sub t2liste_sv(fin$)
Dim o%, fn$, i%

'd2infile = "tplan": d2insub = "t2liste_sv"
o% = FreeFile
fn$ = form1.mydatadir() & "\" & form1.mkfn(fin$)
If Text2.ListCount = 0 Then
  On Error Resume Next
  Kill fn$
  On Error GoTo 0
  Exit Sub
End If
If nexist(fn$) Then Exit Sub
Open fn$ For Output As #o%
For i% = 0 To Text2.ListCount - 1
  If trm(Text2.List(i)) <> "" Then Print #o%, Text2.List(i)
Next i%
Close #o%
End Sub
Sub t2liste_ld(fin$)
Dim o%, fn$, i%

'd2infile = "tplan": d2insub = "t2liste_ld"
o% = FreeFile
Text2.Clear
fn$ = form1.mydatadir() & "\" & form1.mkfn(fin$)
If exist(fn$) = 0 Then Exit Sub
Open fn$ For Input As #o%
While Not EOF(o%)
  Line Input #o%, l$: l$ = trm(l$)
  If l$ <> "" Then Text2.AddItem l$
Wend
Close #o%

End Sub

Sub openadr(sid$)
Dim bg, sida$, sidk$

'd2infile = "tplan": d2insub = "openadr"
bg = BackColor
Load shwAdrDetail
Call shwAdrDetail.savecheck
p% = InStr(sid$, "{")
sida$ = sid$: sidk$ = ""
If p% > 0 Then
  sidk$ = trm(Left(sid$, p% - 1))
  sida$ = trm(Mid(sid$, p% + 1)): sida$ = Left(sida$, Len(sida$) - 1)
End If
Call shwAdrDetail.refreshadrdetail(sida$, sidk$)
Call shwAdrDetail.SetFocus
BackColor = bg

End Sub

Public Sub rkalklist()
Dim tr, r As ADODB.Recordset, typ$, tb0$, cmd$, tb0l%, i%, prvn$, kn$, j%

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "rkalklist"
kalklist.Clear
typ$ = "Projekt"
prvn$ = "-------"
tb0$ = "tabkalk_" & typ$ & "_": tb0l% = Len(tb0$)
cmd$ = "select * from auftritthigru where auftrittsid='" & Text1(0).text & "' and instr(auftrittstyp,'" & tb0$ & "')=1 order by auftrittstyp asc, feldname desc"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
While Not r.EOF
  If prvn$ <> trm(r!auftrittstyp) Then
    prvn$ = trm(r!auftrittstyp)
    kalklist.AddItem Mid$(r!auftrittstyp, tb0l% + 1) & "=" & trm(r!felddaten)
  End If
  r.MoveNext
Wend
tb0l% = Len("tabkalk_projekt__")
tr = Dir(form1.s0dir() & "\" + form1.getdbname() & ".rtf\tabkalk_projekt__*.tbl")
While tr <> ""
  tb0$ = basename(Mid$(tr, tb0l% + 1), ".tbl")
  For i% = 0 To kalklist.ListCount - 1
    kn$ = kalklist.List(i%)
    j% = InStr(kn$, "="): If j% > 0 Then kn$ = Left$(kn$, j% - 1)
    If kn$ = tb0$ Then
      i% = -10
      Exit For
    End If
  Next i%
  If i% >= 0 Then kalklist.AddItem tb0$ & " " + transe("(Vorlage)")
  tr = Dir
Wend

End Sub

Sub tpltchk(id, fld, wrt)
Dim r As ADODB.Recordset, cmd$, nid$, aid$, w$

cmd$ = "select id from auftritt where Auftrittstyp='Tournee' and TourneeplanID='" + id + "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If r.EOF Then Exit Sub
aid$ = r!id

cmd$ = "select Feldname from auftrittsfelder where typ='Tournee' and (Feldname='" & fld & "' or instr(Feldname,'." & fld & ".')>0)"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If r.EOF Then Exit Sub

cmd$ = "select * from auftritthigru where Auftrittstyp='Tournee' and FeldName='" & fld & "' and auftrittsid='" & aid$ & "'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
If r.EOF Then
  If wrt <> "" Then
    nid$ = form1.newid("auftritthigru", "id", 8)
    cmd$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values('" + _
                 nid$ & "','" + aid$ & "','Tournee','" + fld + "','" + wrt + "')"
    Call form1.sqlqry(cmd$)
  End If
Else
  nid$ = r!id
  w$ = wrt
  If w$ = "" Then w$ = " "
  cmd$ = "update usr_tournee set " + fld + "='" + w$ + "' where id='" + aid$ + "'"
  Call form1.sqlqry(cmd$)
  If wrt = "" Then
    cmd$ = "update auftritthigru set Felddaten='" & wrt & "' where auftrittsid='" & aid$ & "' and auftrittstyp='Tournee' and id='" + nid$ + "'"
    Call form1.sqlqry(cmd$)
  Else
    cmd$ = "update auftritthigru set Felddaten='" + wrt + "' where auftrittsid='" + aid$ + "' and auftrittstyp='Tournee' and id='" + nid$ + "'"
    Call form1.sqlqry(cmd$)
  End If
End If
End Sub

Public Sub C26a()
Dim r As ADODB.Recordset, r1 As ADODB.Recordset
Dim nid$, immeranlegen As Boolean, anlegen As Boolean

Dim d2infile As String, d2insub As String
d2infile = "tplan": d2insub = "C26a"
tpid$ = trm(Text1(0).text)
If tpid$ = "" Then Exit Sub
If trm(Text1(4).text) = "" Then
  MsgBox transe("Anfangsdatum des Projekts fehlt.")
  Exit Sub
End If
MousePointer = 11
On Error Resume Next
d0 = trm(Text1(4).text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then d0 = CDate(Date)
d1 = d0
If Text1(5).text <> "" Then
  On Error Resume Next
  d1 = trm(Text1(5).text)
  rrr = Err
  On Error GoTo 0
  If rrr <> 0 Then d1 = CDate(Date)
Else
  Text1(5).text = Text1(4).text
End If
cmd$ = "select * from auftritt where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Tournee'"
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

If r.EOF Then
  nid$ = form1.newid("auftritt", "id", 20)
  Call form1.sqlqry("INSERT INTO auftritt (id, TourneeplanID,Auftrittstyp,bezeichnung,datum) VALUES ('" + _
                 nid$ & "','" + tpid$ + _
                 "','Tournee','" + tpid$ & "','" + _
                 datum2sql(CDate(d0)) & "')")
  Call tpltchk(tpid$, "Tourneeleitung", Text1(1).text)
  Call tpltchk(tpid$, "Projektbetreuer", Text1(14).text)
  Call tpltchk(tpid$, "Orchester", Text1(3).text)
  Call tpltchk(tpid$, "Solist", Text1(13).text)
  Call tpltchk(tpid$, "Dirigent", Text1(9).text)
  Call tpltchk(tpid$, "Veranstalter", Text1(8).text)
  Call tpltchk(tpid$, "mehr_Solisten", Text1(12).text)
  Call tpltchk(tpid$, "Enddatum", trm(d1))
Else
  cmd$ = "update auftritt set datum='" + datum2sql(CDate(d0)) & "' where TourneeplanID='" + tpid$ & "' and Auftrittstyp='Tournee'"
  Call form1.sqlqry(cmd$)
  Call tpltchk(tpid$, "Enddatum", trm(d1))
End If
Call rlist6(tpid$)
MousePointer = 0

End Sub

Private Sub delproj(pid$)
Dim id$, sq$, r As ADODB.Recordset, c$
  
id$ = pid$
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id FROM tpprogli where tpid ='" + id$ & "'", form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If Not r.EOF Then
  MsgBox (transe("Projekt") + " " + id$ & " " + transe("kann nicht gelöscht werden. Es sind Programme verknüpft."))
  Exit Sub
End If
Set r = New ADODB.Recordset
r.CursorLocation = adUseServer
rrr = form1.adoopen(r, "SELECT id FROM auftritt where tourneeplanid ='" + id$ & "' and auftrittstyp<>'Tournee' and auftrittstyp<>'Zeitraum'", form1.adoc, adOpenDynamic, adLockReadOnly)
If Not r.EOF Then
  MsgBox (transe("Projekt") + " " + id$ & " " + transe("kann nicht gelöscht werden. Es sind Auftritte vorhanden."))
  Exit Sub
End If
ask% = MsgBox(transe("Wirklich löschen?"), vbYesNo + vbCritical + vbDefaultButton2, transe("Projekt löschen?"))
If ask% = vbYes Then
  If id$ = "" Then Exit Sub
  MousePointer = 11: DoEvents
  Set r = New ADODB.Recordset
  r.CursorLocation = adUseServer
  rrr = form1.adoopen(r, "SELECT id FROM auftritt where tourneeplanid ='" + id$ & "' and (auftrittstyp='Tournee' or auftrittstyp='Zeitraum')", form1.adoc, adOpenDynamic, adLockReadOnly)
  While Not r.EOF
    sq$ = "delete from auftritthigru where auftrittsid='" + trm(r!id) & "'"
    Call form1.sqlqry(sq$)
    r.MoveNext
  Wend
    If Not form1.isfieldmissing("opt_prios", "id") Then
      c$ = "delete from opt_prios where evnt='T:" + id$ + "'"
      Call form1.sqlqry(c$)
    End If
    sq$ = "delete from tplan where id='" + id$ & "'"
    Call form1.sqlqry(sq$)
    sq$ = "delete from auftritt where tourneeplanid='" + id$ & "'"
    Call form1.sqlqry(sq$)
    sq$ = "delete from tpwernoch where tpid='" + id$ & "'"
    Call form1.sqlqry(sq$)
  End If
  If Not form1.isfieldmissing("opt_topics", "id") Then
    c$ = "delete from opt_topics where topicid='" & id$ & "'"
    Call form1.sqlqry(c$)
    c$ = "delete from sysvars where owner like 'sysvar_system_tlnk_" + id$ + "_%'"
    Call form1.sqlqry(c$)
    If form1.getusersetting("extralogtlnk", "no") = "ja" Then Call form1.log2f(c$, "tplan", "delproj")
  End If
  
  MousePointer = 0: DoEvents
  Call rlist1
End Sub
