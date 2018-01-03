VERSION 5.00
Object = "{E5A19D51-DD6B-11D4-AB81-BBEAD055682C}#1.0#0"; "Resizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCtl.ocx"
Begin VB.Form adrselect 
   Caption         =   "Adresse wählen"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command37 
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
      Left            =   1920
      Picture         =   "adrselect.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   80
      ToolTipText     =   "Neue Adresse anlegen"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Cloud Export"
      Enabled         =   0   'False
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
      Left            =   9720
      TabIndex        =   79
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command35 
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
      Left            =   840
      TabIndex        =   78
      ToolTipText     =   "Hilfeseite öffnen"
      Top             =   5520
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   8760
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10800
      Picture         =   "adrselect.frx":0392
      Style           =   1  'Grafisch
      TabIndex        =   76
      ToolTipText     =   "copy from clipboard"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command33 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10800
      TabIndex        =   75
      ToolTipText     =   "Eine andere Gruppe in die ausgewählte integrieren"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command32 
      Caption         =   "+"
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
      Height          =   240
      Left            =   10800
      TabIndex        =   73
      ToolTipText     =   "Zusätzliche Adresstypdaten exportieren"
      Top             =   2640
      Width           =   375
   End
   Begin VB.ListBox tempdel 
      Height          =   1575
      IntegralHeight  =   0   'False
      Left            =   12240
      TabIndex        =   72
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command31 
      Caption         =   "+ -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   71
      ToolTipText     =   "löscht die Adresse und fügt die Beziehung hinzu"
      Top             =   720
      Width           =   495
   End
   Begin VB.ListBox templist 
      Height          =   1575
      IntegralHeight  =   0   'False
      Left            =   12240
      TabIndex        =   70
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command30 
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
      Height          =   240
      Left            =   8640
      TabIndex        =   69
      ToolTipText     =   "Folgt der gewählten Beziehung und fügt die Adressen hinzu"
      Top             =   720
      Width           =   375
   End
   Begin VB.ComboBox Combo6 
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
      IntegralHeight  =   0   'False
      Left            =   9480
      TabIndex        =   68
      ToolTipText     =   "Beziehungen verfolgen"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Height          =   255
      Left            =   4320
      TabIndex        =   67
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6000
      Picture         =   "adrselect.frx":08C4
      Style           =   1  'Grafisch
      TabIndex        =   64
      ToolTipText     =   "Adressen im Umkreis um eine Postleitzahl suchen"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command29 
      Caption         =   "CSV-Liste"
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
      Left            =   9720
      TabIndex        =   63
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9840
      Picture         =   "adrselect.frx":0F36
      Style           =   1  'Grafisch
      TabIndex        =   62
      ToolTipText     =   "Serienbriefverwaltung"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   10680
      Picture         =   "adrselect.frx":0FD5
      Style           =   1  'Grafisch
      TabIndex        =   61
      ToolTipText     =   "Vorlage bearbeiten"
      Top             =   6840
      Width           =   375
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
      Left            =   10560
      Picture         =   "adrselect.frx":2127
      Style           =   1  'Grafisch
      TabIndex        =   60
      ToolTipText     =   "Neuen Serienbrief"
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command8 
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
      Left            =   10080
      Picture         =   "adrselect.frx":24B9
      Style           =   1  'Grafisch
      TabIndex        =   59
      ToolTipText     =   "Neue Gruppe anlegen"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Caption         =   "ODER-dazu"
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
      TabIndex        =   58
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Caption         =   "UND-dazu"
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
      Left            =   4320
      TabIndex        =   57
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   885
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   56
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5880
      Picture         =   "adrselect.frx":284B
      Style           =   1  'Grafisch
      TabIndex        =   55
      ToolTipText     =   "Speichern"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command22 
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
      Height          =   495
      Left            =   5880
      Picture         =   "adrselect.frx":28AB
      Style           =   1  'Grafisch
      TabIndex        =   54
      ToolTipText     =   "Als neuen Filter speichern"
      Top             =   3120
      Width           =   495
   End
   Begin VB.ListBox List5 
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   52
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "adrselect.frx":2C3D
      Left            =   6480
      List            =   "adrselect.frx":2C3F
      Sorted          =   -1  'True
      TabIndex        =   51
      Text            =   "schliesse aus:"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   7800
      Picture         =   "adrselect.frx":2C41
      Style           =   1  'Grafisch
      TabIndex        =   50
      ToolTipText     =   "schliesse aus"
      Top             =   2160
      Width           =   375
   End
   Begin VB.ListBox List3 
      Height          =   555
      Index           =   2
      IntegralHeight  =   0   'False
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   49
      ToolTipText     =   "ignoriere diese Adressgruppen"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command19 
      Cancel          =   -1  'True
      Caption         =   "CSV-Alle"
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
      Left            =   2760
      TabIndex        =   48
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CSV-Liste"
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
      Left            =   2760
      TabIndex        =   47
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   6120
      TabIndex        =   44
      ToolTipText     =   "Suchkriterien auf Kontakte (oder Adressen) anwenden"
      Top             =   5760
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.ListBox l4h 
      Height          =   2220
      IntegralHeight  =   0   'False
      Left            =   7560
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   2400
      TabIndex        =   41
      Top             =   3000
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "CSV-Liste"
      Height          =   255
      Left            =   2400
      TabIndex        =   40
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "adrselect.frx":3073
      Left            =   1200
      List            =   "adrselect.frx":3075
      Sorted          =   -1  'True
      TabIndex        =   38
      Text            =   "PLZ"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "adrselect.frx":3077
      Left            =   1200
      List            =   "adrselect.frx":3079
      Sorted          =   -1  'True
      TabIndex        =   37
      Text            =   "PLZ"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4320
      TabIndex        =   34
      ToolTipText     =   "Oder-Trennung mit |"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "adrselect.frx":307B
      Left            =   4320
      List            =   "adrselect.frx":3088
      TabIndex        =   33
      Text            =   "gleich"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6480
      TabIndex        =   32
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "adrselect.frx":30AA
      Left            =   6480
      List            =   "adrselect.frx":30AC
      TabIndex        =   31
      Text            =   "gleich"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   2220
      IntegralHeight  =   0   'False
      Left            =   6480
      MultiSelect     =   1  '1 -Einfach
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Serienbrief"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      TabIndex        =   29
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   8760
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "alle dazu"
      Height          =   255
      Left            =   9720
      TabIndex        =   27
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      Picture         =   "adrselect.frx":30AE
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Den markierten Satz löschen"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10800
      Picture         =   "adrselect.frx":359E
      Style           =   1  'Grafisch
      TabIndex        =   25
      ToolTipText     =   "Den markierten Satz löschen"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1200
      Picture         =   "adrselect.frx":3A8E
      Style           =   1  'Grafisch
      TabIndex        =   23
      ToolTipText     =   "gespeicherte Selektionen"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Kontakte merken"
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   22
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Adressen merken"
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   21
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      Text            =   "50"
      ToolTipText     =   "finde nur so viele Einträge, 0=alle"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   360
      Picture         =   "adrselect.frx":3C18
      Style           =   1  'Grafisch
      TabIndex        =   19
      Top             =   5520
      Width           =   495
   End
   Begin Resizer.axsResizer axsResizer1 
      Left            =   8160
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Kontakt(e) dazu"
      Height          =   735
      Left            =   8640
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "alle dazu"
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "A&dresse(n) dazu"
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   240
      Width           =   2175
   End
   Begin VB.ListBox grpmembers 
      Height          =   1425
      Left            =   8640
      TabIndex        =   15
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox currgrp 
      Height          =   285
      Left            =   8640
      TabIndex        =   14
      Top             =   3300
      Width           =   1335
   End
   Begin VB.ListBox allgrps 
      Height          =   2205
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   360
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&LOS"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "adrselect.frx":3E68
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Index           =   1
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "ignoriere diese Adressgruppen"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<=="
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "==>"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<--"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-->"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Index           =   0
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "finde nur diese Adressgruppen (leer=suche alles)"
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   360
      MultiSelect     =   2  'Erweitert
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   360
      MultiSelect     =   2  'Erweitert
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   280
      Width           =   1455
   End
   Begin VB.Image higrusuch 
      Height          =   360
      Left            =   5940
      Picture         =   "adrselect.frx":3E6E
      Stretch         =   -1  'True
      ToolTipText     =   "query"
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8760
      TabIndex        =   74
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "UND statt ODER"
      Height          =   255
      Left            =   4610
      TabIndex        =   66
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label umplz 
      BackStyle       =   0  'Transparent
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
      Left            =   6480
      TabIndex        =   65
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
      Height          =   255
      Left            =   4320
      TabIndex        =   53
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Sortierung"
      Height          =   255
      Left            =   360
      TabIndex        =   46
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Kriterien f. Kontakte"
      Height          =   255
      Left            =   4320
      TabIndex        =   45
      ToolTipText     =   "Suchkriterien auf Kontakte (oder Adressen) anwenden"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "mit Kontakten"
      Height          =   255
      Left            =   2640
      TabIndex        =   42
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10080
      TabIndex        =   39
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sortierung"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "suche nicht:"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "suche nur:"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   180
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   8520
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   4080
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "adrselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sel_vld As Integer, selct$
Dim break%, kontsel$, fl_rl3%, suchstr$, ksuchstr$, sel_break As Integer
Dim snotb4 As Double, msec As Double, suchvz As Double, plzlimit As String
Public kontselid$

Public Sub rlist1(sinp$)
Dim rtmp As ADODB.Recordset, r As ADODB.Recordset, cmd$, cmd1$, topsel$, cmd2$, wh$, i%, k%, s$
Dim i1%, na$, nx$, rrr, ads$, ppp%, usrwh$, usrwh1$, whx$, whx1$, cmdlen%, c$, l$, c1$, c2$
Dim whxkmerk$, whxkmerk1$, prvi$, whx2$, whx3$, lf$, c10m$, iot As Boolean, limitlen%
Dim filtersatz As String, whxall$, filter$, operator$, l5i%, plzfok As Boolean, whv$, ij$
Dim tmpv As String, sgrp As String, swert As String, typlist As String, fnlist As String

Dim d2infile As String, d2insub As String

d2infile = "adrselect": d2insub = "rlist1"

s$ = sinp$
List1.Clear
List2.Clear
Command7.Enabled = False
Command9.Enabled = False
limitlen% = trm0(form1.getusersetting("selectlimitlength", "0"))

retryadrcmd:

topsel$ = Val(Text3.text) * 2
If topsel$ = "0" Then
  topsel$ = ""
Else
  topsel$ = "top " + topsel$ + " "
End If

If Left(s$, 1) = ":" Then
  c$ = trm(s$): If Len(c$) < 2 Then Exit Sub
  c$ = Mid$(c$, 2)
  sgrp = cut_d1(c$, ":")
  swert = cut_d2bis(c$, ":")
  l$ = ""
  If sgrp <> "" Then
    c$ = "SELECT auftrittsfelder.FeldName,auftrittsfelder.typ FROM adresstypen INNER JOIN auftrittsfelder ON adresstypen.id = auftrittsfelder.typ "
    c$ = c$ + "Where auftrittsfelder.feldname Like '" + sgrp + "%' ORDER BY auftrittsfelder.FeldName"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
    rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
    While Not r.EOF
      c$ = trm(r!feldname)
      For i% = 0 To List3(0).ListCount - 1
        If List3(0).List(i%) = trm(r!typ) Then Exit For
      Next i%
      If i% < List3(0).ListCount Then
        If InStr(LCase(l$), "|" + LCase(c$)) = 0 Then l$ = l$ + "|" + c$
      End If
      r.MoveNext
    Wend
  End If
  s$ = ""
  sgrp = l$
End If

cmd1$ = "SELECT adresse.ID,adresse.name,adresse.plz,adresse.ort FROM "

If s$ <> "" Then
  cmd2$ = " ( instr(lcase(adresse.id),'" + LCase(s$) + "')>0 " + _
        "or instr( lcase(adresse.name),'" + LCase(s$) + "')>0 " + _
        "or instr(lcase(adresse.ort),'" + LCase(s$) + "')>0 ) "
Else
  cmd2$ = ""
End If

filtersatz = trm(Text6.text)
whxall$ = ""
Do
filter$ = "": operator$ = ""
If filtersatz <> "" Then
  filter$ = word1(filtersatz): filtersatz = word2bis(filtersatz)
  operator$ = word1(filtersatz): filtersatz = word2bis(filtersatz)
  For l5i% = 0 To List5.ListCount - 1
    If List5.List(l5i%) = filter$ Then
      List5.ListIndex = l5i%: DoEvents
      Exit For
    End If
  Next l5i%
End If
whx1$ = "": l4h.Clear
For i% = 0 To List4.ListCount - 1
  If List4.Selected(i%) And Text4.text <> "" Then
    l4h.AddItem List4.List(i%)
  End If
Next i%
For i% = 0 To List4.ListCount - 1
  If List4.Selected(i%) And (Text4.text <> "" Or transo(Combo2.text) = "leer" Or transo(Combo2.text) = "nicht leer") Then
    usrwh1$ = "adresse." + transo(List4.List(i%))
    Select Case transo(Combo2.text)
      Case "leer": whx2$ = "(isnull(" + usrwh1$ + ") or " + usrwh1$ + "='')"
      Case "nicht leer": whx2$ = "((not isnull(" + usrwh1$ + ")) and " + usrwh1$ + "<>'')"
      Case "gleich"
        whx2$ = usrwh1$ + "='" + Text4.text + "'"
      Case "enthält"
        whx2$ = "instr(lcase(" + usrwh1$ + "),lcase('" + Text4.text + "'))>0 "
        If l4h.ListCount > 1 Then
          For i1% = 0 To l4h.ListCount - 1
            If l4h.List(i1%) <> List4.List(i%) Then
              If whx3$ <> "" Then whx3$ = whx3$ + " and "
              whx3$ = whx3$ + " instr(lcase(" + "adresse." + l4h.List(i1%) + "),lcase('" + Text4.text + "'))>0 "
            End If
          Next i1%
          whx2$ = whx2$ + whx3$ + ")) "
        End If
      Case "beginnt mit":
        whx2$ = "instr(lcase(" + usrwh1$ + "),lcase('" + Text4.text + "'))=1 "
        If l4h.ListCount > 1 Then
          whx2$ = "(" + whx2$ + " or (isnull(" + usrwh1$ + ")) and ("
          whx3$ = ""
          For i1% = 0 To l4h.ListCount - 1
            If l4h.List(i1%) <> List4.List(i%) Then
              If whx3$ <> "" Then whx3$ = whx3$ + " and "
              whx3$ = whx3$ + " instr(lcase(" + "adresse." + l4h.List(i1%) + "),lcase('" + Text4.text + "'))=1 "
            End If
          Next i1%
          whx2$ = whx2$ + whx3$ + ")) "
        End If
      Case Else:
        whx2$ = ""
    End Select
    If whx2$ <> "" Then
      If whx1$ <> "" Then whx1$ = whx1$ + " or "
      whx1$ = whx1$ + " " + whx2$
    End If
  End If
Next i%
If whx1$ <> "" Then
  whx1$ = "(" + whx1$ + ")"
End If
whxall$ = whxall$ + " " + whx1$ + " " + operator$
Loop Until filtersatz = ""
If trm(whxall$) <> "" Then whx1$ = "(" + whxall$ + ")"
whxkmerk$ = whx1$
whx$ = ""
If List3(0).ListIndex >= 0 And Text5.text <> "" Then
  usrwh$ = "'" + transo(List3(0).List(List3(0).ListIndex)) + "'"
  Select Case Combo3.text
    Case "gleich"
      whx$ = "(" + wertsubst(Text5.text, transo(List3(0).List(List3(0).ListIndex))) + ")"
    Case "enthält"
      whx$ = wertsubst2(Text5.text, transo(List3(0).List(List3(0).ListIndex)))
    Case "beginnt mit":
      whx$ = wertsubst3(Text5.text, transo(List3(0).List(List3(0).ListIndex)))
    Case Else:
        whx$ = ""
  End Select
End If
If whx$ <> "" Then
  whx$ = " and " + whx$
  whxkmerk1$ = whx$
End If

wh$ = " where "
If List3(0).ListCount > 0 Then
  ij$ = "adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid": whv$ = ""

          If sgrp <> "" And swert <> "" Then
            ij$ = "(adresse INNER JOIN adresstyp ON adresse.ID = adresstyp.vid) INNER JOIN auftritthigru ON adresse.ID = auftritthigru.auftrittsid"
            l$ = trm(cut_d1(ads$, "("))
            c2$ = Mid$(sgrp, 2)
            fnlist = "": typlist = ""
            While c2$ <> ""
              c1$ = cut_d1(c2$, "|")
              c2$ = cut_d2bis(c2$, "|")
              If fnlist <> "" Then fnlist = fnlist + " and "
              fnlist = fnlist + "auftritthigru.FeldName='" + c1$ + "'"
            Wend
            For i% = 0 To List3(0).ListCount - 1
              If i% > 0 Then typlist = typlist + " or "
              typlist = typlist + "auftritthigru.auftrittstyp='" + List3(0).List(i%) + "'"
            Next i%
            whv$ = "(auftritthigru.FeldDaten like '%" + swert + "%' and (" + fnlist + ") and (" + typlist + ")) and "
'            Debug.Print whv$
          End If



  wh$ = ij$ + " WHERE " + whv$ + "(((kid='-1') or isnull(kid)) and ("
  For i% = 0 To List3(0).ListCount - 1
    If i% = List3(0).ListIndex Then
      wh$ = wh$ + "( ((adresstyp.typ)='" + transo(List3(0).List(i%)) + "') " + whx$ + ")"
    Else
      wh$ = wh$ + "( ((adresstyp.typ)='" + transo(List3(0).List(i%)) + "') )"
    End If
    If i% < List3(0).ListCount - 1 Then wh$ = wh$ + "or "
  Next i%
  wh$ = wh$ + ")) ": If cmd2$ <> "" Then wh$ = wh$ + "and "
End If
If LCase(Right$(cmd1$, 5)) = "from " Then cmd1$ = cmd1$ + "adresse "
If trm(wh$ + cmd2$) <> "where" Then
  cmd$ = cmd1$ + wh$ + cmd2$
Else
  cmd$ = cmd1$
End If
If whx1$ <> "" Then
  If cmd2$ <> "" Or InStr(LCase(cmd$), "where") > 0 Then
    cmd$ = cmd$ + " and "
  Else
    cmd$ = cmd$ + " where "
  End If
  cmd$ = cmd$ + " " + whx1$
End If
If LCase(Right$(cmd$, 5)) = "from " Then cmd$ = cmd$ + "adresse "
If trm(Combo5.text) <> "" Then cmd$ = cmd$ + " order by adresse." + transo(trm(Combo5.text))
Text2.text = cmd$

suchstr$ = ""
ksuchstr$ = ""
k% = Val("0" + trm(Text3.text))
cmdlen% = Len(cmd$)
If limitlen% > 0 And cmdlen% > limitlen% Then
  If List3(0).ListCount > 0 Then
    List3(0).ListIndex = List3(0).ListCount - 1
    Call moveme(0, 1)
    GoTo retryadrcmd
  End If
End If

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
If Not rtmp.EOF Then
  suchstr$ = cmd$
  rtmp.MoveFirst
  Command7.Enabled = True
  Command9.Enabled = True
  c10m$ = transo(Combo10.text)
  While Not rtmp.EOF And List1.ListCount < k% And break% = 0
    plzfok = True
    If plzlimit <> "" Then
      plzfok = False
      tmpv = trm(rtmp!plz)
      If tmpv <> "" And InStr(plzlimit, "|" + tmpv + "|") > 0 Then
        plzfok = True
      Else
        tmpv = word1(trm(rtmp!ort))
        If tmpv <> "" And isnumber(tmpv) Then
          If InStr(plzlimit, "|" + tmpv + "|") > 0 Then plzfok = True
        End If
      End If
    End If
    'UND statt ODER für adressgruppen
    If Check3.value = 1 Then
      For ppp% = 0 To List3(0).ListCount - 1
        If form1.isoftype(rtmp!id, transo(List3(0).List(ppp%))) = "-1" Then
          plzfok = False
          Exit For
        End If
      Next ppp%
    End If
    If plzfok Then
      ads$ = rtmp!id + "(" + trm(rtmp!name) + ")"
      For ppp% = 0 To List1.ListCount - 1
        If List1.List(ppp%) = ads$ Then
          ads$ = ""
          Exit For
        End If
      Next ppp%
      If ads$ <> "" Then
        If List3(2).ListCount > 0 Then
          For ppp% = 0 To List3(2).ListCount - 1
            iot = form1.isoftype(rtmp!id, transo(List3(2).List(ppp%))) <> "-1"
            If iot And c10m$ = "schliesse aus:" _
            Or Not iot And c10m$ <> "schliesse aus:" _
            Then
              ads$ = ""
              Exit For
            End If
          Next ppp%
        End If
        If ads$ <> "" Then
          List1.AddItem form1.crlffake(ads$)
          DoEvents
        End If
      End If
    End If   'plzfok
    rtmp.MoveNext
  Wend
  rtmp.Close
End If
End If
Label3.Caption = "(" + trm(List1.ListCount) + ")"

retrykcmd:
wh$ = ""
ij$ = "(kontakt INNER JOIN adresstyp ON kontakt.id = adresstyp.kid) INNER JOIN adresse ON kontakt.vid = adresse.ID"

          If sgrp <> "" And swert <> "" Then
            ij$ = "((kontakt INNER JOIN adresstyp ON kontakt.id = adresstyp.kid) INNER JOIN adresse ON kontakt.vid = adresse.ID) INNER JOIN auftritthigru ON concat(kontakt.vid, kontakt.id) = auftritthigru.auftrittsid"
'            whv$ = "(auftritthigru.FeldDaten like '%" + swert + "%' and (" + fnlist + ") and (" + typlist + ")) and "
'            Debug.Print whv$
          End If


cmd1$ = "SELECT kontakt.name as name ,kontakt.id as id,kontakt.vid as vid,kontakt.plz as kplz,kontakt.ort as kort,adresse.plz as aplz,adresse.ort as aort, adresstyp.typ as atyp " + _
        "FROM " + ij$ + " where " + whv$ + "(instr(lcase(kontakt.name),'" + LCase(s$) + "')>0 or instr(lcase(adresse.name),'" + LCase(s$) + "')>0 or instr(lcase(adresse.ort),'" + LCase(s$) + "')>0) "
If whxkmerk$ <> "" Then
  If Check2.value = 1 Then
    cmd1$ = cmd1$ + " and (" + strrepl(strrepl(whxkmerk$, "adresse.", "kontakt."), ".Land", ".lkz") + ") "
  Else
    cmd1$ = cmd1$ + " and (" + whxkmerk$ + ") "
  End If
End If
If List3(0).ListCount > 0 Then
  wh$ = "and ("
  For i% = 0 To List3(0).ListCount - 1
    wh$ = wh$ + " ((adresstyp.typ)='" + transo(List3(0).List(i%)) + "') "
    If i% < List3(0).ListCount - 1 Then wh$ = wh$ + " or "
  Next i%
  If whxkmerk1$ <> "" Then wh$ = wh$ + whxkmerk1$
  wh$ = wh$ + ") "
End If
prvi$ = "-1xxxxxxxxxxxxxxxx"
cmd1$ = cmd1$ + wh$
If trm(Combo4.text) <> "" Then
  If Check2.value = 0 Then
    cmd1$ = cmd1$ + " order by adresse." + transo(trm(Combo5.text))
  Else
    lf$ = transo(trm(Combo4.text))
    If LCase(lf$) = "land" Then lf$ = "lkz"
    cmd1$ = cmd1$ + " order by kontakt." + lf$
  End If
End If
cmdlen% = Len(cmd1$)
Call form1.dbg2f("adrselect.rlist1(" + trm(cmdlen%) + "):" + cmd1$)
If limitlen% > 0 And cmdlen% > limitlen% Then
  If List3(0).ListCount > 0 Then
    List3(0).ListIndex = List3(0).ListCount - 1
    Call moveme(0, 1)
    GoTo retrykcmd
  End If
End If
ksuchstr$ = cmd1$
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, cmd1$, form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr = 0 Then
If Not rtmp.EOF Then
rtmp.MoveFirst
c10m$ = transo(Combo10.text)
While Not rtmp.EOF And List2.ListCount < k% And break% = 0
  If prvi$ <> rtmp!id Then
    prvi$ = rtmp!id: ads$ = prvi$
    DoEvents
    na$ = trm(rtmp!name)
    nx$ = form1.getkontaktpositionbyid(rtmp!id)
    If nx$ <> "" Then na$ = na$ + " (" + nx$ + ")"
    If List3(2).ListCount > 0 Then
      For ppp% = 0 To List3(2).ListCount - 1
        iot = form1.kisoftype(ads$, transo(List3(2).List(ppp%))) <> "-1"
        If iot And c10m$ = "schliesse aus:" _
        Or Not iot And c10m$ <> "schliesse aus:" _
        Then
          ads$ = ""
          Exit For
        End If
      Next ppp%
    End If
    If ads$ <> "" Then
      
      plzfok = True
      If plzlimit <> "" Then
        plzfok = False
        tmpv = trm(rtmp!kplz)
        If tmpv <> "" And InStr(plzlimit, "|" + tmpv + "|") > 0 Then
          plzfok = True
        Else
          tmpv = word1(trm(rtmp!kort))
          If tmpv <> "" And isnumber(tmpv) Then
            If InStr(plzlimit, "|" + tmpv + "|") > 0 Then
              plzfok = True
            End If
          Else
            tmpv = trm(rtmp!aplz)
            If tmpv <> "" And InStr(plzlimit, "|" + tmpv + "|") > 0 Then
              plzfok = True
            Else
              tmpv = word1(trm(rtmp!aort))
              If tmpv <> "" And isnumber(tmpv) Then
                If InStr(plzlimit, "|" + tmpv + "|") > 0 Then
                  plzfok = True
                End If
              End If
            End If
          End If
        End If
      End If
      If Check3.value = 1 Then
        For ppp% = 0 To List3(0).ListCount - 1
          If Not (form1.kisoftype(rtmp!id, transo(List3(0).List(ppp%))) <> "-1") Then
            plzfok = False
            Exit For
          End If
        Next ppp%
      End If
      wh$ = na$ + Space$(160) + " (VID:" + rtmp!vid + ") " + "ID:" + rtmp!id
      For ppp% = 0 To List2.ListCount - 1
        If List2.List(ppp%) = wh$ Then
          plzfok = False
          Exit For
        End If
      Next ppp%
      If plzfok Then
        List2.AddItem form1.crlffake(na$) + Space$(160) + " (VID:" + rtmp!vid + ") " + "ID:" + rtmp!id
      End If
    End If
  End If
  rtmp.MoveNext
Wend
rtmp.Close
If List1.ListCount = 0 Then Command7.Enabled = False
If List2.ListCount = 0 Then Command9.Enabled = False
End If
End If
Label4.Caption = "(" + trm(List2.ListCount) + ")"
Exit Sub

errhdl:
  rrr = Err
  If rrr <> 0 Then
    If rrr <> 3420 Then MsgBox "Fehler #" + rrr + " " + Error$(rrr)
    On Error GoTo 0
    break% = 1
    Unload adrselect
    Exit Sub
  End If
  Resume Next
End Sub

Sub rlist3()
Dim rtmp As ADODB.Recordset, inifile As String, o%, l$, i%, rrr
Dim d2infile As String, d2insub As String

d2infile = "adrselect": d2insub = "rlist3"
If fl_rl3% = 1 Then Exit Sub
fl_rl3% = 1
List3(0).Clear

inifile = form1.mydatadir() + "\" + Me.name + ".ini"
If exist(inifile) = 1 Then
  o% = FreeFile
  Open inifile For Input As #o%
  While Not EOF(o%)
    Line Input #o%, l$
    If Left(l$, 4) <> "rel:" Then List3(1).AddItem transe(l$)
  Wend
  Close #o%
End If
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresstypen", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)

While Not rtmp.EOF
  For i% = 0 To List3(1).ListCount - 1
    If List3(1).List(i%) = transe(rtmp!id) Then i% = List3(1).ListCount + 100
  Next i%
  If Left(rtmp!id, 4) <> "rel:" Then
    If i% < List3(1).ListCount + 50 Then List3(0).AddItem transe(rtmp!id)
  Else
    Combo6.AddItem Mid(rtmp!id, 5)
  End If
  rtmp.MoveNext
Wend
Text1.text = ""
fl_rl3% = 0

End Sub

Private Sub allgrps_Click()
Dim rtmp As ADODB.Recordset, gn$, rrr

MousePointer = 11: DoEvents
gn$ = allgrps.List(allgrps.ListIndex)
currgrp.text = gn$
grpmembers.Clear

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppen where grpid='" + gn$ + "'", form1.adoc, adOpenDynamic, adLockReadOnly)

While Not rtmp.EOF
  If Not IsNull(rtmp!kid) And rtmp!kid <> "-1" Then
    grpmembers.AddItem form1.get_kontaktname_by_id(rtmp!kid) + Space$(40) + "(ID:" + rtmp!id
  Else
    grpmembers.AddItem rtmp!adressid + Space$(40) + "(ID:" + rtmp!id
  End If
  rtmp.MoveNext
Wend
Command32.Enabled = True
Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"
MousePointer = 0
End Sub

Private Sub allgrps_DblClick()
Dim i%

i% = allgrps.ListIndex
If i% < 0 Then Exit Sub

selct$ = "GRUPPEGEWÄHLT:" + allgrps.List(i%)
kontsel$ = ""
If sel_vld <> -1 Then
  sel_vld = 1
End If

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
  Combo4.Enabled = True
Else
  Combo4.text = ""
  Combo4.Enabled = False
End If
End Sub

Private Sub Check3_Click()
Call Command6_Click

End Sub

Private Sub Combo1_DropDown()
Call rcombo1
End Sub

Private Sub Command1_Click()
Dim d0

break% = 1
MousePointer = 11
sel_break = 1
d0 = Time
sel_vld = 0
While (Time - d0) * 24 * 60 * 60 < 1: DoEvents: Wend
MousePointer = 0
sel_break = 1
sel_vld = 0
Unload adrselect
sel_vld = 0
sel_break = 1
End Sub

Private Sub Command10_Click()
Dim an$, nop, i%, gn$, nid$, j%, glan$, p%

For j% = 0 To List1.ListCount - 1
If List1.Selected(j%) Then
  an$ = List1.List(j%)
  If InStr(an$, "(") > 1 Then an$ = Left$(an$, InStr(an$, "(") - 1)
  nop = 0
  For i% = 0 To grpmembers.ListCount - 1
    glan$ = grpmembers.List(i%)
    p% = InStr(glan$, "(ID:")
    glan$ = trm(Left$(glan$, p% - 1))
    If glan$ = an$ Then
      nop = 1
      i% = grpmembers.ListCount
    End If
  Next i%
  If nop = 0 Then
    gn$ = currgrp.text
    nop = 1
    If gn$ <> "" Then
      nid$ = form1.newid("adressgruppen", "id", 20)
      form1.sqlqry ("insert into adressgruppen (id,adressid,grpid,kid) values('" + nid$ + "','" + an$ + "','" + gn$ + "','-1')")
      grpmembers.AddItem an$ + Space$(40) + "(ID:" + nid$
      nop = 0
    End If
  End If
  List1.Selected(j%) = False
  DoEvents
End If
Next j%

Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"

End Sub

Private Sub Command11_Click()
Dim an$, gn$
If grpmembers.ListIndex < 0 Then Exit Sub
an$ = grpmembers.List(grpmembers.ListIndex)
If InStr(an$, "(ID:") > 0 Then
  an$ = trm(Mid$(an$, InStr(an$, "(ID:") + 4))
  gn$ = trm(currgrp.text)
  If gn$ <> "" Then
    form1.sqlqry ("delete from adressgruppen where id='" + an$ + "'")
    grpmembers.RemoveItem grpmembers.ListIndex
  End If
End If
Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"

End Sub

Private Sub Command12_Click()
Dim i%
MousePointer = 11: DoEvents
For i% = 0 To List1.ListCount - 1
  List1.Selected(i%) = True
Next i%
Call Command10_Click
Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"
MousePointer = 0

End Sub

Private Sub Command13_Click()
Dim an$, kid$, vid$, i%, gn$, nid$, apn1%, ant$, j%
Dim rtmp As ADODB.Recordset, cmd$, rrr

gn$ = currgrp.text
For j% = 0 To List2.ListCount - 1
If List2.Selected(j%) = True Then
  an$ = List2.List(j%)
  apn1% = InStr(an$, "     ")
  ant$ = ""
  If apn1% > 0 Then ant$ = trm(Left(an$, apn1%))
  If InStr(an$, "(VID:") > 1 Then vid$ = Mid$(an$, InStr(an$, "(VID:") + 5)
  If InStr(vid$, ") ID:") > 1 Then
    kid$ = Mid$(vid$, InStr(vid$, ") ID:") + 5)
    vid$ = Left$(vid$, InStr(vid$, ") ID:") - 1)
  End If
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
  cmd$ = "SELECT id FROM adressgruppen where AdressID='" + vid$ + "' and kid='" + kid$ + "' and grpid='" + gn$ + "'"
  rrr = form1.adoopen(rtmp, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly, "adrselect", "commad13_click")
  If rtmp.EOF Then
    nid$ = form1.newid("adressgruppen", "id", 20)
    form1.sqlqry ("insert into adressgruppen  (id,adressid,grpid,kid) values('" + nid$ + "','" + vid$ + "','" + gn$ + "','" + kid$ + "')")
    grpmembers.AddItem form1.get_kontaktname_by_id(kid$) + Space$(40) + "(ID:" + nid$
  End If
  List2.Selected(j%) = False
End If
Next j%
Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"

End Sub



Private Sub Command14_Click()
Dim an$, c$

an$ = allgrps.List(allgrps.ListIndex)
c$ = "delete FROM adressgruppen where grpid='" + an$ + "'": Call form1.sqlqry(c$)
c$ = "delete FROM adressgruppenindex where id='" + an$ + "'": Call form1.sqlqry(c$)
'On Error Resume Next
'Kill form1.mydir() + "\" + an$ + ".cso"
'On Error GoTo 0
grpmembers.Clear
Call rallgroups

End Sub

Private Sub Command15_Click()
Dim i%, i1%
MousePointer = 11: DoEvents
For i% = 0 To List2.ListCount - 1
    List2.Selected(i%) = True
Next i%
Call Command13_Click
Label7.Caption = "(" + trm(grpmembers.ListCount) + ")"
MousePointer = 0
End Sub

Public Sub Command16_Click()
Dim tr$, ffn$, i%, rrr, X, an$, r As ADODB.Recordset, tr1$

ffn$ = trm(Combo1.text)
If ffn$ = "" Then
  MsgBox "Bitte erst eine Vorlage auswählen."
  Exit Sub
End If
If grpmembers.ListCount = 0 Then
  MsgBox "Bitte erst Adressen auswählen."
  Exit Sub
End If
tr$ = form1.vorlagenverzeichnis() + "\"
tr1$ = "serienbrief_" + ffn$ + ".rtf"
If InStr(ffn$, " ") > 0 Then
  MsgBox "Bitte entfernen Sie alle Leerzeichen aus dem Dateinamen:" + vbCrLf + tr$ + tr1$
  Exit Sub
End If
tr$ = tr$ + tr1$
If exist(tr$) = 0 Then
  MsgBox "Die Vorlage '" + tr$ + "' existiert nicht."
  Exit Sub
End If
ffn$ = form1.mydatadir() + "\" + ffn$ + "_" + datum2sql(Date) + "_" + strrepl(Time, ":", "")
On Error Resume Next
MkDir ffn$
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  MsgBox "Das Verzeichnis '" + ffn$ + "' kann nicht angelegt werden."
  Exit Sub
End If
shwAdrDetail.Check3.value = 1: DoEvents
For i% = 0 To grpmembers.ListCount - 1
  grpmembers.ListIndex = i%
  DoEvents
  an$ = grpmembers.List(i%)
  If InStr(an$, "(ID:") > 0 Then
    an$ = "select * from adressgruppen where id='" + trm(Mid$(an$, InStr(an$, "(ID:") + 4)) + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, an$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      Call shwAdrDetail.savecheck
      Call shwAdrDetail.refreshadrdetail(r!adressid, r!kid)
      Call shwAdrDetail.SetFocus
      DoEvents
      Call form1.faxan(r!adressid, r!kid, tr1$, Right$(ffn$, 60), "", ffn$, "defaultname|noshow")
      DoEvents
    End If
  End If
Next i%

On Error Resume Next
X = Shell("explorer.exe " + ffn$, vbNormalFocus)
On Error GoTo 0

End Sub


Private Sub Command17_Click()
Dim an$, nop, c$, j%, plzort$, plzport$, p%, anid$, r As ADODB.Recordset, s As ADODB.Recordset
Dim fn$, o%, X, acnt%, apcnt%, kcnt%, plzstr$, rrr, xld$

acnt% = 0: apcnt% = 0: kcnt% = 0
fn$ = form1.myuniquedocname("", "csv")
If fn$ = "" Then Exit Sub
xld$ = form1.getusersetting("exceldelimiter", ",")
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, """" + "Sortiername" + """" + xld$ + """" + "Name" + """" + xld$ + """" + "PLZ/Ort / Postfach" + """" + xld$ + """" + "Kontakt" + """"
MousePointer = 11: DoEvents
For j% = 0 To List1.ListCount - 1
  anid$ = List1.List(j%)
  p% = InStr(anid$, "(")
  If p% > 1 Then
    anid$ = Left$(anid$, p% - 1)
    c$ = "select id,name,plz,ort,strasse,plzpostfach,postfach from adresse where id='" + anid$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      plzort$ = trm(r!plz): If plzort$ <> "" Then plzort$ = plzort$ + " "
      plzort$ = plzort$ + " " + trm(r!ort)
      plzport$ = trm(r!plzpostfach): If plzport$ <> "" Then plzport$ = plzport$ + " "
      plzport$ = plzport$ + " " + trm(r!postfach)
      plzstr$ = trm(r!plz) + trm(r!strasse)
      If trm(plzstr$) <> "" Or (trm(plzstr$) = "" And trm(plzport$) = "") Then
        c$ = """" + trm(r!id) + """" + xld$ + """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + plzort$ + """"
        Print #o%, c$
        acnt% = acnt% + 1
      End If
      If trm(plzport$) <> "" Then
        c$ = """" + trm(r!id) + """" + xld$ + """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + plzport$ + """"
        Print #o%, c$
        apcnt% = apcnt% + 1
      End If
      If Check1.value = 1 Then
        c$ = "select name from kontakt where vid='" + anid$ + "'"
        Set s = New ADODB.Recordset
        s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
        While Not s.EOF
          c$ = """" + "" + """" + xld$ + """" + "" + """" + xld$ + """" + "" + """" + xld$ + """" + strrepl(trm(s!name), vbCrLf, " - ") + """"
          Print #o%, c$
          kcnt% = kcnt% + 1
          s.MoveNext
        Wend
      End If
      Print #o%, """" + "" + """" + xld$ + """" + "" + """" + xld$ + """" + "" + """" + xld$ + """" + "" + """"
    End If
  End If
Next j%
Print #o%, """" + trm(acnt%) + " Adressen" + """" + xld$ + """" + trm(apcnt%) + " Postfach-Adressen" + """" + xld$ + """" + trm(kcnt%) + " Kontakte" + """" + xld$ + """" + "" + """"
Close #o%
X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)

MousePointer = 0

End Sub

Private Sub kcsvlist(fn$)
Dim an$, nop, c$, j%, plzort$, plzport$, p%, anid$, r As ADODB.Recordset, s As ADODB.Recordset
Dim o%, X, acnt%, apcnt%, kcnt%, c1$, c2$, plzstr$, rrr, xld$

acnt% = 0: apcnt% = 0: kcnt% = 0
If fn$ = "" Then Exit Sub
xld$ = form1.getusersetting("exceldelimiter", ",")
o% = FreeFile
Open fn$ For Append As #o%
MousePointer = 11: DoEvents
For j% = 0 To List2.ListCount - 1
  List2.ListIndex = j%: DoEvents
  anid$ = List2.List(j%)
  p% = InStr(anid$, " ID:")
  If p% > 1 Then
    anid$ = Mid$(anid$, p% + 4)
    c$ = "select id,vid,name,plz,ort,strasse,plzpostfach,postfach from kontakt where id='" + anid$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      plzort$ = trm(r!plz): If plzort$ <> "" Then plzort$ = plzort$ + " "
      plzort$ = plzort$ + " " + trm(r!ort)
      plzport$ = trm(r!plzpostfach): If plzport$ <> "" Then plzport$ = plzport$ + " "
      plzport$ = plzport$ + " " + trm(r!postfach)
      c1$ = "": c2$ = ""
      plzstr$ = trm(r!plz) + trm(r!strasse)
      If trm(plzstr$) <> "" Or (trm(plzstr$) = "" And trm(plzport$) = "") Then
        c1$ = """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + plzort$ + """"
      End If
      If trm(plzport$) <> "" Then
        c2$ = """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + plzport$ + """"
      End If
      c$ = "select id,name from adresse where id='" + r!vid + "'"
      Set s = New ADODB.Recordset
      s.CursorLocation = adUseServer
rrr = form1.adoopen(s, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
      If Not s.EOF Then
        If c1$ <> "" Then
          c$ = """" + s!id + """" + xld$ + """" + s!name + """," + c1$
          Print #o%, c$
          acnt% = acnt% + 1
        End If
        If c2$ <> "" Then
          c$ = """" + s!id + """" + xld$ + """" + s!name + """," + c2$
          Print #o%, c$
          apcnt% = acnt% + 1
        End If
        kcnt% = kcnt% + 1
        s.MoveNext
      End If
    End If
  End If
Next j%
Close #o%

MousePointer = 0

End Sub

Private Sub Command18_Click()
Dim fn$, o%, X

fn$ = form1.myuniquedocname("", "csv")
If fn$ = "" Then Exit Sub
On Error Resume Next
Kill fn$
On Error GoTo 0
o% = FreeFile
Open fn$ For Append As #o%
Print #o%, """" + "Sortiername" + """,""" + "Adressname" + """,""" + "Kontakt" + """,""" + "PLZ/Ort / Postfach" + """"
Close #o%
Call kcsvlist(fn$)
X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)

End Sub

Private Sub Command19_Click()
Dim an$, nop, c$, j%, plzort$, plzport$, p%, anid$, r As ADODB.Recordset, s As ADODB.Recordset
Dim fn$, o%, X, acnt%, apcnt%, kcnt%, plzstr$, rrr, xld$

acnt% = 0: apcnt% = 0: kcnt% = 0
fn$ = form1.myuniquedocname("", "csv")
If fn$ = "" Then Exit Sub
xld$ = form1.getusersetting("exceldelimiter", ",")
o% = FreeFile
Open fn$ For Output As #o%
Print #o%, """" + "Sortiername" + """" + xld$ + """" + "Adressname" + """" + xld$ + """" + "Kontakt" + """" + xld$ + """" + "PLZ/Ort / Postfach" + """"
MousePointer = 11: DoEvents
For j% = 0 To List1.ListCount - 1
  anid$ = List1.List(j%)
  p% = InStr(anid$, "(")
  If p% > 1 Then
    anid$ = Left$(anid$, p% - 1)
    c$ = "select id,name,strasse,plz,ort,plzpostfach,postfach from adresse where id='" + anid$ + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      plzort$ = trm(r!plz): If plzort$ <> "" Then plzort$ = plzort$ + " "
      plzort$ = plzort$ + " " + trm(r!ort)
      plzport$ = trm(r!plzpostfach): If plzport$ <> "" Then plzport$ = plzport$ + " "
      plzport$ = plzport$ + " " + trm(r!postfach)
      plzstr$ = trm(r!plz) + trm(r!strasse)
      If trm(plzstr$) <> "" Or (trm(plzstr$) = "" And trm(plzport$) = "") Then
        c$ = """" + trm(r!id) + """" + xld$ + """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + """" + xld$ + """" + plzort$ + """"
        Print #o%, c$
        acnt% = acnt% + 1
      End If
      If trm(plzport$) <> "" Then
        c$ = """" + trm(r!id) + """" + xld$ + """" + strrepl(trm(r!name), vbCrLf, " - ") + """" + xld$ + """" + """" + xld$ + """" + plzport$ + """"
        Print #o%, c$
        apcnt% = apcnt% + 1
      End If
    End If
  End If
Next j%
Close #o%
Call kcsvlist(fn$)
X = Shell("explorer.exe " + DirName(fn$), vbNormalFocus)

MousePointer = 0


End Sub

Private Sub Command2_Click()
break% = 1
Call moveme(0, 1)
Call wrini
Call Command6_Click
End Sub

Sub moveme(f%, t%)
Dim mi%

If List3(f%).ListCount = 0 Then Exit Sub
mi% = List3(f%).ListIndex
If mi% < 0 Then
  List3(f%).ListIndex = 0
  mi% = 0
End If
List3(t%).AddItem List3(f%).List(mi%)
List3(f%).RemoveItem mi%

If List3(f%).ListCount > 0 Then
  If mi% >= List3(f%).ListCount Then mi% = List3(f%).ListCount - 1
  List3(f%).ListIndex = mi%
End If

End Sub

Private Sub Command20_Click()
Call moveme(1, 2)

End Sub

Private Sub Command21_Click()
Dim i%, fn$

i% = List5.ListIndex
If i% < 0 Then
  Call Command22_Click
  Exit Sub
End If
Call fltsaveas(form1.mydatadir() + "\" + List5.List(i%) + ".flt")

End Sub

Private Sub Command22_Click()
Dim fn$, o%, i%

fn$ = form1.myuniquedocname("", "flt")
If fn$ = "" Then Exit Sub
Call fltsaveas(fn$)
Call rlist5

End Sub
Private Sub fltsaveas(f$)
Dim fn$, o%, i%

fn$ = f$
o% = FreeFile
Open fn$ For Output As #o%
For i% = 0 To List4.ListCount - 1
  If List4.Selected(i%) Then Print #o%, List4.List(i%)
Next i%
Print #o%, "<<EOL>>"
Print #o%, Combo2.text
Print #o%, Text4.text
Close #o%

End Sub

Private Sub Command23_Click()
Dim i%
i% = List5.ListIndex
If i% < 0 Then
  MsgBox (transe("Wählen Sie einen Filter."))
  Exit Sub
End If

If trm(Text6.text) = "" Then
  Text6.text = List5.List(i%)
Else
  Text6.text = Text6.text + " AND " + List5.List(i%)
End If
End Sub

Private Sub Command24_Click()
Dim i%
i% = List5.ListIndex
If i% < 0 Then
  MsgBox (transe("Wählen Sie einen Filter."))
  Exit Sub
End If

If trm(Text6.text) = "" Then
  Text6.text = List5.List(i%)
Else
  Text6.text = Text6.text + " OR " + List5.List(i%)
End If

End Sub

Private Sub Command25_Click()
Load sels
On Error Resume Next
Call sels.SetFocus
On Error GoTo 0

End Sub

Private Sub Command26_Click()
Dim ffn$, tr$, neuid As String, tr1$, templ$, i%, trgfn$

ffn$ = trm(Combo1.text)
If ffn$ = "" Then
  MsgBox "Bitte erst einen Serienbrief als Vorlage auswählen."
  Exit Sub
End If
tr$ = form1.vorlagenverzeichnis() + "\"
tr1$ = "serienbrief_" + ffn$ + ".rtf"
tr$ = tr$ + tr1$
If exist(tr$) = 0 Then
  MsgBox "Die Vorlage '" + tr$ + "' existiert nicht."
  Exit Sub
End If
templ$ = datum2sql(trm(Date))
i% = allgrps.ListIndex
If i% >= 0 Then templ$ = allgrps.List(i%) + "_" + templ$
neuid = InputBox(transe("Name der neue Serienbriefvorlage:"), "Neue Serienbriefvorlage erstellen.", templ$)
If trm(neuid) = "" Then Exit Sub
trgfn$ = form1.vorlagenverzeichnis() + "\serienbrief_" + neuid + ".rtf"
If exist(trgfn$) <> 0 Then
  MsgBox transe("Diese Vorlage existiert bereits.")
  Exit Sub
End If
Call FileCopy(tr$, trgfn$)
Combo1.text = neuid
Call Command27_Click
End Sub

Private Sub Command27_Click()
Dim ffn$, tr$, neuid As String, tr1$, templ$, i%, trgfn$

ffn$ = trm(Combo1.text)
If ffn$ = "" Then Exit Sub
tr$ = form1.vorlagenverzeichnis() + "\"
tr1$ = "serienbrief_" + ffn$ + ".rtf"
tr$ = tr$ + tr1$
If exist(tr$) = 0 Then
  MsgBox "Die Vorlage '" + tr$ + "' existiert nicht."
  Exit Sub
End If
Call form1.openthisdoc(tr$, "noconvert")
End Sub

Private Sub Command28_Click()
Load verwalt_sbf
On Error Resume Next
Call verwalt_sbf.SetFocus
On Error GoTo 0
End Sub

Private Sub Command29_Click()
Dim i%, anid$, p%, an$, r As ADODB.Recordset, r1 As ADODB.Recordset, r2 As ADODB.Recordset, r3 As ADODB.Recordset
Dim fld As Field, c$, j%, rc$, hdk$, hda$, tw$, rtw As ADODB.Recordset, rrr, o%
Dim fk%, fa%, ofn$, X, xld$, adrnam$, ofna$, ofnk$, plzo$, plzs$, k%
Dim grpl(99) As String, grplp, cmd$, t0 As Double, tn As Double, Dt As Double, spr As Double
Dim test As Double, nrecs As Long, tot As Double, anr$
Dim wabtab(0 To 2, 0 To 99) As String, wabptr As Integer, wabl As Integer, wabcsv$, wabrc$, wabo%
Dim csvxdelim As String

grplp = -1
csvxdelim = form1.getusersetting("csvexportdelimiter", ";")
Label12.Caption = ""
xld$ = form1.getusersetting("exceldelimiter", ",")
hda$ = ""
hdk$ = ""
fk% = -1: fa% = -1
ofn$ = form1.mydir() + "\" + trm(currgrp.text) + ".cso"
If Not nexist(ofn$) Then
  p% = FreeFile
  Open ofn$ For Input As #p%
  wabl = 2
  wabptr = -1
  While Not EOF(p%)
    Line Input #p%, ofn$
    If trm(ofn$) <> "" Then
      If InStr(LCase(ofn$), "wab:") = 1 Then
        ofn$ = Mid$(ofn$, 5)
        wabptr = 0
        If InStr(LCase(ofn$), "out:") = 1 Then
          ofn$ = strrepl(Mid$(ofn$, 5), """", "")
          ofn$ = strrepl(ofn$, ";", ",")
          wabl = 1
        End If
        If InStr(LCase(ofn$), "in:") = 1 Then
          ofn$ = Mid$(ofn$, 4)
          ofn$ = strrepl(ofn$, """", "")
          ofn$ = strrepl(ofn$, ";", ",")
          wabl = 0
        End If
        If wabl < 2 Then
          While trm(ofn$) <> ""
            wabtab(wabl, wabptr) = cut_d1(ofn$, ",")
Debug.Print wabl; ","; wabptr; " "; wabtab(wabl, wabptr)
            ofn$ = cut_d2bis(ofn$, ",")
            wabptr = wabptr + 1
          Wend
          wabl = 2
        End If
      Else
        grplp = grplp + 1
        grpl(grplp) = trm(ofn$)
      End If
    End If
  Wend
  Close #p%
End If
ofn$ = form1.myuniquedocname("", "csv")
If ofn$ = "" Then Exit Sub
ofna$ = ofn$ + ".adressen.csv"
ofnk$ = ofn$ + ".kontakte.csv"
wabcsv$ = ofn$ + ".wab.csv"
On Error Resume Next
Kill ofna$
Kill ofnk$
Kill wabcsv$
On Error GoTo 0
nrecs = grpmembers.ListCount
t0 = Date + Time
If wabptr >= 0 Then
  wabo% = FreeFile
  Open wabcsv$ For Output As #wabo%
End If
fk% = FreeFile
Open ofnk$ For Output As #fk%
fa% = FreeFile
Open ofna$ For Output As #fa%
For i% = 0 To nrecs - 1
  grpmembers.ListIndex = i%
  DoEvents
  an$ = grpmembers.List(i%)
  If InStr(an$, "(ID:") > 0 Then
    an$ = "select * from adressgruppen where id='" + trm(Mid$(an$, InStr(an$, "(ID:") + 4)) + "'"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, an$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r1.EOF Then
      If trmvalidate(r1!kid) <> "-1" Then
        c$ = "select vid as adressid,Name,Strasse,ort,tel,FAX,email,Handy,plz,lkz as land,PLZPostfach,Postfach,Postanrede,' ' as hinweise from kontakt where id='" + r1!kid + "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
        If r.EOF Then
          Print #fk, """" + "Lesefehler bei kid " + trm(r1!kid) + """"
        Else
          rc$ = ""
          If hdk$ = "" Then
            If wabptr >= 0 Then
              hdk$ = ""
              For j% = 0 To wabptr - 1
                hdk$ = hdk$ + wabtab(1, j%) + csvxdelim
              Next j
              Print #wabo%, hdk$
            End If
            hdk$ = """" + "Sortiername" + """" + xld$ + """" + "PLZORT" + """" + xld$ + """" + "POSTFACHSTRASSE" + """"
            For j% = 1 To r.Fields.Count - 1
              If hdk$ <> "" Then hdk$ = hdk$ + xld$
              hdk$ = hdk$ + """" + r.Fields(j%).name + """"
              If j% = 1 Then hdk$ = hdk$ + xld$ + """Kontakt"""
            Next j%
            hdk$ = hdk$ + xld$ + """Anrede"""
            hdk$ = hdk$ + xld$ + """Abrede"""
            If grplp >= 0 Then
              For j% = 0 To grplp
                If hdk$ <> "" Then hdk$ = hdk$ + xld$
                hdk$ = hdk$ + """" + grpl(j%) + """"
                cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='" + grpl(j%) + "' ORDER BY typ, position"
                Set r2 = New ADODB.Recordset
                r2.CursorLocation = adUseServer
                rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
                While Not r2.EOF
                  If hdk$ <> "" Then hdk$ = hdk$ + xld$
                  hdk$ = hdk$ + """" + r2!feldname + """"
                  r2.MoveNext
                Wend
              Next j%
            End If
            Print #fk%, hdk$
          End If
          tw$ = "select name from adresse where id='" + r1!adressid + "';"
          adrnam$ = ""
          Set rtw = New ADODB.Recordset
          rtw.CursorLocation = adUseServer
rrr = form1.adoopen(rtw, tw$, form1.adoc, adOpenDynamic, adLockReadOnly)
          If Not rtw.EOF Then adrnam$ = trmvalidate(rtw!name)
          tw$ = ""
          For k% = 0 To wabptr - 1: wabtab(2, k%) = "": Next k%
          For j% = 1 To r.Fields.Count - 1
            If j% = 1 Then
              plzs$ = form1.pfstrasse(trmvalidate(r!plz), trmvalidate(r!plzpostfach), trmvalidate(r!postfach), trmvalidate(r!strasse))
              plzo$ = form1.plzortpostfach(trmvalidate(r!land), trmvalidate(r!plz), trmvalidate(r!plzpostfach), trmvalidate(r!ort))
              'trmvalidate (trmvalidate(r!plz) + " " + trmvalidate(r!ort))
              If plzo$ = "" Then
                plzo$ = form1.plzoofadr(trmvalidate(r1!adressid))
              End If
              rc$ = """" + adrnam$ + """" + xld$ + """" + plzo$ + """" + xld$ + """" + plzs$ + """" + xld$ + """" + adrnam$ + """"
              wabrc$ = ""
            End If
            If rc$ <> "" Then rc$ = rc$ + xld$
            tw$ = trmvalidate(r.Fields(j%).value)
            If tw$ = "" Then
              tw$ = r.Fields(j%).name
              If tw$ = "lkz" Then tw$ = "land"
              tw$ = "select " + tw$ + " from adresse where id='" + r1!adressid + "';"
              Set rtw = New ADODB.Recordset
              rtw.CursorLocation = adUseServer
rrr = form1.adoopen(rtw, tw$, form1.adoc, adOpenDynamic, adLockReadOnly)
              tw$ = ""
              If rrr = 0 Then
                If Not rtw.EOF Then
                  tw$ = trmvalidate(rtw.Fields(0).value)
                End If
              End If
            End If
            If isnumber(tw$) Then tw$ = "'" + tw$
            rc$ = rc$ + """" + tw$ + """"
            If wabptr > 0 Then
              For k% = 0 To wabptr - 1
                If wabtab(0, k%) = r.Fields(j%).name Then
                  wabtab(2, k%) = tw$
                  Exit For
                End If
              Next k%
            End If
          Next j%
          If rc$ <> "" Then rc$ = rc$ + xld$
          rc$ = rc$ + """" + form1.meineanrede(trm(r1!kid)) + """"
          If rc$ <> "" Then rc$ = rc$ + xld$
          rc$ = rc$ + """" + form1.meineabrede(trm(r1!kid)) + """"
          For j% = 0 To grplp
            c$ = "SELECT typ,wert FROM adresstyp where vid='" + r1!adressid + "' and kid='" + r1!kid + "' and typ='" + grpl(j%) + "'"
            Set r2 = New ADODB.Recordset
            r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
            If rc$ <> "" Then rc$ = rc$ + xld$
            cmd$ = ""
            If Not r2.EOF Then cmd$ = trmvalidate(r2!wert)
            rc$ = rc$ + """" + cmd$ + """"
            cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='" + grpl(j%) + "' ORDER BY typ, position"
            Set r2 = New ADODB.Recordset
            r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
            While Not r2.EOF
              cmd$ = "SELECT felddaten as wert FROM auftritthigru where auftrittsid='" + r1!adressid + r1!kid + "' and auftrittstyp='" + grpl(j%) + "' and feldname='" + trmvalidate(r2!feldname) + "'"
              Set r3 = New ADODB.Recordset
              r3.CursorLocation = adUseServer
rrr = form1.adoopen(r3, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
              cmd$ = ""
              If Not r3.EOF Then cmd$ = trmvalidate(r3!wert)
              If rc$ <> "" Then rc$ = rc$ + xld$
              rc$ = rc$ + """" + cmd$ + """"
              r2.MoveNext
            Wend
            r2.Close
          Next j%

          Print #fk%, rc$
          If wabptr >= 0 Then
            For k% = 0 To wabptr% - 1
              Print #wabo, wabtab(2, k%); csvxdelim;
Debug.Print wabtab(0, k%); "->"; wabtab(1, k%); "="; wabtab(2, k%)
            Next k%
            Print #wabo,
          End If
        End If
      Else
        c$ = "select id as adressid,Name,Strasse,ort,tel,FAX,email,Handy,plz,land,PLZPostfach,Postfach,Postanrede,hinweise from adresse where id='" + r1!adressid + "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
        If r.EOF Then
          Print #fa, """" + "Lesefehler bei adressid " + trm(r1!adressid) + """"
        Else
          rc$ = ""
          If hda$ = "" Then
            If wabptr >= 0 Then
              hda$ = ""
              For j% = 0 To wabptr - 1
                hda$ = hda$ + wabtab(1, j%) + csvxdelim
              Next j
              Print #wabo%, hda$
            End If
            'hda$ = """" + "Sortiername" + """"
            hda$ = """" + "Sortiername" + """" + xld$ + """" + "PLZORT" + """" + xld$ + """" + "POSTFACHSTRASSE" + """"
            For j% = 1 To r.Fields.Count - 1
              If hda$ <> "" Then hda$ = hda$ + xld$
              hda$ = hda$ + """" + r.Fields(j%).name + """"
              If j% = 1 Then hda$ = hda$ + xld$ + """Kontakt"""
            Next j%
            hda$ = hda$ + xld$ + """Anrede"""
            hda$ = hda$ + xld$ + """Abrede"""
            If grplp >= 0 Then
              For j% = 0 To grplp
                If hda$ <> "" Then hda$ = hda$ + xld$
                hda$ = hda$ + """" + grpl(j%) + """"
                cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='" + grpl(j%) + "' ORDER BY typ, position"
                Set r2 = New ADODB.Recordset
                r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
                While Not r2.EOF
                  If hda$ <> "" Then hda$ = hda$ + xld$
                  hda$ = hda$ + """" + r2!feldname + """"
                  r2.MoveNext
                Wend
              Next j%
            End If
            Print #fa%, hda$
          End If
          rc$ = """" + trmvalidate(r!adressid) + """"
          c$ = trmvalidate(r!adressid)
          For j% = 1 To r.Fields.Count - 1
            If rc$ <> "" Then rc$ = rc$ + xld$
            If j = 1 Then
              plzs$ = form1.pfstrasse(trmvalidate(r!plz), trmvalidate(r!plzpostfach), trmvalidate(r!postfach), trmvalidate(r!strasse))
              plzo$ = form1.plzortpostfach(trmvalidate(r!land), trmvalidate(r!plz), trmvalidate(r!plzpostfach), trmvalidate(r!ort))
              rc$ = rc$ + """" + plzo$ + """" + xld$ + """" + plzs$ + """" + xld$
            End If
            tw$ = trmvalidate(r.Fields(j%).value)
            If isnumber(tw$) Then tw$ = "'" + tw$
            If wabptr > 0 Then
              For k% = 0 To wabptr - 1
                If wabtab(0, k%) = r.Fields(j%).name Then
                  wabtab(2, k%) = tw$
                  Exit For
                End If
              Next k%
            End If
            rc$ = rc$ + """" + tw$ + """"
            If j% = 1 Then
              rc$ = rc$ + xld$
              rc$ = rc$ + """ """
            End If
          Next j%
          If rc$ <> "" Then rc$ = rc$ + xld$
          rc$ = rc$ + """" + form1.meineanrede("-1." & trm(r1!adressid)) + """"
          If rc$ <> "" Then rc$ = rc$ + xld$
          rc$ = rc$ + """" + form1.meineabrede("-1." & trm(r1!adressid)) + """"
          For j% = 0 To grplp
            c$ = "SELECT typ,wert FROM adresstyp where vid='" + trmvalidate(r!adressid) + "' and kid='-1' and typ='" + grpl(j%) + "'"
            Set r2 = New ADODB.Recordset
            r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
            If rc$ <> "" Then rc$ = rc$ + xld$
            rc$ = rc$ + """" + trmvalidate(r2!wert) + """"
            cmd$ = "SELECT id,typ,FeldName,zeilen From auftrittsfelder where typ='" + grpl(j%) + "' ORDER BY typ, position"
            Set r2 = New ADODB.Recordset
            r2.CursorLocation = adUseServer
rrr = form1.adoopen(r2, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
            While Not r2.EOF
              cmd$ = "SELECT felddaten as wert FROM auftritthigru where auftrittsid='" + trmvalidate(r!adressid) + "' and auftrittstyp='" + grpl(j%) + "' and feldname='" + trm(r2!feldname) + "'"
              Set r3 = New ADODB.Recordset
              r3.CursorLocation = adUseServer
              Call form1.dbg2f(cmd$)
rrr = form1.adoopen(r3, cmd$, form1.adoc, adOpenDynamic, adLockReadOnly)
              cmd$ = ""
              If Not r3.EOF Then
                cmd$ = trmvalidate(r3!wert)
              End If
              If rc$ <> "" Then rc$ = rc$ + xld$
              rc$ = rc$ + """" + cmd$ + """"
              r2.MoveNext
            Wend
            r2.Close
          Next j%
          
          
          Print #fa%, rc$
          If wabptr >= 0 Then
            For k% = 0 To wabptr% - 1
              Print #wabo, wabtab(2, k%); csvxdelim;
Debug.Print wabtab(0, k%); "->"; wabtab(1, k%); "="; wabtab(2, k%)
            Next k%
            Print #wabo,
          End If
       End If
      End If
    End If
  End If
  tn = Date + Time
  Dt = tn - t0
  spr = Dt / (i% + 1)
  tot = Int(nrecs * spr * 86400)
  test = nrecs - (i% + 1)
  test = Int(test * spr * 86400)
  Label12.Caption = "gesamt: " + trm(tot) + "s, " + trm(test) + "s verbleibend"
Next i%
If fa% >= 0 Then Close #fa%
If fk% >= 0 Then Close #fk%
If wabptr >= 0 Then Close #wabo
o% = FreeFile
Open ofn$ + ".alle.csv" For Output As #o%
If Not nexist(ofna$) Then
  fa% = FreeFile: Open ofna$ For Input As #fa%
  While Not EOF(fa%)
    Line Input #fa%, c$: Print #o%, c$
  Wend
  Close #fa%
End If
If Not nexist(ofnk$) Then
  fa% = FreeFile: Open ofnk$ For Input As #fa%
  If Not nexist(ofna$) Then
    On Error Resume Next
    Line Input #fa%, c$
    rrr = Err
    On Error GoTo 0
  End If
  While Not EOF(fa%) And rrr = 0
    Line Input #fa%, c$: Print #o%, c$
  Wend
  Close #fa%
End If
Close #o%
X = Shell("explorer.exe " + DirName(ofn$), vbNormalFocus)
Label12.Caption = ""
End Sub

Private Sub Command3_Click()
break% = 1
Call moveme(1, 0)
Call wrini
Call Command6_Click
End Sub

Private Sub Command30_Click()
Dim i%, r As ADODB.Recordset, an$, j As Integer, l As String, k As Integer
Dim sid$, sida$, sidk$, p As Integer, nid$, gn$, rrr

i% = allgrps.ListIndex
If i% < 0 Then Exit Sub
gn$ = allgrps.List(i%)
templist.Clear
For i% = 0 To grpmembers.ListCount - 1
  grpmembers.ListIndex = i%
  DoEvents
  an$ = grpmembers.List(i%)
  If InStr(an$, "(ID:") > 0 Then
    an$ = "select * from adressgruppen where id='" + trm(Mid$(an$, InStr(an$, "(ID:") + 4)) + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, an$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      Call shwAdrDetail.savecheck
      Call shwAdrDetail.refreshadrdetail(trmvalidate(r!adressid), trmvalidate(r!kid))
      Call shwAdrDetail.SetFocus
      DoEvents
      For j = 0 To shwAdrDetail.List1b.ListCount - 1
        If InStr(shwAdrDetail.List1b.List(j), Combo6.text) = 1 Then
          l = cut_d2bis(shwAdrDetail.List1b.List(j), ":")
          For k = 0 To templist.ListCount - 1
            If l = templist.List(k) Then Exit For
          Next k
          If k >= templist.ListCount Then templist.AddItem l
        End If
      Next j
    End If
  End If
Next i%
For i% = 0 To templist.ListCount - 1
  sid$ = templist.List(i%)
  sida$ = sid$: sidk$ = ""
  p = InStr(sida$, "{")
  If p > 0 Then
    sidk$ = trm(Left(sid$, p - 1))
    sida$ = trm(Mid(sid$, p + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If sidk$ = "" Then
    sidk$ = "-1"
  Else
    sidk$ = form1.getkontaktidbyname(sida$, sidk$)
  End If
  nid$ = form1.newid("adressgruppen", "id", 20)
  form1.sqlqry ("insert into adressgruppen (id,adressid,grpid,kid) values('" + nid$ + "','" + sida$ + "','" + gn$ + "','" + sidk$ + "')")
Next i%
Call allgrps_Click

End Sub

Private Sub Command31_Click()
Dim i%, r As ADODB.Recordset, an$, j As Integer, l As String, k As Integer
Dim sid$, sida$, sidk$, p As Integer, nid$, gn$, rrr

i% = allgrps.ListIndex
If i% < 0 Then Exit Sub
gn$ = allgrps.List(i%)
templist.Clear
tempdel.Clear
For i% = 0 To grpmembers.ListCount - 1
  grpmembers.ListIndex = i%
  DoEvents
  an$ = grpmembers.List(i%)
  If InStr(an$, "(ID:") > 0 Then
    an$ = "select * from adressgruppen where id='" + trm(Mid$(an$, InStr(an$, "(ID:") + 4)) + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseServer
rrr = form1.adoopen(r, an$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r.EOF Then
      Call shwAdrDetail.savecheck
      Call shwAdrDetail.refreshadrdetail(r!adressid, r!kid)
      Call shwAdrDetail.SetFocus
      DoEvents
      For j = 0 To shwAdrDetail.List1b.ListCount - 1
        If InStr(shwAdrDetail.List1b.List(j), Combo6.text) = 1 Then
          tempdel.AddItem grpmembers.List(i%)
          l = cut_d2bis(shwAdrDetail.List1b.List(j), ":")
          For k = 0 To templist.ListCount - 1
            If l = templist.List(k) Then Exit For
          Next k
          If k >= templist.ListCount Then
            templist.AddItem l
          End If
        End If
      Next j
    End If
  End If
Next i%
For i% = 0 To templist.ListCount - 1
  sid$ = templist.List(i%)
  sida$ = sid$: sidk$ = ""
  p = InStr(sida$, "{")
  If p > 0 Then
    sidk$ = trm(Left(sid$, p - 1))
    sida$ = trm(Mid(sid$, p + 1)): sida$ = Left(sida$, Len(sida$) - 1)
  End If
  If sidk$ = "" Then
    sidk$ = "-1"
  Else
    sidk$ = form1.getkontaktidbyname(sida$, sidk$)
  End If
  nid$ = form1.newid("adressgruppen", "id", 20)
  form1.sqlqry ("insert into adressgruppen (id,adressid,grpid,kid) values('" + nid$ + "','" + sida$ + "','" + gn$ + "','" + sidk$ + "')")
Next i%
For i% = 0 To tempdel.ListCount - 1
  For j = grpmembers.ListCount - 1 To 0 Step -1
    If tempdel.List(i) = grpmembers.List(j) Then
      grpmembers.ListIndex = j
      DoEvents
      Call Command11_Click
      Exit For
    End If
  Next j
Next i%
tempdel.Clear
Call allgrps_Click

End Sub

Private Sub Command32_Click()
Dim fn As String, X

If trm(currgrp.text) = "" Then
  Command32.Enabled = False
  Exit Sub
End If

fn = form1.mydir() + "\" + trm(currgrp.text) + ".cso"
X = Shell("notepad.exe " + fn, vbNormalFocus)
End Sub

Private Sub Command33_Click()
Dim mergeid As String, i%, gn$, cmd$
Dim rtmp As ADODB.Recordset, l$, rrr, ai As String, ki As String

i% = allgrps.ListIndex
If i% < 0 Then
  MsgBox ("Bitte wählen Sie erst eine Gruppe aus," + vbCrLf + "in die importiert werden soll")
  Exit Sub
End If
gn$ = allgrps.List(i%)
mergeid = InputBox(transe("Welche Adressgruppe soll importiert werden in") + " " + gn$, transe("Adressgruppe importieren"))
MousePointer = 11: DoEvents
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT AdressID,kid FROM adressgruppen where grpid='" + mergeid + "'", form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
List1.Clear
List2.Clear: i% = 0
i% = 0
While Not rtmp.EOF
  ki = trm(rtmp!kid)
  ai = rtmp!adressid
  If "" = form1.get1erg("select id as wert from adressgruppen where AdressID='" + ai + "' and kid='" + ki + "' and grpid='" + gn$ + "'") Then
    If ki <> "" And ki <> "-1" Then
      List2.AddItem form1.get_kontaktname_by_id(ki) + Space$(80) + "(ID:" + ai
      List2.ListIndex = List1.ListCount - 1
    Else
      List1.AddItem ai + Space$(80) + "(ID:" + ai
      List1.ListIndex = List2.ListCount - 1
    End If
    DoEvents
    cmd$ = "insert into adressgruppen (id,AdressID,grpid,kid) values("
    cmd$ = cmd$ + "'" + form1.newid("adressgruppen", "id", 40) + "',"
    cmd$ = cmd$ + "'" + ai + "',"
    cmd$ = cmd$ + "'" + gn$ + "',"
    cmd$ = cmd$ + "'" + ki + "')"
    Call form1.sqlqry(cmd$)
    i% = i% + 1
  End If
  rtmp.MoveNext
Wend
List1.Clear
List2.Clear
MousePointer = 0
Call allgrps_Click
MsgBox ("Es wurden " + trm(i%) + " Adressen übernommen")
End Sub

Private Sub Command34_Click()
Dim i%, n%, l$, na$, e$, cb$, c$, nalist$, o%
Dim rtmp As ADODB.Recordset, rrr, id$, gn$, nid$, an$, ank$

gn$ = currgrp.text
If gn$ = "" Then Exit Sub
cb$ = Clipboard.GetText
If cb$ = "" Then Exit Sub

MousePointer = 11
o% = FreeFile
Open "aptmpadr.txt" For Output As #o%
Print #o%, cb$
Close #o%
o% = FreeFile
Open "aptmpadr.txt" For Input As #o%
n% = linesof(cb$)
pb1.Max = n% + 1
pb1.Top = Command30.Top
pb1.Visible = True
na$ = ""
While Not EOF(o%)
  Line Input #o%, l$
  If i% < n% Then i% = i% + 1
  pb1.value = i%
  DoEvents
  If na$ = "" And l$ <> "" Then
    na$ = l$
    e$ = strrepl(na$, ",", " "): e$ = strrepl(e$, "  ", " ")
    While e$ <> ""
      l$ = word1(e$): e$ = word2bis(e$)
      If l$ <> "" Then
        If nalist$ <> "" Then nalist$ = nalist$ + " and "
        nalist$ = nalist$ + "name like '%" + l$ + "%'"
      End If
    Wend
  Else
    If l$ = "" Then
      na$ = ""
      nalist$ = ""
    Else
      e$ = emailonly(trm(l$))
      If e$ <> "" Then
        Debug.Print na$, e$
        an$ = "": ank$ = "-1"
        c$ = "select id from adresse where email='" + e$ + "'"
        Set rtmp = New ADODB.Recordset
        rtmp.CursorLocation = adUseServer
        rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
        If rrr = 0 Then
          If Not rtmp.EOF Then
            an$ = trm(rtmp!id)
          End If
        End If
        If an$ = "" Then
          c$ = "select id,vid from kontakt where email='" + e$ + "'"
          Set rtmp = New ADODB.Recordset
          rtmp.CursorLocation = adUseServer
          rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          If rrr = 0 Then
            If Not rtmp.EOF Then
              an$ = trm(rtmp!vid)
              ank$ = trm(rtmp!id)
            End If
          End If
        End If
        If an$ = "" Then
          c$ = "insert into adresse (id,name,email) values('" & na$ & "','" & na$ & "','" & e$ & "')"
          Call form1.sqlqry(c$)
          c$ = "insert into adresstyp (id,vid,typ,wert,kid) values('" + form1.newid("adresstyp", "id", 20) + "','" + na$ + "','Person',NULL,'-1')"
          Call form1.sqlqry(c$)
          an$ = na$
        End If
        If an$ <> "" And gn$ <> "" Then
          Call addmailing(an$, ank$, gn$)
          c$ = "SELECT id FROM adressgruppen where AdressID='" + an$ + "' and kid='" + ank$ + "' and grpid='" + gn$ + "'"
          Set rtmp = New ADODB.Recordset
          rtmp.CursorLocation = adUseServer
          rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
          If rrr = 0 Then
          If rtmp.EOF Then
            nid$ = form1.newid("adressgruppen", "id", 20)
            form1.sqlqry ("insert into adressgruppen (id,adressid,grpid,kid) values('" + nid$ + "','" + an$ + "','" + gn$ + "','" + ank$ + "')")
            grpmembers.AddItem an$ + Space$(40) + "(ID:" + nid$
            grpmembers.ListIndex = grpmembers.ListCount - 1
            Label7.Caption = trm(grpmembers.ListCount)
            DoEvents
          End If
          End If
        End If
      End If
    End If
  End If
Wend
Close #o%
On Error Resume Next
Kill "aptmpadr.txt"
On Error GoTo 0
pb1.Visible = False
MousePointer = 0
End Sub

Private Sub addmailing(aid$, koid$, mgid$)
Dim rtmp As ADODB.Recordset, rrr, id$, nid$, c$, w$, kid$

kid$ = koid$
If kid$ = "-1" Then kid$ = ""
c$ = "select id,FeldDaten from auftritthigru where auftrittsid='" + aid$ + kid$ + "' and auftrittstyp='Person' and FeldName='mailings'"
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, c$, form1.adoc, adOpenDynamic, adLockReadOnly, "", "")
If rrr = 0 Then
  If rtmp.EOF Then
    nid$ = form1.newid("auftritthigru", "id", 26)
    c$ = "insert into auftritthigru (id,auftrittsid,auftrittstyp,FeldName,FeldDaten) values('"
    c$ = c$ + nid$ + "','" + aid$ + kid$ + "','Person','mailings','," + mgid$ + ",')"
    Call form1.sqlqry(c$)
    If Not form1.isfieldmissing("auftritthigru", "opt_kid") And kid$ <> "" Then
      c$ = "update auftritthigru set opt_kid='" + koid$ + "' where id='" + nid$ + "'"
      Call form1.sqlqry(c$)
    End If
  Else
    w$ = trm(rtmp!felddaten)
    nid$ = trm(rtmp!id)
    If InStr(w$, "," + mgid$ + ",") = 0 Then
      If Right$(w$, 1) <> "," Then w$ = w$ + ","
      w$ = w$ + mgid$ + ","
      c$ = "update auftritthigru set FeldDaten='" + w$ + "' where id='" + nid$ + "'"
      Call form1.sqlqry(c$)
    End If
  End If
End If

End Sub

Private Sub Command35_Click()
Call form1.handbuchcall("06.1-SucheUndSelektion.htm")
End Sub

Private Sub Command36_Click()
Dim i%, anid$, p%, an$, r As ADODB.Recordset, r1 As ADODB.Recordset, r2 As ADODB.Recordset, r3 As ADODB.Recordset
Dim fld As Field, c$, j%, rc$, hdk$, hda$, tw$, rtw As ADODB.Recordset, rrr, o%
Dim fk%, fa%, ofn$, X, xld$, adrnam$, ofna$, ww$
Dim cmd$, t0 As Double, tn As Double, Dt As Double, spr As Double
Dim test As Double, nrecs As Long, tot As Double, anr$

nrecs = grpmembers.ListCount
form1.hordexlock = True
t0 = Date + Time
For i% = 0 To nrecs - 1
  grpmembers.ListIndex = i%
  DoEvents
  an$ = grpmembers.List(i%)
  adrnam$ = ""
  If InStr(an$, "(ID:") > 0 Then
    an$ = "select * from adressgruppen where id='" + trm(Mid$(an$, InStr(an$, "(ID:") + 4)) + "'"
    Set r1 = New ADODB.Recordset
    r1.CursorLocation = adUseServer
rrr = form1.adoopen(r1, an$, form1.adoc, adOpenDynamic, adLockReadOnly)
    If Not r1.EOF Then
      If trmvalidate(r1!kid) <> "-1" Then
        c$ = "select vid as adressid from kontakt where id='" + r1!kid + "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
        If Not r.EOF Then
          adrnam$ = trmvalidate(r!adressid)
        End If
      Else
        c$ = "select id as adressid from adresse where id='" + r1!adressid + "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseServer
rrr = form1.adoopen(r, c$, form1.adoc, adOpenDynamic, adLockReadOnly)
        If Not r.EOF Then
          adrnam$ = trmvalidate(r!adressid)
       End If
      End If
    End If
  End If
  If adrnam$ <> "" Then
    ww$ = form1.cloudmanager + form1.cloudstaff: c$ = "x"
    While ww$ <> ""
      c$ = cut_d1(ww$, "|"): ww$ = cut_d2bis(ww$, "|")
      If c$ <> "" Then
        anr$ = adrnam$ + "|" + c$
        For j = 0 To form1.hordex.ListCount - 1
          If form1.hordex.List(j) = anr$ Then Exit For
        Next j
        If j >= form1.hordex.ListCount Or form1.hordex.ListCount = 0 Then
          Call form1.add2hordex(anr$)
        End If
        DoEvents
      End If
    Wend
  End If
  tn = Date + Time
  Dt = tn - t0
  spr = Dt / (i% + 1)
  tot = Int(nrecs * spr * 86400)
  test = nrecs - (i% + 1)
  test = Int(test * spr * 86400)
  Label12.Caption = "gesamt: " + trm(tot) + "s, " + trm(test) + "s verbleibend"
Next i%
form1.hordexlock = False
Label12.Caption = ""

End Sub

Private Sub Command37_Click()
Call shwAdrDetail.Command11_Click
End Sub

Private Sub Command4_Click()
break% = 1
While List3(0).ListCount > 0
  List3(0).ListIndex = 0
  Call moveme(0, 1)
Wend
Call wrini
Call Command6_Click
End Sub

Private Sub Command42_Click()
Dim dst$, d As Integer, rrr, plza$
Dim s As ADODB.Recordset, c$
Dim asinbreit As Double, acoslang As Double, alang As Double, radf As String, acosbreit As Double

If Not form1.geodbok Then
  MsgBox "Die Geodatenbank ist nicht konfiguriert"
  Exit Sub
End If
umplz.Caption = ""
plzlimit = ""
plza$ = word1(shwAdrDetail.datf(3).text)
If shwAdrDetail.datf(13).text <> "" Then plza$ = shwAdrDetail.datf(13).text
If Not isnumber(plza$) Then
  plza$ = trm(InputBox(transe("Suche im Umkreis der Postleitzahl:"), transe("Umkreissuche"), ""))
  If plza$ = "" Then
    Exit Sub
  End If
End If
dst$ = form1.mylastFormVar(Me.name, "umkrsuch", "30")
dst$ = trm(InputBox(transe("Suche im Umkreis (km):"), transe("Umkreissuche"), dst$))
On Error Resume Next
d = Val(dst$)
rrr = Err
On Error GoTo 0
If rrr <> 0 Or d = 0 Then Exit Sub
d = Abs(d)
Call form1.setmylastFormVar(Me.name, "umkrsuch", dst$)

c$ = "SELECT zc_id, zc_location_name, zc_lat, zc_lon from zip_coordinates WHERE zc_zip = '" + plza$ + "'"
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
  If plzlimit = "" Then plzlimit = "|"
  plzlimit = plzlimit + trm(s!zc_zip) + "|"
  s.MoveNext
Wend
umplz.Caption = dst$ + " km um " + plza$
DoEvents
MousePointer = 0
DoEvents
Call Command6_Click
End Sub

Private Sub Command5_Click()
break% = 1
While List3(1).ListCount > 0
  List3(1).ListIndex = 0
  Call moveme(1, 0)
Wend
Call wrini
Call Command6_Click
End Sub

Public Sub Command6_Click()

break% = 1
Timer1.Enabled = False
DoEvents
Timer1.Interval = suchvz
Timer1.Enabled = True
snotb4 = now()

End Sub

Private Sub Command7_Click()
Dim s$, o%, fn$

If suchstr$ <> "" Then
  s$ = strrepl(suchstr$, "adresse.ID,adresse.name", "*")
  fn$ = form1.myuniquedocname("", "sqs")
  If fn$ <> "" Then
    o% = FreeFile
    Open fn$ For Output As #o%
    Print #o%, s$
    Close #o%
    Call Command25_Click
  End If
End If

End Sub

Private Sub Command8_Click()
Dim rtmp As ADODB.Recordset, gn$, i%, rrr

gn$ = trm(currgrp.text)
If gn$ <> "" Then
  form1.sqlqry ("insert into adressgruppenindex (id) values('" + gn$ + "')")
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppenindex", form1.adoc, adOpenDynamic, adLockReadOnly)
  allgrps.Clear
  While Not rtmp.EOF
    allgrps.AddItem rtmp!id
    rtmp.MoveNext
  Wend
  currgrp.text = ""
  For i% = 0 To allgrps.ListCount - 1
    If allgrps.List(i%) = gn$ Then
      allgrps.ListIndex = i%
      Exit For
    End If
  Next i%
End If

End Sub

Private Sub Command9_Click()
Dim s$, fn$, o%

If suchstr$ <> "" Then
  s$ = strrepl(ksuchstr$, "name,id,vid", "*")
  fn$ = form1.myuniquedocname("", "sqs")
  If fn$ <> "" Then
    o% = FreeFile
    Open fn$ For Output As #o%
    Print #o%, s$
    Close #o%
    Call Command25_Click
  End If
End If

End Sub

Private Sub Form_Load()
Dim fgdbn$

axsResizer1.SaveControlPositions
'Set wrkJet = CreateWorkspace("", "Admin", "", dbUseJet)
Call form1.dbg2f("adrselect:load")
kontselid$ = ""
plzlimit = ""
fl_rl3% = 0
Me.Top = form1.mylasttop(Me.name)
Me.Left = form1.mylastleft(Me.name)
If Me.Top = 20 And Me.Left = 20 Then
  Me.Top = form1.Top + 200
  Me.Left = form1.Left + form1.Height
End If
Call form1.formpos(Me)
sel_break = 0
msec = 1# / (24# * 3600# * 1000#)
suchvz = form1.getsuchvz()
Text1.text = ""
List1.Clear
List2.Clear
If form1.cloud Then Command36.Enabled = True
Text3.text = form1.mylastFormVar(Me.name, "suche_nur", "20")
adrselect.Caption = transe("Adresse wählen")
List3(2).ToolTipText = transe("weitere Adressgruppenkriterien")
Command18.Caption = transe("CSV-Liste")
Check2.ToolTipText = transe("Suchkriterien auf Kontakte (oder Adressen) anwenden")
Command17.Caption = transe("CSV-Liste")
Command19.Caption = transe("CSV-Alle")
Command6.Caption = transe("&LOS")
Combo5.text = transe(form1.mylastFormVar(Me.name, "SortAdresse", "PLZ"))
Combo4.text = transe(form1.mylastFormVar(Me.name, "SortKontakt", "PLZ"))
Combo3.Clear
Combo3.AddItem transe("gleich")
Combo3.AddItem transe("enthält")
Combo3.AddItem transe("beginnt mit")
Combo3.text = transe("gleich")
Combo2.Clear
Combo2.AddItem transe("gleich")
Combo2.AddItem transe("enthält")
Combo2.AddItem transe("beginnt mit")
Combo2.AddItem transe("leer")
Combo2.AddItem transe("nicht leer")
Combo2.text = transe("gleich")
Command16.Caption = transe("Serienbrief")
Command15.Caption = transe("alle dazu")
Command14.ToolTipText = transe("Den markierten Satz löschen")
Command21.ToolTipText = transe("Speichern")
Command30.ToolTipText = transe("Folgt der gewählten Beziehung und fügt die Adressen hinzu")
Command31.ToolTipText = transe("löscht die Adresse und fügt die Beziehung hinzu")
Command42.ToolTipText = transe("Adressen im Umkreis um eine Postleitzahl suchen")
Command22.ToolTipText = transe("Als neuen Filter speichern")
Command11.ToolTipText = transe("Den markierten Satz löschen")
Command25.ToolTipText = transe("gespeicherte Selektionen")
Command9.Caption = transe("Kontakte merken")
Command7.Caption = transe("Adressen merken")
Text3.ToolTipText = transe("finde nur so viele Einträge, 0=alle")
Command13.Caption = transe("&Kontakt(e) dazu")
Command23.Caption = transe("UND-dazu")
Command24.Caption = transe("ODER-dazu")
Command12.Caption = transe("alle dazu")
Command10.Caption = transe("A&dresse(n) dazu")
Command8.Caption = transe("neu")
List3(1).ToolTipText = transe("ignoriere diese Adressgruppen")
List3(0).ToolTipText = ("finde nur diese Adressgruppen (leer=suche alles)")
Combo10.Clear
Combo10.AddItem transe("schliesse aus:")
Combo10.AddItem transe("muss auch sein:")
Combo10.text = transe("schliesse aus:")
Command20.ToolTipText = transe("schliesse aus:") + "/" + transe("muss auch sein:")
Label6.Caption = transe("Sortierung")
Label9.Caption = transe("Kriterien f. Kontakte")
Label8.Caption = transe("mit Kontakten")
Label5.Caption = transe("Sortierung")
Label2.Caption = transe("suche nicht:")
Label1.Caption = transe("suche nur:")
Label11.Caption = transe("UND statt ODER:")
fgdbn$ = LCase(form1.getdbname())
If InStr(LCase(App.EXEName), "apadmin") = 1 Or _
         fgdbn$ = "aaimport" Or _
         fgdbn$ = "bundestag" Or _
         fgdbn$ = "weinstadt" Or _
         fgdbn$ = "wgbm" Or _
         fgdbn$ = "haas" Or _
         fgdbn$ = "klangkultur" Or _
         fgdbn$ = "example" Or fgdbn$ = "agencyprof" Or _
         fgdbn$ = "sks" Or _
         fgdbn$ = "frb" Or _
         fgdbn$ = "frb.mdb" Or _
         fgdbn$ = "krautheim" Or _
         fgdbn$ = "diapason" Or _
         fgdbn$ = "demodaten" Or _
         fgdbn$ = "muenchenmusik" Or _
         fgdbn$ = "musikforum" Or _
         fgdbn$ = "maierartists" Or _
         fgdbn$ = "maierartists.mdb" Or _
         fgdbn$ = "artino" Or _
         fgdbn$ = "kktest" Or _
         fgdbn$ = "hampl" Or _
         fgdbn$ = "schoenherr" Or _
         fgdbn$ = "orfeo" Or _
         LCase(Left(form1.getdbname(), 9)) = "bakjk_ap1" Or _
         form1.getdbname() = "kk" Then
  Command29.Visible = True
  Command32.Visible = True
Else
  Command29.Visible = False
  Command32.Visible = False
End If
Show

Call rlist3
Call rcombo1
Call rlist4
Call rlist5
Call initstring("")
sel_vld = -1

grpmembers.Clear

Call rallgroups
Timer1.Enabled = False
DoEvents
Timer1.Interval = suchvz
Timer1.Enabled = True
End Sub

Sub rallgroups()
Dim rtmp As ADODB.Recordset, rrr
Dim d2infile As String, d2insub As String

d2infile = "adrselect": d2insub = "rallgroups"

Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppenindex", form1.adoc, adOpenDynamic, adLockReadOnly, d2infile, d2insub)
If rrr <> 0 Then Exit Sub

allgrps.Clear
While Not rtmp.EOF
  allgrps.AddItem rtmp!id
  rtmp.MoveNext
Wend
currgrp.text = ""
Command32.Enabled = False

End Sub
Public Sub initstring(s$)
Text1.text = s$
End Sub

Private Sub Form_Resize()
axsResizer1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
On Error GoTo exuld
Call form1.setmylasttop(Me.name, Me.Top)
Call form1.setmylastleft(Me.name, Me.Left)
Call form1.setmylastFormVar(Me.name, "SortAdresse", transo(Combo5.text))
Call form1.setmylastFormVar(Me.name, "SortKontakt", transo(Combo4.text))

exuld:
On Error GoTo 0
End Sub


Private Sub grpmembers_DblClick()
Dim i As Integer, id As String, j As Integer
Dim rtmp As ADODB.Recordset, rrr

i = grpmembers.ListIndex
If i < 0 Then Exit Sub

j = InStr(grpmembers.List(i), "(ID:")
If j > 0 Then
  id = Mid(grpmembers.List(i), j + 4)
  Set rtmp = New ADODB.Recordset
  rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT * FROM adressgruppen where id='" + id + "';", form1.adoc, adOpenDynamic, adLockReadOnly)
  If Not rtmp.EOF Then
    Call shwAdrDetail.refreshadrdetail(rtmp!adressid, rtmp!kid)
    On Error Resume Next
    Call shwAdrDetail.SetFocus
    On Error GoTo 0
  End If
End If
End Sub

Private Sub higrusuch_Click()
Load higruselect
End Sub

Private Sub List1_DblClick()
Dim sid$, i As Integer, c As String

sid$ = ""
For i = 0 To List1.ListCount - 1
  If List1.Selected(i) Then
    c = List1.List(i)
    If InStr(c, "(") > 1 Then c = Left$(c, InStr(c, "(") - 1)
    If sid$ <> "" Then sid$ = sid$ + vbCrLf
    sid$ = sid$ + c
  End If
Next i

'sid$ = List1.List(List1.ListIndex)
'If InStr(sid$, "(") > 1 Then sid$ = Left$(sid$, InStr(sid$, "(") - 1)
selct$ = sid$
kontsel$ = ""
If sel_vld = -1 Then
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(sid$, "")
  Call shwAdrDetail.SetFocus
Else
  sel_vld = 1
End If

End Sub

Private Sub List2_DblClick()
Dim sid$, cid$
sid$ = List2.List(List2.ListIndex)
cid$ = ""
kontsel$ = ""
If InStr(sid$, "(VID:") > 1 Then
  kontselid$ = trm(Mid$(sid$, InStr(sid$, " ID:") + 4))
  cid$ = trm(Left$(sid$, InStr(sid$, "(VID:") - 1))
  kontsel$ = cid$
  If InStr(kontsel$, "(") > 0 Then kontsel$ = trm(Left$(kontsel$, InStr(kontsel$, "(") - 1))
  sid$ = Mid$(sid$, InStr(sid$, "(VID:") + 5)
End If
If InStr(sid$, ")") > 1 Then sid$ = Left$(sid$, InStr(sid$, ")") - 1)
selct$ = sid$
If sel_vld = -1 Then
  Load shwAdrDetail
  Call shwAdrDetail.savecheck
  Call shwAdrDetail.refreshadrdetail(sid$, cid$)
  Call shwAdrDetail.SetFocus
Else
  sel_vld = 1
End If

End Sub

Private Sub List3_DblClick(Index As Integer)
break% = 1
If Index = 2 Then
  Call moveme(2, 1)
  Exit Sub
End If
If Index = 0 Then
  Call moveme(0, 1)
Else
  Call moveme(1, 0)
End If
Call wrini
Call Command6_Click

End Sub

Private Sub List5_Click()
Dim i%, fn$, o%


For i% = 0 To List4.ListCount - 1
  List4.Selected(i%) = False
Next i%
Text4.text = ""
i% = List5.ListIndex
If i% < 0 Then Exit Sub

fn$ = form1.mydir() + "\" + List5.List(i%) + ".flt"
o% = FreeFile
Open fn$ For Input As #o%
While Not EOF(o%)
  Line Input #o%, fn$
  If fn$ = "<<EOL>>" Then
    Line Input #o%, fn$: Combo2.text = fn$
    Line Input #o%, fn: Text4.text = fn$
  Else
    For i% = 0 To List4.ListCount - 1
      If fn$ = List4.List(i%) Then
        List4.Selected(i%) = True
        Exit For
      End If
    Next i%
  End If
Wend
Close #o%
End Sub

Private Sub List5_dblClick()
DoEvents
Call Command6_Click
End Sub

Private Sub Text1_Change()
break% = 1
Timer1.Enabled = False
Timer1.Interval = suchvz
Timer1.Enabled = True

'snotb4 = Now() + suchvz * msec
End Sub
Sub wrini()
Dim rtmp As ADODB.Recordset, inifile As String, o%, i%, rrr

inifile = form1.mydatadir() + "\" + Me.name + ".ini"
o% = FreeFile
On Error Resume Next
Open inifile For Output As #o%
rrr = Err
On Error GoTo 0
If rrr <> 0 Then Exit Sub
For i% = 0 To List3(1).ListCount - 1
  Print #o%, List3(1).List(i%)
Next i%
Close #o%

End Sub

Public Function sel_getselected()
Dim sid$, i As Integer, sida As String, c As String

sid$ = ""
For i = 0 To List1.ListCount - 1
  If List1.Selected(i) Then
    c = List1.List(i)
    If InStr(c, "(") > 1 Then c = Left$(c, InStr(c, "(") - 1)
    If sid$ <> "" Then sid$ = sid$ + vbCrLf
    sid$ = sid$ + c
  End If
Next i

If sid$ <> "" And selct$ = "" Then
  sel_getselected = sid$
Else
  sel_getselected = selct$
End If
form1.dbg2f (trm("sel_getselected='" + sel_getselected + "'"))
End Function
Public Function sel_valid()

sel_valid = sel_vld

End Function
Public Function sel_brk()

sel_brk = sel_break

End Function
Public Sub sel_init(such0$, typl$)
Dim l$, rtmp As ADODB.Recordset, i%, rrr

kontselid$ = ""
Set rtmp = New ADODB.Recordset
rtmp.CursorLocation = adUseServer
rrr = form1.adoopen(rtmp, "SELECT id FROM adresstypen", form1.adoc, adOpenDynamic, adLockReadOnly)

List3(1).Clear
List3(0).Clear
While Not rtmp.EOF
  For i% = 0 To List3(0).ListCount - 1
    If List3(0).List(i%) = transe(rtmp!id) Then i% = List3(1).ListCount + 100
  Next i%
  If i% < List3(0).ListCount + 100 Then List3(1).AddItem transe(rtmp!id)
  rtmp.MoveNext
Wend

List3(0).Clear
While InStr(typl$, "|") > 0
  l$ = Left$(typl$, InStr(typl$, "|") - 1)
  List3(0).AddItem transe(l$)
  typl$ = Mid$(typl$, InStr(typl$, "|") + 1)
Wend
If Right$(typl$, 1) = "*" Then
  typl$ = Left$(typl$, Len(typl$) - 1)
  For i% = 0 To List3(1).ListCount - 1
    If InStr(LCase(List3(1).List(i%)), typl$) = 1 Then
      List3(0).AddItem List3(1).List(i%)
    End If
  Next i%
Else
  If typl$ <> "" Then List3(0).AddItem transe(typl$)
End If
Text1.text = such0$

sel_vld = 0
selct$ = ""
sel_break = 0

End Sub

Private Sub Text1_LostFocus()
Timer1.Enabled = False

End Sub

Private Sub Text3_Change()
Dim ts%, rrr

On Error Resume Next
ts% = Val(Text3.text)
rrr = Err
On Error GoTo 0
If rrr <> 0 Then
  Text3.text = "50"
  Exit Sub
End If
If ts% < 5 Then ts% = 5
If ts% > 9999 Then ts% = 9999
If trm(ts%) <> Text3.text Then Text3.text = trm(ts%)
Call Text1_Change
Call form1.setmylastFormVar(Me.name, "suche_nur", trm(Text3.text))

End Sub

Private Sub Timer1_Timer()
Dim s$

'If Now() < snotb4 Then Exit Sub
Call form1.dbg2f("adrselect Timer1 start")
Timer1.Enabled = False
s$ = Text1.text
break% = 0
Call rlist1(s$)
Call form1.dbg2f("adrselect Timer1 exit")
End Sub

Public Function get_kontsel() As String

get_kontsel = kontsel$

End Function

Public Function get_kontselid() As String

get_kontselid = kontselid$

End Function

Private Sub rcombo1()
Dim tr, rrr, ffn$

Combo1.Clear
tr = form1.vorlagenverzeichnis() + "\serienbrief_*.rtf"
tr = Dir(tr)
rrr = Err
On Error GoTo 0
While tr <> "" And rrr = 0
  ffn$ = basename(Mid$(tr, InStr(tr, "_") + 1), ".rtf")
  Combo1.AddItem ffn$
  tr = Dir
Wend

End Sub
Sub rlist4()
Dim i%

List4.Clear
For i% = 0 To form1.sqla.TableDefs("adresse").Fields.Count - 1
  List4.AddItem transe(form1.sqla.TableDefs("adresse").Fields(i%).name)
  Combo4.AddItem transe(form1.sqla.TableDefs("adresse").Fields(i%).name)
  Combo5.AddItem transe(form1.sqla.TableDefs("adresse").Fields(i%).name)
Next i%
End Sub

Sub rlist5()
Dim tr

List5.Clear
tr = Dir(form1.mydatadir() + "\*.flt")
While tr <> ""
  List5.AddItem basename(trm(tr), ".flt")
  tr = Dir
Wend
End Sub
Function wertsubst(w$, typ$) As String
Dim lw$, l$

wertsubst = "(wert='" + w$ + "' and typ='" + typ$ + "')"
If InStr(w$, "|") > 0 Then
  lw$ = w$
  l$ = "(wert='" + cut_d1(lw$, "|") + "' and typ='" + typ$ + "')": lw$ = cut_d2bis(lw$, "|")
  While lw$ <> ""
    l$ = l$ + " or (wert='" + cut_d1(lw$, "|") + "'  and typ='" + typ$ + "')"
    lw$ = cut_d2bis(lw$, "|")
  Wend
  wertsubst = l$
End If

End Function
Function wertsubst2(w$, typ$) As String
Dim lw$, l$

wertsubst2 = "(instr(lcase(wert),'" + LCase(w$) + "')>0 and typ='" + typ$ + "')"
If InStr(w$, "|") > 0 Then
  lw$ = w$
  l$ = "(instr(lcase(wert),'" + LCase(cut_d1(lw$, "|")) + "')>0 and typ='" + typ$ + "')": lw$ = cut_d2bis(lw$, "|")
  While lw$ <> ""
    l$ = l$ + " or (instr(lcase(wert),'" + LCase(cut_d1(lw$, "|")) + "')>0 and typ='" + typ$ + "')"
    lw$ = cut_d2bis(lw$, "|")
  Wend
  wertsubst2 = l$
End If

End Function

Function wertsubst3(w$, typ$) As String
Dim lw$, l$

wertsubst3 = "(instr(lcase(wert),'" + LCase(w$) + "')=1  and typ='" + typ$ + "') "
If InStr(w$, "|") > 0 Then
  lw$ = w$
  l$ = "(instr(lcase(wert),'" + LCase(cut_d1(lw$, "|")) + "')=1  and typ='" + typ$ + "') ": lw$ = cut_d2bis(lw$, "|")
  While lw$ <> ""
    l$ = l$ + " or (instr(lcase(wert),'" + LCase(cut_d1(lw$, "|")) + "')=1  and typ='" + typ$ + "') "
    lw$ = cut_d2bis(lw$, "|")
  Wend
  wertsubst3 = l$
End If

End Function


